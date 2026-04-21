public class MainWindow : Gtk.ApplicationWindow {
    private Gtk.Entry folder_entry;
    private Gtk.Button convert_button;
    private Gtk.Button cancel_button;
    private Gtk.Button browse_button;
    private Gtk.ProgressBar progress_bar;
    private Gtk.CheckButton recursive_check;
    private Gtk.TextView log_view;
    private Gtk.TextBuffer log_buffer;

    private string soffice_path;
    private string? custom_output_folder = null;
    private int retry_count = 2;
    private int total_files = 0;
    private int processed_files = 0;
    private Cancellable? cancellable = null;
    private string log_file_path;
    private string paper_size = "Letter";
    private string profile_root;
    private string template_profile_dir;
    private bool template_ready = false;
    private string settings_path;

    public MainWindow (Gtk.Application app) {
        Object (
            application: app,
            title: "DOC to PDF Converter",
            default_width: 700,
            default_height: 500
        );
    }

    construct {
        soffice_path = detect_libreoffice ();

        var log_dir = Path.build_filename (Environment.get_current_dir (), "logs");
        DirUtils.create_with_parents (log_dir, 0755);
        var now = new DateTime.now_local ();
        log_file_path = Path.build_filename (log_dir, "log_%s.txt".printf (now.format ("%Y%m%d_%H%M%S")));

        profile_root = Path.build_filename (Environment.get_tmp_dir (), "LibreDocToPdf");
        template_profile_dir = Path.build_filename (profile_root, "_template");
        DirUtils.create_with_parents (profile_root, 0755);

        var config_dir = Path.build_filename (Environment.get_user_config_dir (), "LibreDocToPdf");
        DirUtils.create_with_parents (config_dir, 0755);
        settings_path = Path.build_filename (config_dir, "settings.txt");
        load_settings ();

        build_ui ();
        setup_drop_target ();
        setup_actions ();

        log_message ("LibreOffice Path: %s".printf (soffice_path));
        check_dependencies ();
    }

    private void build_ui () {
        var header = new Gtk.HeaderBar ();

        // Menu button
        var options_section = new Menu ();
        options_section.append ("Set Retry Count…", "win.set-retry");
        options_section.append ("Set Output Folder…", "win.set-output-folder");
        options_section.append ("Export Log…", "win.export-log");

        var paper_submenu = new Menu ();
        paper_submenu.append ("Letter (8.5 x 11)", "win.paper-size::Letter");
        paper_submenu.append ("A4", "win.paper-size::A4");
        paper_submenu.append ("Legal (8.5 x 14)", "win.paper-size::Legal");
        options_section.append_submenu ("Default Paper Size", paper_submenu);

        var about_section = new Menu ();
        about_section.append ("About", "win.about");

        var menu_model = new Menu ();
        menu_model.append_section (null, options_section);
        menu_model.append_section (null, about_section);

        var menu_button = new Gtk.MenuButton ();
        menu_button.icon_name = "open-menu-symbolic";
        menu_button.menu_model = menu_model;
        header.pack_end (menu_button);

        set_titlebar (header);

        // Main layout
        var vbox = new Gtk.Box (Gtk.Orientation.VERTICAL, 8);
        vbox.margin_start = 12;
        vbox.margin_end = 12;
        vbox.margin_top = 12;
        vbox.margin_bottom = 12;

        // Folder row
        var folder_box = new Gtk.Box (Gtk.Orientation.HORIZONTAL, 6);

        folder_entry = new Gtk.Entry ();
        folder_entry.hexpand = true;
        folder_entry.placeholder_text = "Drop a folder here or browse…";

        browse_button = new Gtk.Button.with_label ("Browse");
        browse_button.clicked.connect (on_browse_clicked);

        convert_button = new Gtk.Button.with_label ("Convert");
        convert_button.add_css_class ("suggested-action");
        convert_button.clicked.connect (on_convert_clicked);

        cancel_button = new Gtk.Button.with_label ("Cancel");
        cancel_button.add_css_class ("destructive-action");
        cancel_button.sensitive = false;
        cancel_button.clicked.connect (on_cancel_clicked);

        folder_box.append (folder_entry);
        folder_box.append (browse_button);
        folder_box.append (convert_button);
        folder_box.append (cancel_button);

        vbox.append (folder_box);

        // Progress bar
        progress_bar = new Gtk.ProgressBar ();
        progress_bar.show_text = true;
        vbox.append (progress_bar);

        // Recursive checkbox
        recursive_check = new Gtk.CheckButton.with_label ("Include subfolders");
        recursive_check.active = true;
        vbox.append (recursive_check);

        // Log view
        var scrolled = new Gtk.ScrolledWindow ();
        scrolled.vexpand = true;
        scrolled.hexpand = true;

        log_buffer = new Gtk.TextBuffer (null);
        log_view = new Gtk.TextView.with_buffer (log_buffer);
        log_view.editable = false;
        log_view.monospace = true;
        log_view.wrap_mode = Gtk.WrapMode.WORD_CHAR;

        scrolled.child = log_view;
        vbox.append (scrolled);

        set_child (vbox);
    }

    private void setup_drop_target () {
        var drop_target = new Gtk.DropTarget (typeof (Gdk.FileList), Gdk.DragAction.COPY);

        drop_target.drop.connect ((value) => {
            var file_list = (Gdk.FileList) value;
            var files = file_list.get_files ();

            if (files.length () > 0) {
                var file = files.nth_data (0);
                var path = file.get_path ();

                if (path != null && FileUtils.test (path, FileTest.IS_DIR)) {
                    folder_entry.text = path;
                    log_message ("Folder dropped: %s".printf (path));
                }
            }

            return true;
        });

        folder_entry.add_controller (drop_target);
    }

    private void setup_actions () {
        var retry_action = new SimpleAction ("set-retry", null);
        retry_action.activate.connect (() => { show_retry_dialog (); });
        add_action (retry_action);

        var output_action = new SimpleAction ("set-output-folder", null);
        output_action.activate.connect (() => { show_output_folder_dialog (); });
        add_action (output_action);

        var export_action = new SimpleAction ("export-log", null);
        export_action.activate.connect (() => { show_export_log_dialog (); });
        add_action (export_action);

        var about_action = new SimpleAction ("about", null);
        about_action.activate.connect (() => { show_about_dialog (); });
        add_action (about_action);

        var paper_action = new SimpleAction.stateful ("paper-size", new VariantType ("s"), new Variant.string (paper_size));
        paper_action.activate.connect ((parameter) => {
            if (parameter != null) {
                var val = parameter.get_string ();
                if (val == "Letter" || val == "A4" || val == "Legal") {
                    paper_size = val;
                    paper_action.set_state (new Variant.string (paper_size));
                    log_message ("Default paper size set to %s".printf (paper_size));
                    save_settings ();
                }
            }
        });
        add_action (paper_action);
    }

    private string detect_libreoffice () {
        string[] paths = {
            "/usr/bin/soffice",
            "/usr/lib/libreoffice/program/soffice",
            "/usr/local/bin/soffice",
            "/snap/bin/libreoffice.soffice",
            "/var/lib/flatpak/exports/bin/org.libreoffice.LibreOffice"
        };

        foreach (var p in paths) {
            if (FileUtils.test (p, FileTest.EXISTS)) {
                return p;
            }
        }

        // Try PATH lookup
        var env_path = Environment.get_variable ("PATH");
        if (env_path != null) {
            foreach (var dir in env_path.split (":")) {
                var full = Path.build_filename (dir, "soffice");
                if (FileUtils.test (full, FileTest.EXISTS)) {
                    return full;
                }
            }
        }

        return "soffice";
    }

    private void log_message (string msg) {
        var now = new DateTime.now_local ();
        var line = "%s - %s\n".printf (now.format ("%H:%M:%S"), msg);

        Gtk.TextIter end;
        log_buffer.get_end_iter (out end);
        log_buffer.insert (ref end, line, -1);

        // Auto-scroll to bottom
        Gtk.TextIter end2;
        log_buffer.get_end_iter (out end2);
        var mark = log_buffer.create_mark (null, end2, false);
        log_view.scroll_to_mark (mark, 0, false, 0, 0);

        // Write to log file
        try {
            FileUtils.set_contents (log_file_path,
                log_buffer.text);
        } catch (Error e) {
            // Silently ignore file write errors
        }
    }

    private void on_browse_clicked () {
        var dialog = new Gtk.FileDialog ();
        dialog.title = "Select Folder";

        dialog.select_folder.begin (this, null, (obj, res) => {
            try {
                var folder = dialog.select_folder.end (res);
                if (folder != null) {
                    folder_entry.text = folder.get_path ();
                }
            } catch (Error e) {
                // User cancelled
            }
        });
    }

    private async void on_convert_clicked () {
        var folder = folder_entry.text;

        if (!FileUtils.test (folder, FileTest.IS_DIR)) {
            show_error ("Invalid folder path.");
            return;
        }

        cancellable = new Cancellable ();
        convert_button.sensitive = false;
        cancel_button.sensitive = true;

        yield ensure_template_profile ();

        var files = new GenericArray<string> ();
        enumerate_docs (folder, files);

        total_files = (int) files.length;
        processed_files = 0;

        progress_bar.fraction = 0;
        progress_bar.text = "0 / %d".printf (total_files);

        log_message ("Found %d files.".printf (total_files));

        // Process files with concurrency limit (capped at 8 — soffice doesn't benefit past this)
        int max_concurrent = (int) GLib.get_num_processors ();
        if (max_concurrent > 8) max_concurrent = 8;
        var semaphore = new AsyncSemaphore (max_concurrent);

        var pending = new GenericArray<TaskCompletionSource> ();

        for (int i = 0; i < files.length; i++) {
            if (cancellable.is_cancelled ()) break;

            var file = files[i];
            var task = new TaskCompletionSource ();
            pending.add (task);

            convert_single_file.begin (file, semaphore, task);
        }

        // Wait for all tasks
        for (int i = 0; i < pending.length; i++) {
            yield pending[i].wait_async ();
        }

        convert_button.sensitive = true;
        cancel_button.sensitive = false;

        if (cancellable.is_cancelled ()) {
            log_message ("Operation cancelled.");
        } else {
            log_message ("All conversions completed.");
        }
    }

    private async void convert_single_file (string file, AsyncSemaphore semaphore, TaskCompletionSource task) {
        yield semaphore.acquire ();

        try {
            if (cancellable != null && !cancellable.is_cancelled ()) {
                yield convert_with_retry (file);
            }
        } finally {
            semaphore.release ();
            task.complete ();
        }
    }

    private void enumerate_docs (string folder, GenericArray<string> results) {
        try {
            var dir = Dir.open (folder);
            string? name;

            while ((name = dir.read_name ()) != null) {
                var path = Path.build_filename (folder, name);

                if (FileUtils.test (path, FileTest.IS_DIR)) {
                    if (recursive_check.active) {
                        enumerate_docs (path, results);
                    }
                } else {
                    var lower = name.down ();
                    if (lower.has_suffix (".doc") || lower.has_suffix (".docx")) {
                        results.add (path);
                    }
                }
            }
        } catch (Error e) {
            log_message ("Error reading directory: %s".printf (e.message));
        }
    }

    private async void convert_with_retry (string file) {
        for (int i = 1; i <= retry_count + 1; i++) {
            if (cancellable != null && cancellable.is_cancelled ()) return;

            if (yield convert_to_pdf (file)) return;

            if (i <= retry_count) {
                log_message ("Retry %d/%d for %s".printf (i, retry_count, Path.get_basename (file)));
            }
        }

        log_message ("Failed: %s".printf (file));
    }

    private async bool convert_to_pdf (string file) {
        var output_dir = custom_output_folder ?? Path.get_dirname (file);

        var basename = Path.get_basename (file);
        var dot_index = basename.last_index_of (".");
        var name_without_ext = (dot_index > 0) ? basename.substring (0, dot_index) : basename;
        var expected_output = Path.build_filename (output_dir, name_without_ext + ".pdf");

        var profile_dir = Path.build_filename (profile_root, GLib.Uuid.string_random ());
        try {
            copy_directory (template_profile_dir, profile_dir);
            write_paper_size_xcu (profile_dir);
        } catch (Error e) {
            log_message ("Profile setup failed: %s".printf (e.message));
            return false;
        }
        var user_installation = "-env:UserInstallation=file://" + profile_dir;

        try {
            var launcher = new Subprocess.newv (
                { soffice_path, user_installation,
                  "--headless", "--norestore", "--nologo", "--nofirststartwizard",
                  "--nodefault", "--nolockcheck",
                  "--convert-to", "pdf",
                  "--outdir", output_dir, file },
                SubprocessFlags.STDOUT_PIPE | SubprocessFlags.STDERR_PIPE
            );

            string? stdout_buf;
            string? stderr_buf;
            yield launcher.communicate_utf8_async (null, cancellable, out stdout_buf, out stderr_buf);

            if (FileUtils.test (expected_output, FileTest.EXISTS)) {
                processed_files++;
                log_message ("Complete: %s".printf (basename));
                update_progress ();
                return true;
            }

            // Conversion failed — log any output from LibreOffice
            if (stderr_buf != null && stderr_buf.strip () != "") {
                log_message ("LibreOffice error: %s".printf (stderr_buf.strip ()));
            }
            if (stdout_buf != null && stdout_buf.strip () != "") {
                log_message ("LibreOffice: %s".printf (stdout_buf.strip ()));
            }
        } catch (Error e) {
            if (!(e is IOError.CANCELLED)) {
                log_message (e.message);
            }
        } finally {
            remove_directory_recursive (profile_dir);
        }

        return false;
    }

    private async void ensure_template_profile () {
        if (template_ready) return;
        var user_dir = Path.build_filename (template_profile_dir, "user");
        if (FileUtils.test (user_dir, FileTest.IS_DIR)) {
            template_ready = true;
            return;
        }

        log_message ("Initializing LibreOffice profile template (one-time)...");
        DirUtils.create_with_parents (template_profile_dir, 0755);
        var user_installation = "-env:UserInstallation=file://" + template_profile_dir;

        try {
            var launcher = new Subprocess.newv (
                { soffice_path, user_installation,
                  "--headless", "--norestore", "--nologo", "--nofirststartwizard",
                  "--nodefault", "--nolockcheck", "--terminate_after_init" },
                SubprocessFlags.STDOUT_PIPE | SubprocessFlags.STDERR_PIPE
            );

            var timeout_cancel = new Cancellable ();
            uint timeout_id = Timeout.add_seconds (60, () => {
                timeout_cancel.cancel ();
                return false;
            });

            string? stdout_buf;
            string? stderr_buf;
            try {
                yield launcher.communicate_utf8_async (null, timeout_cancel, out stdout_buf, out stderr_buf);
            } catch (Error e) {
                launcher.force_exit ();
                log_message ("Template init interrupted: %s".printf (e.message));
            }
            Source.remove (timeout_id);
            template_ready = true;
        } catch (Error e) {
            log_message ("Template init failed: %s".printf (e.message));
        }
    }

    private void copy_directory (string source, string dest) throws Error {
        DirUtils.create_with_parents (dest, 0755);
        var dir = Dir.open (source);
        string? name;
        while ((name = dir.read_name ()) != null) {
            var src_path = Path.build_filename (source, name);
            var dst_path = Path.build_filename (dest, name);
            if (FileUtils.test (src_path, FileTest.IS_DIR)) {
                copy_directory (src_path, dst_path);
            } else {
                var src_file = File.new_for_path (src_path);
                var dst_file = File.new_for_path (dst_path);
                src_file.copy (dst_file, FileCopyFlags.OVERWRITE, null, null);
            }
        }
    }

    private void remove_directory_recursive (string path) {
        if (!FileUtils.test (path, FileTest.IS_DIR)) return;
        try {
            var dir = Dir.open (path);
            string? name;
            while ((name = dir.read_name ()) != null) {
                var child = Path.build_filename (path, name);
                if (FileUtils.test (child, FileTest.IS_DIR)) {
                    remove_directory_recursive (child);
                } else {
                    FileUtils.unlink (child);
                }
            }
        } catch (Error e) {
            // ignore
        }
        DirUtils.remove (path);
    }

    private void write_paper_size_xcu (string profile_dir) {
        int width, height;
        switch (paper_size) {
            case "A4":    width = 21000; height = 29700; break;
            case "Legal": width = 21590; height = 35560; break;
            default:      width = 21590; height = 27940; break;
        }

        var user_dir = Path.build_filename (profile_dir, "user");
        DirUtils.create_with_parents (user_dir, 0755);
        var xcu_path = Path.build_filename (user_dir, "registrymodifications.xcu");

        var xml = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Writer/DefaultPageSize"><prop oor:name="Width" oor:op="fuse"><value>%d</value></prop></item>
 <item oor:path="/org.openoffice.Office.Writer/DefaultPageSize"><prop oor:name="Height" oor:op="fuse"><value>%d</value></prop></item>
 <item oor:path="/org.openoffice.Office.Common/Save/Document"><prop oor:name="PrinterIndependentLayout" oor:op="fuse"><value>2</value></prop></item>
</oor:items>
""".printf (width, height);

        try {
            FileUtils.set_contents (xcu_path, xml);
        } catch (Error e) {
            log_message ("Failed to write XCU: %s".printf (e.message));
        }
    }

    private void load_settings () {
        if (!FileUtils.test (settings_path, FileTest.EXISTS)) return;
        try {
            string contents;
            FileUtils.get_contents (settings_path, out contents);
            foreach (var line in contents.split ("\n")) {
                var eq = line.index_of ("=");
                if (eq <= 0) continue;
                var key = line.substring (0, eq).strip ();
                var val = line.substring (eq + 1).strip ();
                if (key == "paper" && (val == "Letter" || val == "A4" || val == "Legal")) {
                    paper_size = val;
                }
            }
        } catch (Error e) {
            // ignore
        }
    }

    private void save_settings () {
        var content = "paper=%s\n".printf (paper_size);
        try {
            FileUtils.set_contents (settings_path, content);
        } catch (Error e) {
            // ignore
        }
    }

    private void update_progress () {
        if (total_files > 0) {
            progress_bar.fraction = (double) processed_files / (double) total_files;
            progress_bar.text = "%d / %d".printf (processed_files, total_files);
        }
    }

    private void on_cancel_clicked () {
        if (cancellable != null) {
            cancellable.cancel ();
            log_message ("Cancel requested...");
        }
    }

    private void show_retry_dialog () {
        var dialog = new Gtk.AlertDialog ("Set Retry Count");
        dialog.detail = "Current retry count: %d\nEnter a new value:".printf (retry_count);
        dialog.buttons = { "Cancel", "OK" };
        dialog.default_button = 1;
        dialog.cancel_button = 0;

        dialog.choose.begin (this, null, (obj, res) => {
            try {
                var button = dialog.choose.end (res);
                if (button == 1) {
                    show_retry_entry_dialog ();
                }
            } catch (Error e) {
                // cancelled
            }
        });
    }

    private void show_retry_entry_dialog () {
        var win = new Gtk.Window ();
        win.title = "Set Retry Count";
        win.transient_for = this;
        win.modal = true;
        win.default_width = 300;

        var box = new Gtk.Box (Gtk.Orientation.VERTICAL, 12);
        box.margin_start = 12;
        box.margin_end = 12;
        box.margin_top = 12;
        box.margin_bottom = 12;

        var label = new Gtk.Label ("Retry count:");
        box.append (label);

        var entry = new Gtk.Entry ();
        entry.text = retry_count.to_string ();
        entry.input_purpose = Gtk.InputPurpose.DIGITS;
        box.append (entry);

        var btn_box = new Gtk.Box (Gtk.Orientation.HORIZONTAL, 6);
        btn_box.halign = Gtk.Align.END;

        var cancel_btn = new Gtk.Button.with_label ("Cancel");
        cancel_btn.clicked.connect (() => { win.close (); });

        var ok_btn = new Gtk.Button.with_label ("OK");
        ok_btn.add_css_class ("suggested-action");
        ok_btn.clicked.connect (() => {
            int val;
            if (int.try_parse (entry.text, out val) && val >= 0) {
                retry_count = val;
                log_message ("Retry set to %d".printf (retry_count));
            }
            win.close ();
        });

        btn_box.append (cancel_btn);
        btn_box.append (ok_btn);
        box.append (btn_box);

        win.child = box;
        win.present ();
    }

    private void show_output_folder_dialog () {
        var dialog = new Gtk.FileDialog ();
        dialog.title = "Select Output Folder";

        dialog.select_folder.begin (this, null, (obj, res) => {
            try {
                var folder = dialog.select_folder.end (res);
                if (folder != null) {
                    custom_output_folder = folder.get_path ();
                    log_message ("Output folder: %s".printf (custom_output_folder));
                }
            } catch (Error e) {
                // cancelled
            }
        });
    }

    private void show_export_log_dialog () {
        var dialog = new Gtk.FileDialog ();
        dialog.title = "Export Log";
        dialog.initial_name = "conversion_log.txt";

        dialog.save.begin (this, null, (obj, res) => {
            try {
                var file = dialog.save.end (res);
                if (file != null) {
                    var dest = file.get_path ();
                    FileUtils.set_contents (dest, log_buffer.text);
                    log_message ("Log exported to: %s".printf (dest));
                }
            } catch (Error e) {
                // cancelled
            }
        });
    }

    private void check_dependencies () {
        string[] checks = {
            "which java",
            "dpkg-query -W libreoffice-writer",
            "dpkg-query -W libreoffice-java-common"
        };

        string[] messages = {
            "Warning: Java not found. Install with: sudo apt install default-jre",
            "Warning: libreoffice-writer not found. Install with: sudo apt install libreoffice-writer",
            "Warning: libreoffice-java-common not found. Install with: sudo apt install libreoffice-java-common"
        };

        for (int i = 0; i < checks.length; i++) {
            try {
                int status;
                Process.spawn_command_line_sync (checks[i], null, null, out status);
                if (status != 0) {
                    log_message (messages[i]);
                }
            } catch (Error e) {
                // Command not available (e.g. non-Debian system), skip
            }
        }
    }

    private void show_about_dialog () {
        var about = new Gtk.AboutDialog ();
        about.transient_for = this;
        about.modal = true;
        about.program_name = "DOC to PDF Converter";
        about.version = "1.0.1";
        about.authors = { "Kevin Adams" };
        about.website = "https://github.com/KevinAdams05/LibreDocToPdf";
        about.website_label = "GitHub Repository";
        about.license_type = Gtk.License.MIT_X11;
        about.copyright = "© 2025 Kevin Adams";
        about.comments = "A simple DOC/DOCX to PDF converter powered by LibreOffice.";
        about.add_credit_section ("Credits", {
            "Icon: icon-icons.com (pdf-reader-pro-macos-bigsur)",
            "Powered by LibreOffice (libreoffice.org)"
        });
        about.present ();
    }

    private void show_error (string message) {
        var dialog = new Gtk.AlertDialog (message);
        dialog.buttons = { "OK" };
        dialog.choose.begin (this, null, (obj, res) => {
            try { dialog.choose.end (res); } catch (Error e) {}
        });
    }
}
