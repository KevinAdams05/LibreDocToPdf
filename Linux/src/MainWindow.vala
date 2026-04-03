public class MainWindow : Gtk.ApplicationWindow {
    private Gtk.Entry folder_entry;
    private Gtk.Button convert_button;
    private Gtk.Button cancel_button;
    private Gtk.Button browse_button;
    private Gtk.ProgressBar progress_bar;
    private Gtk.TextView log_view;
    private Gtk.TextBuffer log_buffer;

    private string soffice_path;
    private string? custom_output_folder = null;
    private int retry_count = 2;
    private int total_files = 0;
    private int processed_files = 0;
    private Cancellable? cancellable = null;
    private string log_file_path;

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

        build_ui ();
        setup_drop_target ();
        setup_actions ();

        log_message ("LibreOffice Path: %s".printf (soffice_path));
    }

    private void build_ui () {
        var header = new Gtk.HeaderBar ();

        // Menu button
        var menu_model = new Menu ();
        menu_model.append ("Set Retry Count…", "win.set-retry");
        menu_model.append ("Set Output Folder…", "win.set-output-folder");
        menu_model.append ("Export Log…", "win.export-log");

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

        var files = new GenericArray<string> ();
        enumerate_docs (folder, files);

        total_files = (int) files.length;
        processed_files = 0;

        progress_bar.fraction = 0;
        progress_bar.text = "0 / %d".printf (total_files);

        log_message ("Found %d files.".printf (total_files));

        // Process files with concurrency limit
        int max_concurrent = (int) GLib.get_num_processors ();
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
                    enumerate_docs (path, results);
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

        try {
            var launcher = new Subprocess.newv (
                { soffice_path, "--headless", "--convert-to", "pdf",
                  "--outdir", output_dir, file },
                SubprocessFlags.STDOUT_SILENCE | SubprocessFlags.STDERR_SILENCE
            );

            yield launcher.wait_async (cancellable);

            if (launcher.get_successful ()) {
                processed_files++;
                log_message ("Complete: %s".printf (Path.get_basename (file)));
                update_progress ();
                return true;
            }
        } catch (Error e) {
            if (!(e is IOError.CANCELLED)) {
                log_message (e.message);
            }
        }

        return false;
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

    private void show_error (string message) {
        var dialog = new Gtk.AlertDialog (message);
        dialog.buttons = { "OK" };
        dialog.choose.begin (this, null, (obj, res) => {
            try { dialog.choose.end (res); } catch (Error e) {}
        });
    }
}
