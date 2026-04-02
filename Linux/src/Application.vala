public class DocToPdfApp : Gtk.Application {
    public DocToPdfApp () {
        Object (
            application_id: "com.github.kevinadams05.doctopdf",
            flags: ApplicationFlags.FLAGS_NONE
        );
    }

    protected override void activate () {
        var win = new MainWindow (this);
        win.present ();
    }

    public static int main (string[] args) {
        var app = new DocToPdfApp ();
        return app.run (args);
    }
}
