using Aras.IOM;

public abstract partial class InnovatorBase {

    private static Innovator Inn;

    public static string Url = "http://localhost/2023/";
    public static string DB = "Dev2023";
    public static string User = "admin";
    public static string Password = "innovator";
    public static int TimeoutSeconds = 90;

    private static void createInnovator() {
        HttpServerConnection MyConnection = IomFactory.CreateHttpServerConnection(Url, DB, User, Password);
        if (TimeoutSeconds > 0) {
            MyConnection.Timeout = 1000 * TimeoutSeconds;
        }
        Inn = new Innovator(MyConnection);
        MyConnection.Login();
    }

    public static Innovator getInnovator() {
        if (Inn is null)
            createInnovator();
        return Inn;
    }

    public static object ToStringAddress() {
        return Url + " DB: " + DB + " User:" + User;
    }

}