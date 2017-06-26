using System.Windows;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Oauth2.v2;
using Google.Apis.Oauth2.v2.Data;
using Google.Apis.Services;
using MailBee;
using MailBee.ImapMail;
using System.Threading;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Connect();
        }

        void Connect()
        {
            // Request Gmail IMAP/SMTP scope and the e-mail address scope.
            string[] scopes = new string[] { "https://mail.google.com/", Oauth2Service.Scope.UserinfoEmail };

            MessageBox.Show("Requesting authorization");
            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = "Your Client ID, like 0123-abcd.apps.googleusercontent.com",
                    ClientSecret = "Your Client Secret, like 0a1b2c3d",
                },
                 scopes,
                 "user",
                 CancellationToken.None).Result;
            MessageBox.Show("Authorization granted or not required (if the saved access token already available)");

            if (credential.Token.IsExpired(credential.Flow.Clock))
            {
                MessageBox.Show("The access token has expired, refreshing it");
                if (credential.RefreshTokenAsync(CancellationToken.None).Result)
                {
                    MessageBox.Show("The access token is now refreshed");
                }
                else
                {
                    MessageBox.Show("The access token has expired but we can't refresh it :(");
                    return;
                }
            }
            else
            {
                MessageBox.Show("The access token is OK, continue");
            }

            MessageBox.Show("Requesting the e-mail address of the user from Google");

            // Sometimes, you may also need to set Initializer.ApplicationName property.
            // In our tests, setting just Initializer.HttpClientInitializer was enough for Google.
            Oauth2Service oauthService = new Oauth2Service(
                new BaseClientService.Initializer() { HttpClientInitializer = credential });

            // Userinfo.Get may crash if you run the app under debugger (a bug in Google API).
            // If this happens, use "Start without debugging" instead.
            Userinfoplus userInfo = oauthService.Userinfo.Get().ExecuteAsync().Result;
            string userEmail = userInfo.Email;
            MessageBox.Show("E-mail address is " + userEmail);

            // Build XOAUTH2 token. Can be used with Gmail IMAP or SMTP.
            string xoauthKey = OAuth2.GetXOAuthKeyStatic(userEmail, credential.Token.AccessToken);

            // Uncomment and set your key if you haven't specified it in app.config or Windows registry.
            // MailBee.Global.LicenseKey = "Your MNXXX-XXXX-XXXX key here";

            // Finally, use MailBee.NET to list the number of e-mails in Inbox.

            Imap imp = new Imap();

            // Logging is not necessary but useful for debugging.
            imp.Log.Enabled = true;
            imp.Log.Filename = @"C:\Temp\log.txt";
            imp.Log.HidePasswords = false;
            imp.Log.Clear();

            imp.Connect("imap.gmail.com");

            // This is the where IMAP XOAUTH2 actually occurs.
            imp.Login(null, xoauthKey, AuthenticationMethods.SaslOAuth2,
                MailBee.AuthenticationOptions.None, null);

            // If we're here, we're lucky.
            imp.SelectFolder("INBOX");
            MessageBox.Show(imp.MessageCount.ToString() + " e-mails in Inbox");
            imp.Disconnect();
        }
    }
}
