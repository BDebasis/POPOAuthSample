/*
 * By David Barrett, Microsoft Ltd. 2021. Use at your own risk.  No warranties are given.
 * 
 * DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 * */

using System;
using Microsoft.Identity.Client;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Text;
using System.Net.Security;

namespace POPOAuthSample
{
    class Program
    {
        private static TcpClient _popClient = null;
        private static SslStream _sslStream = null;

        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine($"Syntax: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.exe <TenantId> <ApplicationId>");
                return;
            }
            var task = TestPOP(args[1], args[0]);
            task.Wait();
        }

        static async Task TestPOP(string ClientId, string TenantId)
        {

            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ClientId,
                TenantId = TenantId
            };

            Console.WriteLine("Building application");
            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions)
                .WithRedirectUri("http://localhost")
                .Build();

            var popScope = new string[] { "https://outlook.office.com/POP.AccessAsUser.All" };

            try
            {
                // Make the interactive token request
                Console.WriteLine("Requesting access token (user must log-in via browser)");
                var authResult = await pca.AcquireTokenInteractive(popScope).ExecuteAsync();
                if (String.IsNullOrEmpty(authResult.AccessToken))
                {
                    Console.WriteLine("No token received");
                    return;
                }
                Console.WriteLine($"Token received for {authResult.Account.Username}");

                // Use the token to connect to POP service
                RetrieveMessages(authResult);
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }
            Console.WriteLine("Finished");

        }

        static string ReadSSLStream()
        {
            int bytes = -1;
            byte[] buffer = new byte[2048];
            bytes = _sslStream.Read(buffer, 0, buffer.Length);
            string response = Encoding.ASCII.GetString(buffer, 0, bytes);
            Console.WriteLine(response);
            return response;
        }

        static void WriteSSLStream(string Data)
        {
            _sslStream.Write(Encoding.ASCII.GetBytes($"{Data}{Environment.NewLine}"));
            _sslStream.Flush();
            Console.WriteLine(Data);
        }

        static void RetrieveMessages(AuthenticationResult authResult)
        {
            try
            {
                _popClient = new TcpClient("outlook.office365.com", 995);
                _sslStream = new SslStream(_popClient.GetStream());
                _sslStream.AuthenticateAsClient("outlook.office365.com");

                ReadSSLStream();

                // Initiate OAuth login
                WriteSSLStream("AUTH XOAUTH2");
                if (ReadSSLStream().StartsWith("+"))
                {
                    // Send OAuth token
                    WriteSSLStream(XOauth2(authResult));
                    if (ReadSSLStream().StartsWith("+OK"))
                    {
                        // Logged in, get status
                        WriteSSLStream("STAT");
                        ReadSSLStream();

                        // And list of messages
                        WriteSSLStream("LIST");
                        ReadSSLStream();
                    }
                }
                WriteSSLStream("QUIT");
                ReadSSLStream();

                // Tidy up
                Console.WriteLine("Closing connection");
                _popClient.Close();
                _sslStream.Dispose();
            }
            catch (SocketException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static string XOauth2(AuthenticationResult authResult)
        {
            // Create the log-in code, which is a base 64 encoded combination of user and auth token

            string ctrlA = $"{(char)1}";
            string login = $"user={authResult.Account.Username}{ctrlA}auth=Bearer {authResult.AccessToken}{ctrlA}{ctrlA}";
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(login);
            return Convert.ToBase64String(plainTextBytes);
        }
    }
}
