using MailKit.Net.Imap;
using MailKit;
using System;
using MailKit.Search;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Linq;

namespace AutomaticMailPrinter
{
    internal class Program
    {
        private static System.Threading.Timer timer;

        private static string ImapServer, MailAddress, Password, PrinterName;
        private static string[] Filter = new string[0];
        private static int ImapPort;

        private static ImapClient client = new ImapClient();
        private static IMailFolder inbox;

        static void Main(string[] args)
        {
            // This doesn't work without a console
            // Console.Title = Properties.Resources.strAppTitle;

            try
            {
                string configPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                var configDocument = System.Text.Json.JsonSerializer.Deserialize<JsonDocument>(System.IO.File.ReadAllText(configPath));

                ImapServer = configDocument.RootElement.GetProperty("imap_server").GetString();
                ImapPort = configDocument.RootElement.GetProperty("imap_port").GetInt32();
                MailAddress = configDocument.RootElement.GetProperty("mail").GetString();
                Password = configDocument.RootElement.GetProperty("password").GetString();
                PrinterName = configDocument.RootElement.GetProperty("printer_name").GetString();
                int intervalInSecods = configDocument.RootElement.GetProperty("timer_interval_in_seconds").GetInt32();
                
                var filterProperty = configDocument.RootElement.GetProperty("filter");
                int counter = 0;
                Filter = new string[filterProperty.GetArrayLength()];
                foreach (var word in filterProperty.EnumerateArray())
                    Filter[counter++] = word.GetString().ToLower();

                Console.WriteLine(string.Format(Properties.Resources.strConnectToMailServer, $"\"{ImapServer}:{ImapPort}\""));

                client = new ImapClient();
                client.Connect(ImapServer, ImapPort, true);
                client.Authenticate(MailAddress, Password);

                // The Inbox folder is always available on all IMAP servers...
                inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                timer = new System.Threading.Timer(Timer_Tick, null, 0, intervalInSecods * 1000);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{Properties.Resources.strFailedToReadConfigFile}: {ex.Message}");
                Console.ResetColor();
            }

            while (true)
            {
                System.Threading.Thread.Sleep(500);
            }
        }

        private static void Timer_Tick(object state)
        {
            try
            {
                Console.WriteLine(Properties.Resources.strLookingForUnreadMails);
                bool found = false;

                if (!client.IsAuthenticated || !client.IsConnected)
                {
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine(Properties.Resources.strMailClientIsNotConnectedAnymore);
                    Console.ResetColor();

                    try
                    {
                        client = new ImapClient();
                        client.Connect(ImapServer, ImapPort, true);
                        client.Authenticate(MailAddress, Password);

                        // The Inbox folder is always available on all IMAP servers...
                        inbox = client.Inbox;

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(Properties.Resources.strConnectionEstablishedSuccess);
                        Console.ResetColor();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"{Properties.Resources.strFailedToConnect}: {ex.Message}!");
                        Console.ResetColor();
                        return;
                    }

                }

                inbox.Open(FolderAccess.ReadWrite);

                foreach (var uid in inbox.Search(SearchQuery.NotSeen))
                {
                    var message = inbox.GetMessage(uid);
                    string subject = message.Subject.ToLower();
                    if (Filter.Any(f => subject.Contains(f)))
                    {
                        // Print text
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"{string.Format(Properties.Resources.strFoundUnreadMail, Filter.Where(f => subject.Contains(f)).FirstOrDefault())} {message.Subject}");

                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine(Properties.Resources.strMarkMailAsRead);

                        // Mark mail as read
                        inbox.SetFlags(uid, MessageFlags.Seen, true);

                        Console.WriteLine(string.Format(Properties.Resources.strPrintMessage, message.Subject, PrinterName));
                        Console.ResetColor();

                        PrintHtmlPage(message.HtmlBody);
                        found = true;

                        PlaySound();
                    }
                }

                if (!found)
                {
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.WriteLine(Properties.Resources.strNoUnreadMailFound);
                    Console.ResetColor();
                }

                // Do not disconnect here!
                // client.Disconnect(true);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{Properties.Resources.strFailedToRecieveMails}: {ex.Message}");
                Console.ResetColor();
            }
        }
        

        public static void PlaySound()
        {
            try
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sounds", "sound.wav"));
                player.Play();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{Properties.Resources.strFailedToPlaySound}: {ex.Message}");
                Console.ResetColor();
            }
        }

        public static void PrintHtmlPage(string htmlContent)
        {
            try
            {
                string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.html");
                System.IO.File.WriteAllText(path, htmlContent);
                PrintHtmlPages(PrinterName, path);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{Properties.Resources.strFailedToPrintMail}: {ex.Message}");
                Console.ResetColor();
            }
        }

        public static bool? PrintHtmlPages(string printer, params string[] urls)
        {
            // Spawn the code to print the packing slips
            var info = new ProcessStartInfo();
            info.Arguments = $"-p \"{printer}\" -a A4 \"{string.Join("\" \"", urls)}\"";
            var pathToExe = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            info.FileName = Path.Combine(pathToExe, "PrintHtml", "PrintHtml.exe");
            Process.Start(info);


            return null;
            /*using (var p = Process.Start(info))
            {
                // Wait until it is finished
                while (!p.HasExited)
                    System.Threading.Thread.Sleep(10);

                // Return the exit code
                return p.ExitCode == 0;
            }*/
        }
    }
}