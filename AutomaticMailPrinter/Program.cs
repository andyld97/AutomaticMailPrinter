using MailKit.Net.Imap;
using MailKit;
using System;
using MailKit.Search;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace AutomaticMailPrinter
{
    internal class Program
    {
        private static System.Threading.Timer timer;
        private static readonly Mutex AppMutex = new Mutex(false, "c75adf4e-765c-4529-bf7a-90dd76cd386a");

        private static string ImapServer, MailAddress, Password, PrinterName;
        public static string WebHookUrl { get; private set; }
        private static string[] Filter = new string[0];
        private static int ImapPort;

        private static ImapClient client = new ImapClient();
        private static IMailFolder inbox;

        private static object sync = new object();

        static void Main(string[] args)
        {
            if (!AppMutex.WaitOne(TimeSpan.FromSeconds(1), false))
            {
                MessageBox.Show(Properties.Resources.strInstanceAlreadyRunning, Properties.Resources.strError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            System.AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            try
            {
                string configPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                var configDocument = System.Text.Json.JsonSerializer.Deserialize<JsonDocument>(System.IO.File.ReadAllText(configPath));

                ImapServer = configDocument.RootElement.GetProperty("imap_server").GetString();
                ImapPort = configDocument.RootElement.GetProperty("imap_port").GetInt32();
                MailAddress = configDocument.RootElement.GetProperty("mail").GetString();
                Password = configDocument.RootElement.GetProperty("password").GetString();
                PrinterName = configDocument.RootElement.GetProperty("printer_name").GetString();

                try
                {
                    // Can be empty or even may not existing ...
                    WebHookUrl = configDocument.RootElement.GetProperty("webhook_url").GetString();
                }
                catch { }
                int intervalInSecods = configDocument.RootElement.GetProperty("timer_interval_in_seconds").GetInt32();

                var filterProperty = configDocument.RootElement.GetProperty("filter");
                int counter = 0;
                Filter = new string[filterProperty.GetArrayLength()];
                foreach (var word in filterProperty.EnumerateArray())
                    Filter[counter++] = word.GetString().ToLower();

                Logger.LogInfo(string.Format(Properties.Resources.strConnectToMailServer, $"\"{ImapServer}:{ImapPort}\""));

                client = new ImapClient();
                client.Connect(ImapServer, ImapPort, true);
                client.Authenticate(MailAddress, Password);

                // The Inbox folder is always available on all IMAP servers...
                inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);

                // Clear all old mails
                Logger.LogInfo(Properties.Resources.strDeleteOldMessagesFromInBox, sendWebHook: true);
                int count = 0;
                foreach (var uid in inbox.Search(SearchQuery.Seen))
                {
                    var message = inbox.GetMessage(uid);
                    string subject = message.Subject.ToLower();

                    if (Filter.Any(f => subject.Contains(f)))
                    {
                        // Delete mail https://stackoverflow.com/a/24204804/6237448
                        inbox.SetFlags(uid, MessageFlags.Deleted, true);
                        count++;
                    }
                }
                if (count > 0)
                    inbox.Expunge();

                Logger.LogInfo(string.Format(Properties.Resources.strDeleteNMessagesFromInBox, count), sendWebHook: true);
                timer = new System.Threading.Timer(Timer_Tick, null, 0, intervalInSecods * 1000);
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToReadConfigFile, ex);
            }

            while (true)
            {
                System.Threading.Thread.Sleep(500);
            }

            AppMutex.ReleaseMutex();
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
                Logger.LogError("Unhandled Exception recieved!", ex);
            else if (e.ExceptionObject != null)
                Logger.LogError($"Unhandled Exception recieved: {e.ExceptionObject}");
            else
                Logger.LogError("Unhandled Exception but exception object is empty :(");
        }

        private static void Timer_Tick(object state)
        {
            try
            {
                lock (sync)
                {
                    Logger.LogInfo(Properties.Resources.strLookingForUnreadMails);
                    bool found = false;

                    if (!client.IsAuthenticated || !client.IsConnected)
                    {
                        Logger.LogWarning(Properties.Resources.strMailClientIsNotConnectedAnymore);

                        try
                        {
                            client = new ImapClient();
                            client.Connect(ImapServer, ImapPort, true);
                            client.Authenticate(MailAddress, Password);

                            // The Inbox folder is always available on all IMAP servers...
                            inbox = client.Inbox;

                            Logger.LogInfo(Properties.Resources.strConnectionEstablishedSuccess, sendWebHook: true);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError(Properties.Resources.strFailedToConnect, ex);
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
                            Logger.LogInfo($"{string.Format(Properties.Resources.strFoundUnreadMail, Filter.Where(f => subject.Contains(f)).FirstOrDefault())} {message.Subject}");

                            // Print mail
                            Logger.LogInfo(string.Format(Properties.Resources.strPrintMessage, message.Subject, PrinterName));
                            PrintHtmlPage(message.HtmlBody);

                            // Delete mail https://stackoverflow.com/a/24204804/6237448
                            Logger.LogInfo(Properties.Resources.strMarkMailAsDeleted);                     
                            inbox.SetFlags(uid, MessageFlags.Deleted, true);

                            found = true;

                            PlaySound();
                        }
                    }

                    if (!found)
                        Logger.LogInfo(Properties.Resources.strNoUnreadMailFound);
                    else
                    {
                        Logger.LogInfo(Properties.Resources.strExpungeMails);
                        inbox.Expunge();
                    }

                    // Do not disconnect here!
                    // client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToRecieveMails, ex);
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
                Logger.LogError(Properties.Resources.strFailedToPlaySound, ex);
            }
        }

        public static void PrintHtmlPage(string htmlContent)
        {
            try
            {
                string path = System.IO.Path.GetTempFileName();
                System.IO.File.WriteAllText(path, htmlContent);
                PrintHtmlPages(PrinterName, path);
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToPrintMail, ex);
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