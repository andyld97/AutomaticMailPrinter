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
using MimeKit;

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
        private static bool PrintPdfAttachments = false;

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
            int intervalInSecods = 60;
            try
            {
                // Prefer YAML if present; else fallback to JSON
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string yamlPath = System.IO.Path.Combine(baseDir, "config.yml");
                string jsonPath = System.IO.Path.Combine(baseDir, "config.json");

                if (System.IO.File.Exists(yamlPath))
                {
                    LoadConfigFromYaml(yamlPath, ref intervalInSecods);
                }
                else
                {
                    var configDocument = System.Text.Json.JsonSerializer.Deserialize<JsonDocument>(System.IO.File.ReadAllText(jsonPath));

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

                    intervalInSecods = configDocument.RootElement.GetProperty("timer_interval_in_seconds").GetInt32();

                    try
                    {
                        PrintPdfAttachments = configDocument.RootElement.GetProperty("print_pdf_attachments").GetBoolean();
                    }
                    catch { }

                    var filterProperty = configDocument.RootElement.GetProperty("filter");
                    int counter = 0;
                    Filter = new string[filterProperty.GetArrayLength()];
                    foreach (var word in filterProperty.EnumerateArray())
                        Filter[counter++] = word.GetString().ToLower();
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToReadConfigFile, ex);
            }

            try
            {
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

            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToRecieveMails, ex);
            }
            finally
            {
                timer = new System.Threading.Timer(Timer_Tick, null, 0, intervalInSecods * 1000);
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
                            Logger.LogInfo($"{string.Format(Properties.Resources.strFoundUnreadMail, Filter.FirstOrDefault(f => subject.Contains(f)))} {message.Subject}");

                            // Print mail
                            Logger.LogInfo(string.Format(Properties.Resources.strPrintMessage, message.Subject, PrinterName));

                            if (PrintPdfAttachments)
                            {
                                var pdfs = SavePdfAttachmentsToTempFiles(message);
                                if (pdfs != null && pdfs.Any())
                                {
                                    PrintPdfFiles(PrinterName, pdfs.ToArray());
                                }
                                else
                                {
                                    // Fallback to HTML if no PDFs found
                                    PrintHtmlPage(message.HtmlBody);
                                }
                            }
                            else
                            {
                                PrintHtmlPage(message.HtmlBody);
                            }

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
            var info = new ProcessStartInfo
            {
                Arguments = $"-p \"{printer}\" -a A4 \"{string.Join("\" \"", urls)}\""
            };
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

        private static void LoadConfigFromYaml(string yamlPath, ref int intervalInSecods)
        {
            try
            {
                var lines = System.IO.File.ReadAllLines(yamlPath);
                var values = ParseSimpleYaml(lines);

                ImapServer = GetDictString(values, "imap_server");
                ImapPort = GetDictInt(values, "imap_port", 993);
                MailAddress = GetDictString(values, "mail");
                Password = GetDictString(values, "password");
                PrinterName = GetDictString(values, "printer_name");
                WebHookUrl = GetDictString(values, "webhook_url");
                intervalInSecods = GetDictInt(values, "timer_interval_in_seconds", intervalInSecods);
                PrintPdfAttachments = GetDictBool(values, "print_pdf_attachments", false);

                if (values.TryGetValue("filter", out var filterObj) && filterObj is System.Collections.Generic.List<string> list)
                {
                    Filter = list.Select(s => (s ?? string.Empty).ToLower()).ToArray();
                }
                else if (values.TryGetValue("filter", out var filterStr) && filterStr is string s && !string.IsNullOrWhiteSpace(s))
                {
                    Filter = s.Split(',').Select(x => x.Trim().ToLower()).Where(x => !string.IsNullOrEmpty(x)).ToArray();
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("Failed to read YAML config.", ex);
                throw;
            }
        }

        private static System.Collections.Generic.Dictionary<string, object> ParseSimpleYaml(string[] lines)
        {
            var dict = new System.Collections.Generic.Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            string currentListKey = null;
            System.Collections.Generic.List<string> currentList = null;

            foreach (var rawLine in lines)
            {
                var line = rawLine;
                if (line == null) continue;
                line = line.Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("#")) continue;

                // list item
                if (line.StartsWith("- "))
                {
                    if (currentListKey != null)
                    {
                        currentList = currentList ?? new System.Collections.Generic.List<string>();
                        var item = line.Substring(2).Trim();
                        currentList.Add(Unquote(item));
                        dict[currentListKey] = currentList;
                    }
                    continue;
                }

                var idx = line.IndexOf(':');
                if (idx <= 0)
                    continue;

                var key = line.Substring(0, idx).Trim();
                var value = line.Substring(idx + 1).Trim();

                if (string.IsNullOrEmpty(value))
                {
                    // start list or nested section; we only support simple lists here
                    currentListKey = key;
                    currentList = new System.Collections.Generic.List<string>();
                    dict[currentListKey] = currentList;
                }
                else
                {
                    currentListKey = null;
                    dict[key] = Unquote(value);
                }
            }

            return dict;
        }

        private static string Unquote(string v)
        {
            if (string.IsNullOrEmpty(v)) return v;
            if ((v.StartsWith("\"") && v.EndsWith("\"")) || (v.StartsWith("'") && v.EndsWith("'")))
                return v.Substring(1, v.Length - 2);
            return v;
        }

        private static string GetDictString(System.Collections.Generic.Dictionary<string, object> dict, string key)
        {
            if (dict.TryGetValue(key, out var v) && v != null)
                return Convert.ToString(v);
            return null;
        }

        private static int GetDictInt(System.Collections.Generic.Dictionary<string, object> dict, string key, int defaultValue)
        {
            if (dict.TryGetValue(key, out var v) && v != null)
            {
                if (v is int i) return i;
                if (int.TryParse(Convert.ToString(v), out var parsed)) return parsed;
            }
            return defaultValue;
        }

        private static bool GetDictBool(System.Collections.Generic.Dictionary<string, object> dict, string key, bool defaultValue)
        {
            if (dict.TryGetValue(key, out var v) && v != null)
            {
                var s = Convert.ToString(v)?.Trim().ToLowerInvariant();
                if (s == "true" || s == "yes" || s == "1") return true;
                if (s == "false" || s == "no" || s == "0") return false;
            }
            return defaultValue;
        }

        private static System.Collections.Generic.List<string> SavePdfAttachmentsToTempFiles(MimeMessage message)
        {
            var result = new System.Collections.Generic.List<string>();
            if (message == null) return result;

            foreach (var attachment in message.Attachments)
            {
                if (attachment is MimePart part)
                {
                    var mimeType = part.ContentType?.MimeType ?? string.Empty;
                    var fileName = part.FileName ?? string.Empty;
                    bool isPdf = mimeType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase)
                                 || fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
                    if (isPdf)
                    {
                        var tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
                        using (var stream = System.IO.File.Create(tempPath))
                        {
                            part.Content.DecodeTo(stream);
                        }
                        result.Add(tempPath);
                    }
                }
            }

            return result;
        }

        private static void PrintPdfFiles(string printer, params string[] pdfPaths)
        {
            if (pdfPaths == null || pdfPaths.Length == 0) return;

            foreach (var pdf in pdfPaths)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = pdf,
                        Verb = "printto",
                        UseShellExecute = true,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        Arguments = $"\"{printer}\""
                    };
                    Process.Start(psi);
                }
                catch (Exception ex)
                {
                    Logger.LogError($"Failed to print PDF '{pdf}'", ex);
                }
            }
        }
    }
}