using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

internal class Program
{
    private static int Main(string[] args)
    {
        try
        {
            LogText("SmartManMailSender.exe started.");
            string argumentDescriptions = @"Example arguments: 
                SmartManMailSender.exe 
                outlook 
                -s ""My title"" 
                -f ""C:\\Desktop\\test1.pdf"" 
                -t ""template.txt"" 
                -r frank@gmail.com 
                -p placeholder1,placeholder2, ""placeholder that has comma,""... 
                
                OR
                
                SmartManMailSender.exe 
                smtp 
                -host smtp.gmail.com
                -port 587
                -ssl true
                -username yourusername@gmail.com
                -password yourpassword
                -sender_name ""志元資訊/Frank""
                -sender_email frank@gmail.com
                -receiver_name ""Receiver/Frank""
                -receiver_email frank@gmail.com
                -s ""My title"" 
                -f ""C:\\Desktop\\test1.pdf"" 
                -t ""template.txt"" 
                -p placeholder1,placeholder2, ""placeholder that has comma,""
            ";
            if (args.Length < 2 || (args[0].ToLower() != "smtp" && args[0].ToLower() != "outlook"))
            {
                LogError(@$"Invalid arguments. {argumentDescriptions}");
                LogText("Program exited.");
                return -1;
            }

            var emailMethod = args[0].ToLower();
            var arguments = ParseArguments(args.Skip(1).ToArray());

            // Validate required arguments
            if (emailMethod == "outlook" && (!arguments.ContainsKey("-s")  || !arguments.ContainsKey("-t") || !arguments.ContainsKey("-r")))
            {
                LogError($"Missing required arguments for Outlook method. {argumentDescriptions}");
                LogText("Program exited.");
                return -1;
            }
            else if (emailMethod == "smtp" && (
                         !arguments.ContainsKey("-host") ||
                         !arguments.ContainsKey("-port") ||
                         !arguments.ContainsKey("-username") ||
                         !arguments.ContainsKey("-password") ||
                         !arguments.ContainsKey("-sender_email") ||
                         !arguments.ContainsKey("-receiver_email") ||
                         !arguments.ContainsKey("-s") ||  !arguments.ContainsKey("-t")))
            {
                LogError($"Missing required arguments for SMTP method. {argumentDescriptions}");
                LogText("Program exited.");
                return -1;
            }

            // Extract common arguments
            var subject = arguments["-s"];
            var filePath = arguments.ContainsKey("-f") ? arguments["-f"] : null;
            var templatePath = arguments["-t"];
            var placeholders = arguments.ContainsKey("-p") ? arguments["-p"].Split(',') : new string[0];

            // get full path of a file
            if (filePath != null && !Path.IsPathRooted(filePath))
            {
                filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filePath);
            }

            // Validate file paths
            if (filePath != null && !File.Exists(filePath))
            {
                LogError($"File: {filePath} does not exist.");
                LogText("Program exited.");
                return -1;
            }

            if (!File.Exists(templatePath))
            {
                LogError($"Email template file does not exist: {templatePath}");
                LogText("Program exited.");
                return -1;
            }

            // Read and process email content
            string emailContent = File.ReadAllText(templatePath);
            for (int i = 0; i < placeholders.Length; i++)
            {
                emailContent = emailContent.Replace($"[Placeholder{i + 1}]", placeholders[i]);
            }

            // Send email based on the chosen method
            if (emailMethod == "outlook")
            {
                SendWithOutlook(subject: subject, filePath: filePath, emailContent: emailContent, recipient: arguments["-r"]);
            }
            else // smtp
            {
                SendWithSmtp(host: arguments["-host"],
                             port: int.TryParse(arguments["-port"], out int port) ? port : 587,
                             useSsl: arguments.ContainsKey("-ssl") ? bool.TryParse(arguments["-ssl"], out bool ssl) && ssl : false,
                             username: arguments["-username"],
                             password: arguments["-password"],
                             senderName: arguments.ContainsKey("-sender_name") ? arguments["-sender_name"] : arguments["-sender_email"],
                             senderEmail: arguments["-sender_email"],
                             receiverName: arguments.ContainsKey("-receiver_name") ? arguments["-receiver_name"] : arguments["-receiver_email"],
                             receiverEmail: arguments["-receiver_email"],
                             subject: subject,
                             filePath: filePath,
                             emailContent: emailContent);
            }

            LogText("Email sent successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            LogError($"An unexpected error occurred. [Message]: {ex.Message}, [StackTrace]: {ex.StackTrace}");
            LogText("Program exited");
            return -1;
        }
    }

    private static void SendWithOutlook(string subject, string? filePath, string emailContent, string recipient)
    {
        try
        {
            Outlook.Application outlookApp = new Outlook.ApplicationClass();
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.HTMLBody = emailContent;
            if (filePath != null)
            {
                LogText($"File path: {filePath}");
                mailItem.Attachments.Add(filePath);
            }
            mailItem.Recipients.Add(recipient);
            //mailItem.SendUsingAccount = outlookApp.Session.Accounts.Cast<Outlook.Account>().FirstOrDefault(a => a.SmtpAddress.Equals(senderEmail, StringComparison.OrdinalIgnoreCase));
            LogText($"Sending to {recipient} using Outlook...");
            mailItem.Send();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    private static void SendWithSmtp(string host, int port, bool useSsl, string username, string password, string senderName, string senderEmail, string receiverName, string receiverEmail, string subject, string? filePath, string emailContent)
    {
        try
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(senderName, senderEmail));
            message.To.Add(new MailboxAddress(receiverName, receiverEmail));
            message.Subject = subject;

            var builder = new BodyBuilder();
            builder.HtmlBody = emailContent;
            if (filePath != null)
            {
                builder.Attachments.Add(filePath);
            }

            message.Body = builder.ToMessageBody();
           
            LogText($"Sending from {senderEmail} to {receiverEmail} using SMTP...");

            using (var client = new SmtpClient())
            {
                client.Connect(host, port, useSsl);
                client.Authenticate(username, password);
                string smtpResponse = client.Send(message);
                LogText($"SMTP response: {smtpResponse}");
                client.Disconnect(true);
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public static void LogText(string message)
    {
        string formattedMessage = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss} --- {message}";
        Console.WriteLine(formattedMessage);

        string logDirectory = "./logs";
        string logFileName = $"mail_log_{DateTime.Now:yyyyMMdd}.txt";
        string logPath = Path.Combine(logDirectory, logFileName);

        string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logDirectory);

        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        using (var logFile = new StreamWriter(logPath, true))
        {
            logFile.WriteLine(formattedMessage);
        }
    }

    public static void LogError(string message, Exception? ex = null)
    {
        string formattedMessage = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss} --- [Error] {message}";
        if (ex != null)
        {
            formattedMessage += $" [Error Message] {ex.Message}, [StackTrace] {ex.StackTrace}";
        }
        Console.WriteLine(formattedMessage);

        string logFileName = $"mail_log_{DateTime.Now:yyyyMMdd}.txt";
        string logDirectory = "./logs";
        string logPath = Path.Combine(logDirectory, logFileName);

        string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logDirectory);

        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        using (var logFile = new StreamWriter(logPath, true))
        {
            logFile.WriteLine(formattedMessage);
        }
    }

    private static Dictionary<string, string> ParseArguments(string[] args)
    {
        var result = new Dictionary<string, string>();
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].StartsWith("-"))
            {
                if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                {
                    result[args[i]] = args[i + 1];
                    i++;
                }
                else
                {
                    result[args[i]] = "";
                }
            }
        }
        return result;
    }
}