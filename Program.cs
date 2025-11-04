// To do 
// Move contacts that do not come back
// May want to make this a tool that poeple can use

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Interop.Outlook;

public class Recruiter
{
    public Recruiter(string email, bool sent = true)
    {
        Email = email;
        Sent = sent;
    }

    public string Email { get; set; } = string.Empty;
    public bool Sent { get; set; }
}

namespace SendEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            MAPIFolder contactsFolder = null;
            MAPIFolder inbox = null;
            NameSpace ns = null;
            Application outlookApp = null;
            List<string> emails = new List<string>();
            List<string> log = new List<string>();

            try
            {
                bool bProcessFailures = true;
                bool bSendEmails = false;
                bool bSendEmailsUsingCategory = false;

                outlookApp = new Application();
                ns = outlookApp.GetNamespace("MAPI");
                ns.Logon("", "", Missing.Value, Missing.Value);

                if (bProcessFailures)
                {
                    int failureCount = 0;
                    string filePathEmails = @"C:\Users\RobertAmenn\Desktop\emails.txt";
                    string filePathLog = @"C:\Users\RobertAmenn\Desktop\logs.txt";

                    inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                    DateTime targetDate = new DateTime(2025, 11, 4);

                    string filter = $"[ReceivedTime] >= '{targetDate:MM/dd/yyyy 00:00}' AND [ReceivedTime] < '{targetDate.AddDays(1):MM/dd/yyyy 00:00}'";
                    Items filteredItems = inbox.Items.Restrict(filter);

                    List<MailItem> mailItemsToDelete = new List<MailItem>();
                    List<ReportItem> reportItemsToDelete = new List<ReportItem>();

                    foreach (object item in filteredItems)
                    {
                        failureCount++;
                        log.Add($"{failureCount}");
                        if (item is ReportItem report)
                        {
                            log.Add($"\nSubject: {report.Subject}");
                            List<string> emailsFromEmail = ExtractEmails(report.Body);

                            if (emailsFromEmail.Count == 0)
                            {
                                log.Add("(No emails found in body)");
                            }
                            else
                            {
                                foreach (string e in emailsFromEmail)
                                {
                                    if (string.Equals(e, "robert@albionnet.com", StringComparison.OrdinalIgnoreCase) || 
                                        e.ToLower().Contains("namprd18.prod.outlook.com".ToLower()))
 
                                    {
                                        log.Add($"Skipping email {e}");
                                        continue;
                                    }

                                    emails.Add(e);
                                    log.Add(e);
                                }
                            }

                            reportItemsToDelete.Add(report);
                        }
                        else if (item is MailItem mail)
                        {
                            Console.WriteLine($"\nSubject: {mail.Subject}");
                            List<string> emailsFromEmail = ExtractEmails(mail.Body);

                            if (emailsFromEmail.Count == 0)
                            {
                                log.Add("(No emails found in body)");
                            }
                            else
                            {
                                foreach (string e in emailsFromEmail)
                                {
                                    if (string.Equals(e, "robert@albionnet.com", StringComparison.OrdinalIgnoreCase) || 
                                        e.ToLower().Contains("namprd18.prod.outlook.com".ToLower()))
                                    {
                                        log.Add($"Skipping email {e}");
                                        continue;
                                    }

                                    emails.Add(e);
                                    log.Add(e);
                                }
                            }

                            mailItemsToDelete.Add(mail);
                        }
                        else
                        {
                            Marshal.ReleaseComObject(item);
                        }
                    }

                    using (StreamWriter writer = new StreamWriter(filePathEmails, false)) // false = overwrite
                    {
                        foreach (string e in emails)
                        {
                            writer.WriteLine($"{e}");
                        }
                    }

                    using (StreamWriter writer = new StreamWriter(filePathLog, false)) // false = overwrite
                    {
                        foreach (string l in log)
                        {
                            writer.WriteLine($"{l}");
                        }
                    }

                    foreach (ReportItem r in reportItemsToDelete)
                    {
                        r.Delete();
                        Thread.Sleep(100);
                    }

                    foreach (MailItem m in mailItemsToDelete)
                    {
                        m.Delete();
                        Thread.Sleep(100);
                    }

                }

                if (bSendEmails)
                {
                    string filePath = @"C:\Users\RobertAmenn\Desktop\recruiters.txt";

                    Dictionary<string, Recruiter> recruiters = LoadRecruitersIntoDictionary(filePath);

                    if (bSendEmailsUsingCategory)
                    { 
                        int contactCount = 0;
                        contactsFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                        string filterContacts = $"[Categories] = 'R8'";
    
                        Items filteredContacts = contactsFolder.Items.Restrict(filterContacts);

                        foreach (object item in filteredContacts)
                        {
                            contactCount++;
                            Console.WriteLine(contactCount);

                            if (item is ContactItem contact)
                            {
                                if (string.IsNullOrWhiteSpace(contact.Email1Address) == false)
                                {
                                    if (recruiters.TryGetValue(contact.Email1Address, out Recruiter r) == true)
                                    {
                                        if (r.Sent == false)
                                        {
                                            r.Sent = SendMailTo(outlookApp, r.Email);
                                        }
                                    }
                                    else
                                    {
                                        bool sent = SendMailTo(outlookApp, contact.Email1Address);
                                        Recruiter r1 = new Recruiter(contact.Email1Address, sent);
                                        if (recruiters.ContainsKey(contact.Email1Address) == false)
                                        {
                                            recruiters.Add(contact.Email1Address, r1);
                                        }
                                        else
                                        {
                                            Console.WriteLine(contact.Email1Address);
                                        }
                                    }
                                }

                                if (string.IsNullOrWhiteSpace(contact.Email2Address) == false)
                                {
                                    if (recruiters.TryGetValue(contact.Email2Address, out Recruiter r) == true)
                                    {
                                        if (r.Sent == false)
                                        {
                                            r.Sent = SendMailTo(outlookApp, r.Email);
                                        }
                                    }
                                    else
                                    {
                                        bool sent = SendMailTo(outlookApp, contact.Email2Address);
                                        Recruiter r2 = new Recruiter(contact.Email2Address, sent);
                                        if (recruiters.ContainsKey(contact.Email2Address) == false)
                                        {
                                            recruiters.Add(contact.Email2Address, r2);
                                        }
                                        else
                                        {
                                            Console.WriteLine(contact.Email2Address);
                                        }
                                    }
                                }

                                if (string.IsNullOrWhiteSpace(contact.Email3Address) == false)
                                {
                                    /* Add this check   
                                    e.ToLower().Contains("aol.com".ToLower()) ||
 e.ToLower().Contains("yahoo.com".ToLower()) ||
 e.ToLower().Contains("gmail.com".ToLower()) ||
 e.ToLower().Contains("hotmail.com".ToLower()) ||
 e.ToLower().Contains("amazon.com".ToLower()) ||
 e.ToLower().Contains("google.com".ToLower()) ||
 e.ToLower().Contains("microsoft.com".ToLower()) ||
 e.ToLower().Contains("verizon.net".ToLower())) */


                                    if (recruiters.TryGetValue(contact.Email3Address, out Recruiter r) == true)
                                    {
                                        if (r.Sent == false)
                                        {
                                            r.Sent = SendMailTo(outlookApp, r.Email);
                                        }
                                    }
                                    else
                                    {
                                        bool sent = SendMailTo(outlookApp, contact.Email3Address);
                                        Recruiter r3 = new Recruiter(contact.Email3Address, sent);
                                        if (recruiters.ContainsKey(contact.Email3Address) == false)
                                        {
                                            recruiters.Add(contact.Email3Address, r3);
                                        }
                                        else
                                        {
                                            Console.WriteLine(contact.Email3Address);
                                        }
                                    }
                                }
                            }

                            Marshal.ReleaseComObject(item);
                        }
                    }
                    else
                    {
                        foreach (Recruiter r in recruiters.Values)
                        {
                            if(r.Sent == false)
                            {
                                bool sent = SendMailTo(outlookApp, r.Email);
                                r.Sent = sent;

                            }
                        }
                    }
                
                    using (StreamWriter writer = new StreamWriter(filePath, false)) // false = overwrite
                    {
                        foreach (Recruiter r in recruiters.Values)
                        {
                            writer.WriteLine($"{r.Email},{r.Sent}");
                        }
                    }
                }

                ns.Logoff();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                if(contactsFolder != null)
                {
                    Marshal.ReleaseComObject(contactsFolder);
                }

                if (inbox != null)
                {
                    Marshal.ReleaseComObject(inbox);
                }

                Marshal.ReleaseComObject(ns);
                Marshal.ReleaseComObject(outlookApp);
            }
        }

        private static List<string> ExtractEmails(string text)
        {
            string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            MatchCollection matches = Regex.Matches(text, pattern);

            List<string> result = new List<string>();
            foreach (Match match in matches)
            {
                if (!result.Contains(match.Value))
                    result.Add(match.Value);
            }

            return result;
        }

        private static bool SendMailTo(Application outlookApp, string emailAddress)
        {
            try
            {
                MailItem mail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                string toAddress = emailAddress;
                string subject = "Here is my updated resume";
                string resumePath = @"D:\Resaech\Files\Robert Amenn SDE -Latest.docx";

                mail.To = toAddress;
                mail.Subject = subject;

                mail.HTMLBody = @"
<html>
  <body style='font-family:Segoe UI, Tahoma, Geneva, Verdana, sans-serif; font-size:14px; color:#000000;'>
    <p>I am a <strong>Full Stack Developer</strong> with extensive experience building applications across multiple cloud platforms — Azure, AWS, and GCP — with a strong focus on AI solutions. My recent work includes:</p>

    <p><strong>Chatbot Solutions for Insurance Customers and Customer Service Teams</strong></p>
    <ul>
      <li>Designed and deployed two AI-powered chatbots: one serving insurance customers and another supporting the customer service team.</li>
      <li>Leveraged <strong>Azure OpenAI</strong> to improve question categorization, ensuring inquiries were accurately routed within the <strong>Azure Bot Framework</strong>.</li>
      <li>Developed and integrated <strong>Retrieval-Augmented Generation (RAG)</strong> using <strong>Azure Cognitive Search</strong> and <strong>Azure OpenAI</strong>, enhancing response accuracy for prescription drug coverage, medical procedures, and general customer service information.</li>
      <li>Implemented drug classification using Azure OpenAI to determine coverage eligibility.</li>
      <li>Utilized <strong>Azure Document Intelligence</strong> to process PDFs for inclusion in <strong>Azure Cognitive Search</strong> indexes.</li>
      <li>Built and maintained indexes using semantic and vector search to optimize information retrieval across large datasets.</li>
    </ul>

    <p><strong>Additional Expertise:</strong></p>
    <ul>
      <li>Extensive experience developing RESTful APIs and SDKs using C#.</li>
      <li>Experience with Node.js and React.js applications.</li>
      <li>Proficiency with <strong>Microsoft Copilot Studio</strong>, <strong>Azure AI Studio</strong>, and <strong>Azure AI Foundry</strong> to accelerate AI development.</li>
      <li>Hands-on experience with LLM evaluation, prompt engineering, and RAG implementations.</li>
      <li>Strong familiarity with Agile and Scrum methodologies.</li>
      <li>Knowledge of OCR, text classification, semantic search, and PDF parsing libraries.</li>
    </ul>

    <p>I am eager to bring my skills and experience to new opportunities where I can deliver meaningful impact.</p>

    <p>Thank you for your consideration.</p>

    <p>Best regards,<br>
    <strong>Robert Amenn</strong><br>
    📞 360.421.5119<br>
    ✉️ <a href='mailto:robert@albionnet.com'>robert@albionnet.com</a></p>
  </body>
</html>
";
                mail.Attachments.Add(resumePath, OlAttachmentType.olByValue, Type.Missing, Type.Missing);

                // Save it to Drafts (does NOT send)
                mail.Save();

                Console.WriteLine("✅ Draft email created successfully in Outlook Drafts folder.");

                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        private static List<Recruiter> LoadRecruiters(string filePath)
        {
            List<Recruiter> recruiters = new List<Recruiter>();

            if (File.Exists(filePath))
            {
                foreach (string line in File.ReadLines(filePath))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split(',');

                    if (parts.Length >= 1)
                    {
                        string email = parts[0].Trim();

                        bool sent = false;

                        if (parts.Length >= 2)
                        {
                            bool.TryParse(parts[1].Trim(), out sent);
                        }

                        recruiters.Add(new Recruiter(email, sent));
                    }
                    else
                    {
                        throw new System.Exception($"Loaded {recruiters.Count} recruiters from file.");
                    }
                }
            }
            else
            {
                throw new System.Exception($"File not found: {filePath}");
            }

            return recruiters;
        }

        private static Dictionary<string, Recruiter> LoadRecruitersIntoDictionary(string filePath)
        {
            Dictionary<string, Recruiter> recruiters = new Dictionary<string, Recruiter>();

            if (File.Exists(filePath))
            {
                foreach (string line in File.ReadLines(filePath))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split(',');

                    if (parts.Length >= 1)
                    {
                        string email = parts[0].Trim();

                        bool sent = false;

                        if (parts.Length >= 2)
                        {
                            bool.TryParse(parts[1].Trim(), out sent);
                        }

                        recruiters.Add(email, new Recruiter(email, sent));
                    }
                    else
                    {
                        throw new System.Exception($"Loaded {recruiters.Count} recruiters from file.");
                    }
                }
            }
            else
            {
                throw new System.Exception($"File not found: {filePath}");
            }

            return recruiters;
        }
    }
}