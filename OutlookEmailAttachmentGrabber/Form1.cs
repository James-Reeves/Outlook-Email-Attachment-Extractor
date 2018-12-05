using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailAttachmentGrabber
{
    public partial class Form1 : Form
    {

        static string basePath = @"C:\Emails\";
        static int totalfilesize = 0;

        public Form1()
        {
            InitializeComponent();
            Console.SetOut(new ControlWriter(textBox3));
            MessageBox.Show("Click the \"Load Email\" button to load your email handle into the program.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1.ActiveForm.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EnumerateAccounts();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExtractAttachments();
        }

        public class ControlWriter : TextWriter
        {
            private Control textbox;
            public ControlWriter(Control textbox)
            {
                this.textbox = textbox;
            }

            public override void Write(char value)
            {
                textbox.Text += value;
            }

            public override void Write(string value)
            {
                textbox.Text += value;
            }

            public override Encoding Encoding
            {
                get { return Encoding.ASCII; }
            }
        }

        public static void EnumerateFoldersInDefaultStore()
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root);
        }

        // Uses recursion to enumerate Outlook subfolders.
        public static void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        // Write the folder path.
                        Console.WriteLine(childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        EnumerateFolders(childFolder);
                    }
                }
            }
            Console.WriteLine("Checking in " + folder.FolderPath);
            IterateMessages(folder);
        }

        public static void IterateMessages(Outlook.Folder folder)
        {
            // attachment extensions to save. Since this is only for my HIN return files I will limit this to .txt only
            string[] extensionsArray = { ".txt" };

            // Iterate through all items ("messages") in a folder
            var fi = folder.Items;
            if (fi != null)
            {

                try
                {
                    foreach (Object item in fi)
                    {
                        Outlook.MailItem mi = (Outlook.MailItem)item;
                        var attachments = mi.Attachments;
                        if (attachments.Count != 0)
                        {

                            // Create a directory to store the attachment 
                            if (!Directory.Exists(basePath + folder.FolderPath))
                            {
                                Directory.CreateDirectory(basePath + folder.FolderPath);
                            }

                            //Console.WriteLine(mi.Sender.Address);
                            //Console.WriteLine(mi.Subject + " [" + attachments.Count + "]");
                            //Console.WriteLine(generateFolder(folder.FolderPath, mi.Sender.Address));
                            for (int i = 1; i <= mi.Attachments.Count; i++)
                            {
                                var fn = mi.Attachments[i].FileName.ToLower();
                                //check wither any of the strings in the extensionsArray are contained within the filename
                                if (extensionsArray.Any(fn.Contains))
                                {

                                    // Create a further sub-folder for the sender
                                    if (!Directory.Exists(basePath + folder.FolderPath + @"\" + mi.Sender.Address))
                                    {
                                        Directory.CreateDirectory(basePath + folder.FolderPath + @"\" + mi.Sender.Address);
                                    }
                                    totalfilesize = totalfilesize + mi.Attachments[i].Size;
                                    if (!File.Exists(basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[i].FileName))
                                    {
                                        Console.WriteLine("Saving " + mi.Attachments[i].FileName);
                                        mi.Attachments[i].SaveAsFile(basePath + folder.FolderPath + @"\" + mi.Sender.Address + @"\" + mi.Attachments[i].FileName);
                                        //mi.Attachments[i].Delete();
                                    }
                                    else
                                    {
                                        Console.WriteLine("Already saved " + mi.Attachments[i].FileName);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine("An error occurred: '{0}'", e);
                }
            }
        }

        // Retrieves the email address for a given account object
        public static string EnumerateAccountEmailAddress(Outlook.Account account)
        {
            try
            {
                if (string.IsNullOrEmpty(account.SmtpAddress) || string.IsNullOrEmpty(account.UserName))
                {
                    Outlook.AddressEntry oAE = account.CurrentUser.AddressEntry as Outlook.AddressEntry;
                    if (oAE.Type == "EX")
                    {
                        Outlook.ExchangeUser oEU = oAE.GetExchangeUser() as Outlook.ExchangeUser;
                        return oEU.PrimarySmtpAddress;
                    }
                    else
                    {
                        return oAE.Address;
                    }
                }
                else
                {
                    return account.SmtpAddress;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }

        static void EnumerateAccounts()
        {
            Console.WriteLine("Outlook Attachment Extraction Tool");
            Console.WriteLine("---------------------------------");
            Console.WriteLine("The Inquisitive Analyst");
            Console.WriteLine("Created by: James Reeves");
            Console.WriteLine();
            int id;
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;

            id = 1;
            foreach (Outlook.Account account in accounts)
            {
                Console.WriteLine("Run: " + EnumerateAccountEmailAddress(account));
                id++;
            }
            Console.WriteLine("Quit: Quit Application");
            Console.WriteLine();
            Console.WriteLine("Select Run or Quit from the provided box, then click the \"Save .txt Attachments\" button.");
        }

        public void ExtractAttachments()
        {
            string response = "";
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;

            response = UsersInput();

            if (response == "Q")
            {
                Console.WriteLine("Quitting...");
                Form1.ActiveForm.Close();
            }
            if (response != "")
            {
            
                    Console.WriteLine("Processing: " + accounts[Int32.Parse(response.Trim())].DisplayName);
                    Console.WriteLine("Processing: " + EnumerateAccountEmailAddress(accounts[Int32.Parse(response.Trim())]));

                    Outlook.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                    selectedFolder = GetFolder(@"\\" + accounts[Int32.Parse(response)].DisplayName);
                    EnumerateFolders(selectedFolder);
                    MessageBox.Show("Job Successfully Completed!" + "\n" + "Your files have been saved here:" +
                        "\n" + @"C:\Emails\");
            }
            else
            {
                Console.WriteLine("Invalid Account Selected");
            }
            
        }

        public string UsersInput()
        {
            string userInput = "";

            if (comboBox1.Text == "Quit")
            {
                userInput = "Q";
            }
            if (comboBox1.Text == "Run")
            {
                userInput = "1";
            }
            return userInput;
        }

        // Returns Folder object based on folder path
        static Outlook.Folder GetFolder(string folderPath)
        {
            //MessageBox.Show("Looking for: " + folderPath);
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray());
                Outlook.Application Application = new Outlook.Application();
                folder = Application.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}