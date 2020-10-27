using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using GroupwareTypeLibrary;
using Productivity;
using Application = GroupwareTypeLibrary.Application;
using Message = GroupwareTypeLibrary.Message;

namespace DataGridView_Import_Excel
{
    public partial class Form2 : Form
    {
        string path = "C:/test/";
        string pathfile = @"C:/test/recipients.txt";
        string pathfileXls = @"C:/test/productivity.xlsx";
        string Login = "";
        string Password = "";

        List<string> mailboxes;

        public Form2(string login, string password)
        {
            //string path = "C:/test/";
            CreateIfMissing(path);
            //string pathfile = @"C:/test/recipients.txt";
            CheckTextFile(pathfile).ToString();
  
            mailboxes = File.ReadAllLines(pathfile).ToList();
            
            InitializeComponent();
            Login = login;
            Password = password;
            
            MailBoxList.DataSource = mailboxes;
        }

        private void CreateIfMissing(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private void RewriteTextFile(string path)
        {


            if (!File.Exists(path))
            {
                File.Create(path).Dispose();
                using (StreamWriter tw = new StreamWriter(path, false))
                {//second parameter is `Append` and false means override content            
                    //tw.WriteLine();
                    tw.Close();
                }

            }

            else if (File.Exists(path))
            {
                using (StreamWriter tw = new StreamWriter(path, false))
                {//second parameter is `Append` and false means override content            
                    //tw.WriteLine(textBox1.Text.ToString());
                    tw.Close();
                }
            }
        }
        private int CheckTextFile(string path)
        {
            int size = 0;

            if (!File.Exists(path))
            {
                File.Create(path).Dispose();
                using (TextWriter tw = new StreamWriter(path))
                {
                    //tw.WriteLine("50");
                    tw.Close();
                }
                size = 50;
            }
            else
            {
                size = 0;
            }
            return size;

        }

        private void CreateMail() 
        {
           
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = SentMail;
        }

        private void Login_label_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            mailboxes.Remove(MailBoxList.SelectedItem.ToString());
            MailBoxList.DataSource = null;
            MailBoxList.DataSource = mailboxes;
            File.WriteAllText(pathfile, string.Empty);
            File.AppendAllLines(pathfile, mailboxes);
        }

        private void SentMail_Click(object sender, EventArgs e)
        {
           
            try
            {
               
                Application gwapplication = new Application();
                Account objAccount = gwapplication.Login(Login, null, Password, LoginConstants.egwAllowPasswordPrompt, null);
                Messages messages1 = objAccount.MailBox.Messages;
                Messages messages = messages1;
                Message message = messages.Add("GW.MESSAGE.MAIL", "Draft", null);
                Recipients recipients = message.Recipients;
                Recipient recipient;
                foreach (string rec in mailboxes)
                {
                   recipient = recipients.Add(rec, null, null);
                }

                _ = message.Attachments.Add(pathfileXls);
                message.Subject.PlainText = "Звіт продуктивності";
                message.BodyText.PlainText = "Звіт продуктивності у додатку";

                Message myMessage = message.Send();
                File.Delete(pathfileXls);
                MyMessages m = new MyMessages("Успішно", "Звіт успішно надісланий", 1);
                m.ShowDialog();

                bool okButtonClicked = m.OKButtonClicked;
                //MessageBox.Show("Звіт успішно відправлений...", "Звіт успішно відправлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                if (ex.Message.ToString() == "Неправильный пароль")
                {
                    Console.WriteLine("Те що треба:)");
                }
            }
        }

        private void AddMail_Click(object sender, EventArgs e)
        {
            if (MailTextBox.Text.ToString() == "")
            {
                Console.WriteLine("empty text box");
            }
            else
            {
                try
                {
                    var eMailValidator = new System.Net.Mail.MailAddress(MailTextBox.Text.ToString());
                    File.AppendAllText(pathfile, MailTextBox.Text + Environment.NewLine);
                    mailboxes.Add(MailTextBox.Text.ToString());
                    MailBoxList.DataSource = null;
                    MailBoxList.DataSource = mailboxes;
                    MailTextBox.Text = "";
                    
                }
                catch (FormatException ex)
                {
                    // wrong e-mail address
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        private void MailTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = AddMail;
        }
    }
}
