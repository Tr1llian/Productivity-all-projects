using System;
using System.Windows.Forms;
using DataGridView_Import_Excel;
using GroupwareTypeLibrary;
using Application = GroupwareTypeLibrary.Application;

namespace Productivity
{
    public partial class LoginGW : Form
    {
        
        public LoginGW()
        {
            InitializeComponent();
            Password.PasswordChar = '*';

        }

      

        private void LoginButton_Click(object sender, EventArgs e)
        {
            Application gwapplication = new Application();
            try
            {
                string login = Login.Text.ToString();
                string password = Password.Text.ToString();
                Account objAccount = gwapplication.Login(login, null, password , LoginConstants.egwNeverPrompt, null);
                Form GW = new Form2( login, password);
                _ = GW.ShowDialog();
                Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Невірний логін або пароль", "Помилка логування", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(ex.ToString());
                
                if (ex.Message.ToString() == "Неправильный пароль")
                {
                    Console.WriteLine("Те що треба:)");
                }
            }
        }

        private void Password_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = LoginButton;
        }

        private void Login_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
