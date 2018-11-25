using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Json;
using System.Windows.Forms;
using ARC_Outlook_Plugin.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;

namespace ARC_Outlook_Plugin
{
    public partial class accountForm : Form
    {
        public accountForm()
        {
            this.InitializeComponent();
            Outlook.Accounts informationAcounts = Globals.ThisAddIn.GetInformationAcounts();
            foreach (object obj in informationAcounts)
            {
                Outlook.Account account = (Outlook.Account)obj;
                try
                {
                    this.emailListSelect.Items.Add(account.DisplayName);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            try
            {
                string text = Globals.ThisAddIn.CheckDefaultLogin();
                bool flag = text != null;
                if (flag)
                {
                    this.emailListSelect.SelectedIndex = this.emailListSelect.Items.IndexOf(text);
                }
                else
                {
                    this.emailListSelect.SelectedIndex = 0;
                }
                this.hostInput.Text = Settings.Default.host;
                this.errorMessage.Hide();
                //this.loadingImage.Hide();
                this.logoutPanel.Hide();
                this.hostData.Text = Settings.Default.host;
                this.emailData.Text = Settings.Default.email;
            }
            catch (System.Exception ex2)
            {
                MessageBox.Show(ex2.Message);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //string text = this.emailListSelect.Text;
            //string text2 = this.passwordInput.Text;
        }
        
        private void cancelLoginForm_Click(object sender, EventArgs e)
        {

        }

        public void closeForm()
        {
            try
            {
                base.Invoke(new MethodInvoker(delegate ()
                {
                    base.Close();
                }));
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        public void showLogin()
        {
            try
            {
                bool flag = System.Windows.Forms.Application.OpenForms["accountForm"] is accountForm;
                if (!flag)
                {
                    base.Show();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        public void showError()
        {
            try
            {
                bool flag = System.Windows.Forms.Application.OpenForms["accountForm"] is accountForm;
                if (flag)
                {
                    base.Invoke(new MethodInvoker(delegate ()
                    {
                        //this.loadingImage.Hide();
                        //this.actionLoginPanel.Show();
                        //this.errorMessage.Show();
                    }));
                }
                else
                {
                    //this.loadingImage.Hide();
                    //this.actionLoginPanel.Show();
                    //this.errorMessage.Show();
                    //base.Show();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        public void showLogout()
        {
            try
            {
                bool flag = System.Windows.Forms.Application.OpenForms["accountForm"] is accountForm;
                if (!flag)
                {
                    //this.formLoginPanel.Hide();
                    //this.logoutPanel.Show();
                    base.Show();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void cancelLogoutBtn_Click(object sender, EventArgs e)
        {

        }
        
        private void logoutBtn_Click(object sender, EventArgs e)
        {

        }

        private void infoPlugin_Click(object sender, EventArgs e)
        {

        }

        private void loginBtn_Click(object sender, EventArgs e)
        {
            try
            {
                bool flag = this.hostInput.Text.Contains("http://") || this.hostInput.Text.Contains("https://");
                if (flag)
                {
                    this.errorMessage.Hide();
                    dynamic jsonObject = new JObject();
                    jsonObject.host = this.hostInput.Text;
                    jsonObject.email = this.emailListSelect.Text;
                    jsonObject.password = this.passwordInput.Text;
                    //this.actionLoginPanel.Hide();
                    //this.loadingImage.Show();
                    Globals.ThisAddIn.SendRequestLogin(jsonObject);
                }
                else
                {
                    this.errorMessage.Text = "Host servers must be a full path, example: http://example.com";
                    this.errorMessage.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
