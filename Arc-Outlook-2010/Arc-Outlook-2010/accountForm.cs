﻿using System;
using System.Windows.Forms;
using Arc_Outlook_2010.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;

namespace Arc_Outlook_2010
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
                this.loadingImage.Hide();
                this.logoutPanel.Hide();
                this.hostData.Text = Settings.Default.host;
                this.emailData.Text = Settings.Default.email;
            }
            catch (System.Exception ex2)
            {
                MessageBox.Show(ex2.Message);
            }
        }

        private void cancelLoginForm_Click(object sender, EventArgs e)
        {
            try
            {
                base.Invoke(new MethodInvoker(delegate ()
                {
                    base.Close();
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cancel login fail: " + ex.Message);
            }
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
                MessageBox.Show("Close form error: " + ex.Message);
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
                MessageBox.Show("Show login fail: " + ex.Message);
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
                        this.loadingImage.Hide();
                        this.actionLoginPanel.Show();
                        this.errorMessage.Show();
                    }));
                }
                else
                {
                    this.loadingImage.Hide();
                    this.actionLoginPanel.Show();
                    this.errorMessage.Show();
                    base.Show();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Show error fail: " + ex.Message);
            }
        }

        public void showLogout()
        {
            try
            {
                bool flag = System.Windows.Forms.Application.OpenForms["accountForm"] is accountForm;
                if (!flag)
                {
                    this.formLoginPanel.Hide();
                    this.logoutPanel.Show();
                    base.Show();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Show Logout Form Error: " + ex.Message);
            }
        }

        private void cancelLogoutBtn_Click(object sender, EventArgs e)
        {
            try
            {
                base.Invoke(new MethodInvoker(delegate ()
                {
                    base.Close();
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Close form logout error: " + ex.Message);
            }
        }

        private void logoutBtn_Click(object sender, EventArgs e)
        {
            try
            {
                base.Invoke(new MethodInvoker(delegate ()
                {
                    base.Close();
                    Globals.ThisAddIn.TriggerLogoutBtn(this.hostData.Text, this.emailData.Text);
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Logout Error: " + ex.Message);
            }
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
                    this.actionLoginPanel.Hide();
                    this.loadingImage.Show();
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
                MessageBox.Show("Login error: " + ex.Message);
            }
        }
    }
}
