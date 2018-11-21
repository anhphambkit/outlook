using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
//using System.Json;
using System.Linq;
using System.Linq.Expressions;
//using System.Net.Http;
//using System.Net.Http.Headers;
using System.Net.Mail;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using ARC_Outlook_Plugin.Properties;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Outlook;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Exception = System.Exception;
//using Newtonsoft.Json.Linq;
//using Redemption;

namespace ARC_Outlook_Plugin
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Interop.Outlook.Application oApp;
        
        private Microsoft.Office.Interop.Outlook._NameSpace oNS;
        
        private Microsoft.Office.Interop.Outlook.MAPIFolder mailsFromThisFolder;
        
        private LinkedResource theContent;
        
        private string img;

        private List<string> listEntrySynced;

        private static System.Timers.Timer timerSync;

        private Items _items;
        
        private int _totalEventAdd = 0;

        private Store selectedStore;

        //private accountForm _formAccount;
       
        //private messageSuccess _messageSuccess;
        
        private dynamic resultLogin;

        private string tmpResult;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.EventCreateNewNoteAfterEmailSent();
                this.ThreadCallCheckSync();
                //this.StartupCheckLogin();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void EventCreateNewNoteAfterEmailSent()
        {
            try
            {
                this.RemoveEventAfterEmailSent();
                Store store = this.GetSelectedStore(Settings.Default.email);
                Folder folder = store.GetDefaultFolder(OlDefaultFolders.olFolderSentMail) as Folder;
                this._items = folder.Items;
                bool flag = Settings.Default.token != null && Settings.Default.token != "";
                if (flag)
                {
                    //new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ItemsEvents_Event), "ItemAdd").AddEventHandler(this._items, new ItemsEvents_ItemAddEventHandler(this, (UIntPtr)ldftn(Items_ItemAdd)));
                    this._totalEventAdd++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void showAccountForm(bool error)
        {
            try
            {
                //this._formAccount = new accountForm();
                bool flag = Settings.Default.token != null && Settings.Default.token != "";
                if (flag)
                {
                    //this._formAccount.ShowLogout();
                }
                else
                {
                    //this._formAccount.showLogin();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void RemoveEventAfterEmailSent()
        {
            for (int i = 0; i < this._totalEventAdd; i++)
            {
                try
                {
                    //new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ItemsEvents_Event), "ItemAdd").RemoveEventHandler(this._items, new ItemsEvents_ItemAddEventHandler(this, (UIntPtr)ldftn(Items_ItemAdd)));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            this._totalEventAdd = 0;
        }

        public void ShowSuccessMessage(string message)
        {
            try
            {
                //this._messageSuccess = new messageSuccess();
                //this._messageSuccess.showMessage(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ThreadCallCheckSync()
        {
            try
            {
                ThisAddIn.timerSync = new System.Timers.Timer();
                ThisAddIn.timerSync.Interval = 180000.0;
                ThisAddIn.timerSync.Elapsed += ThisAddIn.ThreadAutoSyncProcessMail;
                ThisAddIn.timerSync.AutoReset = true;
                ThisAddIn.timerSync.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static void ThreadAutoSyncProcessMail(object source, ElapsedEventArgs e)
        {
            try
            {
                //Thread thread = new Thread(new ThreadStart(ThisAddIn.CheckProcessEmail));
                //thread.Start();
                //thread.Join();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public Store GetSelectedStore(string account)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application application = this.Application;
                Microsoft.Office.Interop.Outlook._NameSpace @namespace = application.GetNamespace("MAPI");
                Stores stores = @namespace.Stores;
                foreach (object obj in stores)
                {
                    Store store = (Store)obj;
                    bool flag = store.DisplayName == account;
                    if (flag)
                    {
                        this.selectedStore = store;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return this.selectedStore;
        }

        private void Items_ItemAdd(object item)
        {
            try
            {
                MailItem mailItem = item as MailItem;
                string[] array = new string[0];
                array = ThisAddIn.GetAllReciptents(mailItem);
                bool flag = (!mailItem.Subject.StartsWith("FW: ") || !mailItem.Subject.StartsWith("RE: ")) && array.Length != 0;
                if (flag)
                {
                    try
                    {
                        Thread thread = new Thread(new ParameterizedThreadStart(ThisAddIn.PrepairDataToRequest));
                        thread.Start(item);
                        thread.Join();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
            }
        }

        public void DeleteAllTemp()
        {
            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(System.Windows.Forms.Application.LocalUserAppDataPath);
                foreach (FileInfo fileInfo in directoryInfo.GetFiles())
                {
                    fileInfo.Delete();
                }
                foreach (DirectoryInfo directoryInfo2 in directoryInfo.GetDirectories())
                {
                    directoryInfo2.Delete(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            try
            {
                Globals.ThisAddIn.DeleteAllTemp();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void PrepairDataToRequest(object item)
        {
            try
            {
                //Thread thread = new Thread(new ParameterizedThreadStart(ThisAddIn.UploadNewNoteToServer));
                //thread.Start(item);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public async void SendRequestLogin(object objData)
        {
            try
            {
                //await Task.Run<string>(() => this.tmpResult = ThisAddIn.LoginAction(objData));
                //if (this.tmpResult != "fail")
                //{
                //    this._formAccount.closeForm();
                //    new messageSuccess().ShowDialog();
                //}
                //else
                //{
                //    this._formAccount.showError();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ResetDefaultInfo(string host, string email)
        {
            try
            {
                Settings.Default.token = null;
                Settings.Default.tmpHost = host;
                Settings.Default.tmpEmail = email;
                Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static string[] GetAllReciptents(MailItem mail)
        {
            List<string> list = new List<string>();
            Recipients recipients = mail.Recipients;
            foreach (object obj in recipients)
            {
                Recipient recipient = (Recipient)obj;
                try
                {
                    bool flag = recipient.Address != null && recipient.Address != "";
                    if (flag)
                    {
                        list.Add(recipient.Address);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return list.ToArray();
        }

        public Accounts GetInformationAcounts()
        {
            return this.Application.Session.Accounts;
        }

        public void SaveYourInfo(string host, string email, string token)
        {
            try
            {
                Settings.Default.host = host;
                Settings.Default.email = email;
                Settings.Default.token = token;
                Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string CheckDefaultLogin()
        {
            bool flag = Settings.Default.email != null;
            string result;
            if (flag)
            {
                result = Settings.Default.email;
            }
            else
            {
                result = null;
            }
            return result;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
