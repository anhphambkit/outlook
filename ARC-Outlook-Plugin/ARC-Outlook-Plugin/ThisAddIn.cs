using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using System.CodeDom.Compiler;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Json;
using System.Linq.Expressions;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;
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
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;
//using redemption;

namespace ARC_Outlook_Plugin
{
    public partial class ThisAddIn
    {
        private Outlook.Application oApp;
        
        private Outlook._NameSpace oNS;
        
        private Outlook.MAPIFolder mailsFromThisFolder;
        
        private LinkedResource theContent;
        
        private string img;

        private List<string> listEntrySynced;

        private static System.Timers.Timer timerSync;

        private Items _items;
        
        private int _totalEventAdd = 0;

        private Store selectedStore;
        
        //private messageSuccess _messageSuccess;

        private dynamic resultLogin;

        private string tmpResult;

        private accountForm _formAccount;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.EventCreateNewNoteAfterEmailSent();
                this.ThreadCallCheckSync();
                this.StartupCheckLogin();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Start Error: " + ex.Message);
            }
        }

        public void StartupCheckLogin()
        {
            try
            {
                if ((Settings.Default.token == null ? true : Settings.Default.token == ""))
                {
                    this.showAccountForm(false);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("StartupCheckLogin: " + exception.Message);
            }
        }

        //public static void CheckProcessEmail()
        //{
        //    if ((Settings.Default.token == null ? false : Settings.Default.token != ""))
        //    {
        //        try
        //        {
        //            string @default = Settings.Default.host;
        //            string str = Settings.Default.email;
        //            string str1 = string.Concat(@default, "/api/mail/syncProcessEmail?email_user=", str);
        //            HttpClient httpClient = new HttpClient();
        //            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", string.Concat("= ", Settings.Default.token, "&&&Email=", Settings.Default.email));
        //            HttpResponseMessage result = httpClient.GetAsync(str1).Result;
        //            Task<string> task = result.Content.ReadAsStringAsync();
        //            dynamic obj = JObject.Parse(task.Result).data;
        //            if (obj != (dynamic)null)
        //            {
        //                foreach (dynamic obj1 in (IEnumerable)obj)
        //                {
        //                    dynamic obj2 = obj1.id;
        //                    dynamic obj3 = obj1.attachments;
        //                    Microsoft.Office.Interop.Outlook.Application application = Globals.ThisAddIn.Application;
        //                    RDOSession mAPIOBJECT = (RDOSession)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("29AB7A12-B531-450E-8F7A-EA94C2F3C05F")));
        //                    Store selectedStore = Globals.ThisAddIn.GetSelectedStore(Settings.Default.email);
        //                    mAPIOBJECT.MAPIOBJECT = selectedStore.Session.MAPIOBJECT;
        //                    RDOFolder defaultFolder = mAPIOBJECT.GetDefaultFolder(rdoDefaultFolders.olFolderSentMail);
        //                    RDOMail now = defaultFolder.Items.Add("IPM.Note");
        //                    now.Sent = true;
        //                    now.SentOn = DateTime.Now;
        //                    now.ReceivedTime = DateTime.Now;
        //                    now.Subject = (string)obj1.subject;
        //                    now.HTMLBody = (string)obj1.body;
        //                    now.To = (string)obj1.to_addr;
        //                    now.BCC = (string)obj1.bcc;
        //                    now.CC = (string)obj1.cc;
        //                    now.Recipients.Add(obj1.to_addr);
        //                    now.Recipients.ResolveAll(Type.Missing, Type.Missing);
        //                    now.SenderName = (string)obj1.from_name;
        //                    now.SenderEmailAddress = (string)obj1.from_addr;
        //                    dynamic obj4 = string.Concat(System.Windows.Forms.Application.LocalUserAppDataPath, "\\contentEmailSync\\mail_") + obj2;
        //                    if ((dynamic)(!Directory.Exists(obj4)))
        //                    {
        //                        Directory.CreateDirectory(obj4);
        //                    }
        //                    if (obj3 != (dynamic)null)
        //                    {
        //                        foreach (dynamic obj5 in (IEnumerable)obj3)
        //                        {
        //                            if (obj5 != (dynamic)null)
        //                            {
        //                                string str2 = ((string)obj5.path_string).Replace("http://localhost", @default);
        //                                now.Attachments.Add(str2, Type.Missing, Type.Missing, Type.Missing);
        //                            }
        //                        }
        //                    }
        //                    now.Save();
        //                    string str3 = string.Concat(@default, "/api/mail/updateStatusEmailSync");
        //                    HttpClient authenticationHeaderValue = new HttpClient();
        //                    authenticationHeaderValue.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", string.Concat("= ", Settings.Default.token, "&&&Email=", Settings.Default.email));
        //                    JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
        //                    jsonObject.Add("email_id", obj1.id.ToString());
        //                    StringContent stringContent = new StringContent(jsonObject.ToString(), Encoding.UTF8, "application/json");
        //                    HttpResponseMessage httpResponseMessage = authenticationHeaderValue.PostAsync(str3, stringContent).Result;
        //                }
        //            }
        //        }
        //        catch (Exception exception)
        //        {
        //            MessageBox.Show(exception.Message);
        //        }
        //    }
        //}

        public void EventCreateNewNoteAfterEmailSent()
        {
            try
            {
                this.RemoveEventAfterEmailSent();
                Store store = this.GetSelectedStore(Settings.Default.email);
                //Folder folder = store.GetDefaultFolder(OlDefaultFolders.olFolderSentMail) as Folder;
                //this._items = folder.Items;
                //bool flag = Settings.Default.token != null && Settings.Default.token != "";
                //if (flag)
                //{
                //    //new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ItemsEvents_Event), "ItemAdd").AddEventHandler(this._items, new ItemsEvents_ItemAddEventHandler(this, (UIntPtr)ldftn(Items_ItemAdd)));
                //    this._totalEventAdd++;
                //}
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
                this._formAccount = new accountForm();
                bool flag = Settings.Default.token != null && Settings.Default.token != "";
                if (flag)
                {
                    this._formAccount.showLogout();
                }
                else
                {
                    this._formAccount.showLogin();
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
                foreach (Store store in this.Application.GetNamespace("MAPI").Stores)
                {
                    if (store.DisplayName == account)
                    {
                        this.selectedStore = store;
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
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
                FileInfo[] files = directoryInfo.GetFiles();
                for (int i = 0; i < (int)files.Length; i++)
                {
                    files[i].Delete();
                }
                DirectoryInfo[] directories = directoryInfo.GetDirectories();
                for (int j = 0; j < (int)directories.Length; j++)
                {
                    directories[j].Delete(true);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
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

        public async void SendRequestLogin(JObject objData)
        {
            try
            {
                await Task.Run<string>(() => this.tmpResult = ThisAddIn.LoginAction(objData));
                if (this.tmpResult != "fail")
                {
                    this._formAccount.closeForm();
                    //new messageSuccess().ShowDialog();
                }
                else
                {
                    this._formAccount.showError();
                }
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

        public void TriggerLogoutBtn(string host, string email)
        {
            try
            {
                this.ResetDefaultInfo(host, email);
                //this._formAccount = new accountForm();
                //this._formAccount.ShowDialog();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public static string LoginAction(JObject objData)
        {
            string thisAddIn;
            dynamic obj = JObject.Parse(objData.ToString());
            string str = (string)(obj["host"].ToString() + "/api/mail/test");
            JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
            jsonObject.Add("email", obj["email"].ToString());
            jsonObject.Add("password", obj["password"].ToString());
            HttpClient httpClient = new HttpClient();
            StringContent stringContent = new StringContent(jsonObject.ToString(), Encoding.UTF8, "application/json");
            Task<HttpResponseMessage> task = httpClient.GetAsync(str);
            HttpResponseMessage result = task.Result;
            AggregateException exception = task.Exception;
            TaskStatus status = task.Status;
            if (!result.IsSuccessStatusCode)
            {
                thisAddIn = "fail";
            }
            else
            {
                Task<string> task1 = result.Content.ReadAsStringAsync();
                if ((task1.Result == null || !(task1.Result != "") ? true : task1.Result == "\"\""))
                {
                    thisAddIn = "fail";
                }
                else
                {
                    dynamic obj1 = JObject.Parse(task1.Result);
                    Globals.ThisAddIn.resultLogin = obj1.data;
                    thisAddIn = (string)((dynamic)Globals.ThisAddIn.resultLogin).ToString();
                }
            }
            if (thisAddIn == "fail")
            {
                Globals.ThisAddIn.ResetDefaultInfo(obj["host"].ToString(), obj["email"].ToString());
            }
            else
            {
                Globals.ThisAddIn.SaveYourInfo(obj["host"].ToString(), obj["email"].ToString(), ((dynamic)Globals.ThisAddIn.resultLogin).ToString());
            }
            Globals.ThisAddIn.EventCreateNewNoteAfterEmailSent();
            return thisAddIn;
        }

        //public static object ParseDataFromEmail(MailItem mail, int indexMail)
        //{
        //    string str = DateTime.Now.ToString("yyyyMMddHHmmss");
        //    string str1 = string.Concat(System.Windows.Forms.Application.LocalUserAppDataPath, "\\contentEmailTemp\\mail_", str);
        //    if (!Directory.Exists(str1))
        //    {
        //        Directory.CreateDirectory(str1);
        //    }
        //    string[] strArrays = new string[] { ".png", ".jpg", ".gif", ".bmp" };
        //    string hTMLBody = null;
        //    string str2 = string.Concat(new object[] { str1, "\\", str, "-attachments-", indexMail });
        //    string str3 = string.Concat(new object[] { str, "-attachments-", indexMail, ".zip" });
        //    string str4 = string.Concat(str1, "\\", str3);
        //    string[] allReciptents = new string[0];
        //    allReciptents = ThisAddIn.GetAllReciptents(mail);
        //    hTMLBody = mail.HTMLBody;
        //    foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mail.Attachments)
        //    {
        //        string base64String = null;
        //        string str5 = string.Concat(str, "_", attachment.FileName);
        //        string str6 = Path.GetExtension(string.Concat(str1, "\\", str5)).Replace(".", "");
        //        string str7 = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";
        //        attachment.SaveAsFile(string.Concat(str1, "\\", str5));
        //        string property = (string)((dynamic)attachment.PropertyAccessor.GetProperty(str7));
        //        try
        //        {
        //            if (strArrays.Any<string>(new Func<string, bool>(attachment.FileName.Contains)))
        //            {
        //                using (Image image = Image.FromFile(string.Concat(str1, "\\", str5)))
        //                {
        //                    using (MemoryStream memoryStream = new MemoryStream())
        //                    {
        //                        image.Save(memoryStream, image.RawFormat);
        //                        base64String = Convert.ToBase64String(memoryStream.ToArray());
        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception exception)
        //        {
        //            MessageBox.Show(exception.Message);
        //        }
        //        try
        //        {
        //            if (property == "")
        //            {
        //                if (!Directory.Exists(str2))
        //                {
        //                    Directory.CreateDirectory(str2);
        //                }
        //                File.Move(string.Concat(str1, "\\", str5), string.Concat(str2, "\\", str5));
        //            }
        //            else if (hTMLBody.ToLower().Contains(string.Concat("cid:", property.ToLower())))
        //            {
        //                hTMLBody = hTMLBody.Replace(string.Concat("\"cid:", property, "\""), string.Concat("data:image/", str6, ";base64,", base64String));
        //            }
        //        }
        //        catch (Exception exception1)
        //        {
        //            MessageBox.Show(exception1.Message);
        //        }
        //    }
        //    JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
        //    jsonObject.Add("body", hTMLBody);
        //    jsonObject.Add("entry_id", mail.EntryID);
        //    if (Directory.Exists(str2))
        //    {
        //        ZipFile.CreateFromDirectory(str2, str4);
        //        jsonObject.Add("pathAttachment", str4);
        //        jsonObject.Add("fileName", str3);
        //    }
        //    return jsonObject;
        //}

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
