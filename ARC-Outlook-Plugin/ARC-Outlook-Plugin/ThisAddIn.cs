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
using System.Net;
using Redemption;

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

        public static void CheckProcessEmail()
        {
            if ((Settings.Default.token == null ? false : Settings.Default.token != ""))
            {
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                    string hostDefault = Settings.Default.host;
                    string emailDefault = Settings.Default.email;
                    string urlSync = string.Concat(hostDefault, "/api/mail/syncProcessEmail?email_user=", emailDefault);
                    HttpClient httpClient = new HttpClient();
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", string.Concat("= ", Settings.Default.token, "&&&Email=", Settings.Default.email));
                    HttpResponseMessage result = httpClient.GetAsync(urlSync).Result;
                    Task<string> task = result.Content.ReadAsStringAsync();
                    if ((task.Result == null || !(task.Result != "") ? true : task.Result == "\"\""))
                    {
                        MessageBox.Show("Process Email Fail: Result Null From Server");
                    }
                    else
                    {
                        dynamic emailSyncs = JObject.Parse(task.Result)["data"];
                        if (emailSyncs != (dynamic)null)
                        {
                            Outlook.Application application = Globals.ThisAddIn.Application;
                            foreach (dynamic emailSync in (IEnumerable)emailSyncs)
                            {
                                dynamic idEmail = emailSync.id;
                                dynamic attachmentEmails = emailSync.attachments;
                                RDOSession mAPIOBJECT = new RDOSession();
                                Store selectedStore = Globals.ThisAddIn.GetSelectedStore(Settings.Default.email);
                                mAPIOBJECT.MAPIOBJECT = selectedStore.Session.MAPIOBJECT;
                                RDOFolder defaultFolder = mAPIOBJECT.GetFolderFromPath(selectedStore.GetDefaultFolder(OlDefaultFolders.olFolderSentMail).FullFolderPath);
                                RDOMail now = defaultFolder.Items.Add("IPM.Note");
                                now.Sent = true;
                                now.SentOn = DateTime.Now;
                                now.ReceivedTime = DateTime.Now;
                                now.Subject = emailSync.subject.ToString();
                                now.HTMLBody = emailSync.body.ToString();
                                now.To = emailSync.to_addr.ToString();
                                now.BCC = emailSync.bcc.ToString();
                                now.CC = emailSync.cc.ToString();
                                //now.Recipients.Add(emailSync.to_addr.ToString());
                                //now.Recipients.ResolveAll();
                                now.SenderName = emailSync.from_name.ToString();
                                now.SenderEmailAddress = emailSync.from_addr.ToString();
                                dynamic folderPath = string.Concat(System.Windows.Forms.Application.LocalUserAppDataPath, "\\contentEmailSync\\mail_") + idEmail;
                                if ((dynamic)(!Directory.Exists(folderPath)))
                                {
                                    Directory.CreateDirectory(folderPath);
                                }
                                if (attachmentEmails != (dynamic)null)
                                {
                                    foreach (dynamic attachmentEmail in (IEnumerable)attachmentEmails)
                                    {
                                        if (attachmentEmail != (dynamic)null)
                                        {
                                            string urlAttachment = ((string)attachmentEmail.path_string).Replace("http://localhost", hostDefault);
                                            now.Attachments.Add(urlAttachment);
                                        }
                                    }
                                }
                                now.Save();
                                string urlUpdateStatusEmailSync = string.Concat(hostDefault, "/api/mail/updateStatusEmailSync");
                                HttpClient authenticationHeaderValue = new HttpClient();
                                authenticationHeaderValue.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", string.Concat("= ", Settings.Default.token, "&&&Email=", Settings.Default.email));
                                JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
                                jsonObject.Add("email_id", emailSync.id.ToString());
                                StringContent stringContent = new StringContent(jsonObject.ToString(), Encoding.UTF8, "application/json");
                                HttpResponseMessage httpResponseMessage = authenticationHeaderValue.PostAsync(urlUpdateStatusEmailSync, stringContent).Result;
                            }
                        }
                    }
                   
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Process Email Fail: " + exception.Message);
                }
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
                Thread thread = new Thread(new ThreadStart(ThisAddIn.CheckProcessEmail));
                thread.Start();
                thread.Join();
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

        //public static void UploadNewNoteToServer(object objectMail)
        //{
        //    try
        //    {
        //        string host = Settings.Default.host;
        //        MailItem mailItem = objectMail as MailItem;
        //        object obj = ThisAddIn.ParseDataFromEmail(mailItem, 0);
        //        object arg = JObject.Parse(obj.ToString());
        //        Dictionary<string, string> dictionary = new Dictionary<string, string>();
        //        if (ThisAddIn.<> o__29.<> p__2 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__2 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        Func<CallSite, object, bool> target = ThisAddIn.<> o__29.<> p__2.Target;
        //        CallSite<> p__ = ThisAddIn.<> o__29.<> p__2;
        //        if (ThisAddIn.<> o__29.<> p__1 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__1 = CallSite<Func<CallSite, object, object, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.Constant, null)
        //            }));
        //        }
        //        Func<CallSite, object, object, object> target2 = ThisAddIn.<> o__29.<> p__1.Target;
        //        CallSite<> p__2 = ThisAddIn.<> o__29.<> p__1;
        //        if (ThisAddIn.<> o__29.<> p__0 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__0 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "fileName", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        bool flag = target(<> p__, target2(<> p__2, ThisAddIn.<> o__29.<> p__0.Target(ThisAddIn.<> o__29.<> p__0, arg), null));
        //        if (flag)
        //        {
        //            if (ThisAddIn.<> o__29.<> p__6 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__6 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof(string), typeof(ThisAddIn)));
        //            }
        //            Func<CallSite, object, string> target3 = ThisAddIn.<> o__29.<> p__6.Target;
        //            CallSite<> p__3 = ThisAddIn.<> o__29.<> p__6;
        //            if (ThisAddIn.<> o__29.<> p__5 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__5 = CallSite<Func<CallSite, string, object, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.Add, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Func<CallSite, string, object, object> target4 = ThisAddIn.<> o__29.<> p__5.Target;
        //            CallSite<> p__4 = ThisAddIn.<> o__29.<> p__5;
        //            string arg2 = host + "/api/mail/uploadAttachmentFromClient?entry_id=";
        //            if (ThisAddIn.<> o__29.<> p__4 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__4 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Func<CallSite, object, object> target5 = ThisAddIn.<> o__29.<> p__4.Target;
        //            CallSite<> p__5 = ThisAddIn.<> o__29.<> p__4;
        //            if (ThisAddIn.<> o__29.<> p__3 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__3 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "entry_id", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            string text = target3(<> p__3, target4(<> p__4, arg2, target5(<> p__5, ThisAddIn.<> o__29.<> p__3.Target(ThisAddIn.<> o__29.<> p__3, arg))));
        //            if (ThisAddIn.<> o__29.<> p__9 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__9 = CallSite<Action<CallSite, Dictionary<string, string>, string, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Add", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Action<CallSite, Dictionary<string, string>, string, object> target6 = ThisAddIn.<> o__29.<> p__9.Target;
        //            CallSite<> p__6 = ThisAddIn.<> o__29.<> p__9;
        //            Dictionary<string, string> arg3 = dictionary;
        //            string arg4 = "attachment";
        //            if (ThisAddIn.<> o__29.<> p__8 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__8 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Func<CallSite, object, object> target7 = ThisAddIn.<> o__29.<> p__8.Target;
        //            CallSite<> p__7 = ThisAddIn.<> o__29.<> p__8;
        //            if (ThisAddIn.<> o__29.<> p__7 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__7 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "pathAttachment", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            target6(<> p__6, arg3, arg4, target7(<> p__7, ThisAddIn.<> o__29.<> p__7.Target(ThisAddIn.<> o__29.<> p__7, arg)));
        //            if (ThisAddIn.<> o__29.<> p__12 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__12 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof(string), typeof(ThisAddIn)));
        //            }
        //            Func<CallSite, object, string> target8 = ThisAddIn.<> o__29.<> p__12.Target;
        //            CallSite<> p__8 = ThisAddIn.<> o__29.<> p__12;
        //            if (ThisAddIn.<> o__29.<> p__11 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__11 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Func<CallSite, object, object> target9 = ThisAddIn.<> o__29.<> p__11.Target;
        //            CallSite<> p__9 = ThisAddIn.<> o__29.<> p__11;
        //            if (ThisAddIn.<> o__29.<> p__10 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__10 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "fileName", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            string fileName = target8(<> p__8, target9(<> p__9, ThisAddIn.<> o__29.<> p__10.Target(ThisAddIn.<> o__29.<> p__10, arg)));
        //            HttpClient httpClient = new HttpClient();
        //            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", "= " + Settings.Default.token + "&&&Email=" + Settings.Default.email);
        //            httpClient.BaseAddress = new Uri(text);
        //            httpClient.DefaultRequestHeaders.Accept.Clear();
        //            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        //            MultipartFormDataContent multipartFormDataContent = new MultipartFormDataContent();
        //            HttpContent content = new StringContent("fileToUpload");
        //            HttpContent content2 = new FormUrlEncodedContent(dictionary);
        //            multipartFormDataContent.Add(content, "fileToUpload");
        //            multipartFormDataContent.Add(content2, "medicineOrder");
        //            if (ThisAddIn.<> o__29.<> p__15 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__15 = CallSite<Func<CallSite, Type, object, FileMode, FileStream>>.Create(Binder.InvokeConstructor(CSharpBinderFlags.None, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.IsStaticType, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, null)
        //                }));
        //            }
        //            Func<CallSite, Type, object, FileMode, FileStream> target10 = ThisAddIn.<> o__29.<> p__15.Target;
        //            CallSite<> p__10 = ThisAddIn.<> o__29.<> p__15;
        //            Type typeFromHandle = typeof(FileStream);
        //            if (ThisAddIn.<> o__29.<> p__14 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__14 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Func<CallSite, object, object> target11 = ThisAddIn.<> o__29.<> p__14.Target;
        //            CallSite<> p__11 = ThisAddIn.<> o__29.<> p__14;
        //            if (ThisAddIn.<> o__29.<> p__13 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__13 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "pathAttachment", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            FileStream content3 = target10(<> p__10, typeFromHandle, target11(<> p__11, ThisAddIn.<> o__29.<> p__13.Target(ThisAddIn.<> o__29.<> p__13, arg)), FileMode.Open);
        //            multipartFormDataContent.Add(new StreamContent(content3)
        //            {
        //                Headers =
        //                {
        //                    ContentDisposition = new ContentDispositionHeaderValue("form-data")
        //                    {
        //                        Name = "file",
        //                        FileName = fileName
        //                    }
        //                }
        //            });
        //            HttpResponseMessage httpResponseMessage = null;
        //            try
        //            {
        //                httpResponseMessage = httpClient.PostAsync(text, multipartFormDataContent).Result;
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show(ex.Message);
        //            }
        //            string result = httpResponseMessage.Content.ReadAsStringAsync().Result;
        //        }
        //        JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
        //        object obj2 = new JsonArray(new JsonValue[0]);
        //        Recipients recipients = mailItem.Recipients;
        //        foreach (object obj3 in recipients)
        //        {
        //            Recipient recipient = (Recipient)obj3;
        //            if (ThisAddIn.<> o__29.<> p__16 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__16 = CallSite<Action<CallSite, object, string>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Add", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null)
        //                }));
        //            }
        //            ThisAddIn.<> o__29.<> p__16.Target(ThisAddIn.<> o__29.<> p__16, obj2, recipient.Address);
        //        }
        //        if (ThisAddIn.<> o__29.<> p__17 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__17 = CallSite<Action<CallSite, JsonObject, string, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Add", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        ThisAddIn.<> o__29.<> p__17.Target(ThisAddIn.<> o__29.<> p__17, jsonObject, "addresses", obj2);
        //        jsonObject.Add("subject", mailItem.Subject);
        //        jsonObject.Add("fromEmail", mailItem.SenderEmailAddress);
        //        if (ThisAddIn.<> o__29.<> p__20 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__20 = CallSite<Action<CallSite, JsonObject, string, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Add", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        Action<CallSite, JsonObject, string, object> target12 = ThisAddIn.<> o__29.<> p__20.Target;
        //        CallSite<> p__12 = ThisAddIn.<> o__29.<> p__20;
        //        JsonObject arg5 = jsonObject;
        //        string arg6 = "contentMail";
        //        if (ThisAddIn.<> o__29.<> p__19 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__19 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        Func<CallSite, object, object> target13 = ThisAddIn.<> o__29.<> p__19.Target;
        //        CallSite<> p__13 = ThisAddIn.<> o__29.<> p__19;
        //        if (ThisAddIn.<> o__29.<> p__18 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__18 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "body", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        target12(<> p__12, arg5, arg6, target13(<> p__13, ThisAddIn.<> o__29.<> p__18.Target(ThisAddIn.<> o__29.<> p__18, arg)));
        //        if (ThisAddIn.<> o__29.<> p__23 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__23 = CallSite<Action<CallSite, JsonObject, string, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Add", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        Action<CallSite, JsonObject, string, object> target14 = ThisAddIn.<> o__29.<> p__23.Target;
        //        CallSite<> p__14 = ThisAddIn.<> o__29.<> p__23;
        //        JsonObject arg7 = jsonObject;
        //        string arg8 = "entry_id";
        //        if (ThisAddIn.<> o__29.<> p__22 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__22 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        Func<CallSite, object, object> target15 = ThisAddIn.<> o__29.<> p__22.Target;
        //        CallSite<> p__15 = ThisAddIn.<> o__29.<> p__22;
        //        if (ThisAddIn.<> o__29.<> p__21 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__21 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "entry_id", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        target14(<> p__14, arg7, arg8, target15(<> p__15, ThisAddIn.<> o__29.<> p__21.Target(ThisAddIn.<> o__29.<> p__21, arg)));
        //        if (ThisAddIn.<> o__29.<> p__26 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__26 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        Func<CallSite, object, bool> target16 = ThisAddIn.<> o__29.<> p__26.Target;
        //        CallSite<> p__16 = ThisAddIn.<> o__29.<> p__26;
        //        if (ThisAddIn.<> o__29.<> p__25 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__25 = CallSite<Func<CallSite, object, object, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null),
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.Constant, null)
        //            }));
        //        }
        //        Func<CallSite, object, object, object> target17 = ThisAddIn.<> o__29.<> p__25.Target;
        //        CallSite<> p__17 = ThisAddIn.<> o__29.<> p__25;
        //        if (ThisAddIn.<> o__29.<> p__24 == null)
        //        {
        //            ThisAddIn.<> o__29.<> p__24 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "fileName", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //            {
        //                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //            }));
        //        }
        //        bool flag2 = target16(<> p__16, target17(<> p__17, ThisAddIn.<> o__29.<> p__24.Target(ThisAddIn.<> o__29.<> p__24, arg), null));
        //        if (flag2)
        //        {
        //            if (ThisAddIn.<> o__29.<> p__29 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__29 = CallSite<Action<CallSite, JsonObject, string, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.ResultDiscarded, "Add", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, null),
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Action<CallSite, JsonObject, string, object> target18 = ThisAddIn.<> o__29.<> p__29.Target;
        //            CallSite<> p__18 = ThisAddIn.<> o__29.<> p__29;
        //            JsonObject arg9 = jsonObject;
        //            string arg10 = "attachment";
        //            if (ThisAddIn.<> o__29.<> p__28 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__28 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", null, typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            Func<CallSite, object, object> target19 = ThisAddIn.<> o__29.<> p__28.Target;
        //            CallSite<> p__19 = ThisAddIn.<> o__29.<> p__28;
        //            if (ThisAddIn.<> o__29.<> p__27 == null)
        //            {
        //                ThisAddIn.<> o__29.<> p__27 = CallSite<Func<CallSite, object, object>>.Create(Binder.GetMember(CSharpBinderFlags.None, "fileName", typeof(ThisAddIn), new CSharpArgumentInfo[]
        //                {
        //                    CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null)
        //                }));
        //            }
        //            target18(<> p__18, arg9, arg10, target19(<> p__19, ThisAddIn.<> o__29.<> p__27.Target(ThisAddIn.<> o__29.<> p__27, arg)));
        //        }
        //        else
        //        {
        //            jsonObject.Add("attachment", null);
        //        }
        //        string requestUri = host + "/api/mail/saveNoteFromEmailDataClient";
        //        HttpClient httpClient2 = new HttpClient();
        //        httpClient2.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", "= " + Settings.Default.token + "&&&Email=" + Settings.Default.email);
        //        StringContent content4 = new StringContent(jsonObject.ToString(), Encoding.UTF8, "application/json");
        //        HttpResponseMessage result2 = httpClient2.PostAsync(requestUri, content4).Result;
        //    }
        //    catch (Exception ex2)
        //    {
        //        MessageBox.Show(ex2.Message);
        //    }
        //}

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
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            string thisAddIn;
            dynamic obj = JObject.Parse(objData.ToString());
            string str = (string)(obj["host"].ToString() + "/api/mail/loginFromOutlook");
            JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
            jsonObject.Add("email", obj["email"].ToString());
            jsonObject.Add("password", obj["password"].ToString());
            HttpClient httpClient = new HttpClient();
            StringContent stringContent = new StringContent(jsonObject.ToString(), Encoding.UTF8, "application/json");
            Task<HttpResponseMessage> task = httpClient.PostAsync(str, stringContent);
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
