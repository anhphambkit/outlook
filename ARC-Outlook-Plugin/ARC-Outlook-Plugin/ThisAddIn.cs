using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using ARC_Outlook_Plugin.Properties;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;
using System.Net;
using Redemption;
using System.Dynamic;
using Newtonsoft.Json;

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

        private messageSuccess _messageSuccess;

        private dynamic resultLogin;

        private string tmpResult;

        private accountForm _formAccount;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.StartupCheckLogin();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Start Error: " + ex.Message);
            }
        }

        // Function check user logged or not:
        public void StartupCheckLogin()
        {
            try
            {
                if ((Settings.Default.token == null ? true : Settings.Default.token == ""))
                {
                    this.showAccountForm(false);
                }
                else
                {
                    this.StartAutoThreads();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("StartupCheckLogin: " + exception.Message);
            }
        }

        // Auto start threads:
        public void StartAutoThreads()
        {
            this.EventCreateNewNoteAfterEmailSent();
            this.ThreadCallCheckSync();
        }

        // Function synce note from server to email in Sent Mail Folder of outlook client:
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
                    HttpClient httpSyncEmailClient = new HttpClient();
                    httpSyncEmailClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", string.Concat("= ", Settings.Default.token, "&&&Email=", Settings.Default.email));
                    HttpResponseMessage httpResponseSyncEmail = httpSyncEmailClient.GetAsync(urlSync).Result;
                    Task<string> resultSyncEmail = httpResponseSyncEmail.Content.ReadAsStringAsync();
                    if ((resultSyncEmail.Result == null || !(resultSyncEmail.Result != "") ? true : resultSyncEmail.Result == "\"\""))
                    {
                        MessageBox.Show("Process Email Fail: Result Null From Server");
                    }
                    else
                    {
                        dynamic emailSyncs = JObject.Parse(resultSyncEmail.Result)["data"];
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
                                JsonObject dataUpdateStatusEmail = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
                                dataUpdateStatusEmail.Add("email_id", emailSync.id.ToString());
                                StringContent stringContent = new StringContent(dataUpdateStatusEmail.ToString(), Encoding.UTF8, "application/json");
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

        // Function listen event create new note after email sent:
        public void EventCreateNewNoteAfterEmailSent()
        {
            try
            {
                this.RemoveEventAfterEmailSent();
                Store store = this.GetSelectedStore(Settings.Default.email);
                Folder folder = store.GetDefaultFolder(OlDefaultFolders.olFolderSentMail) as Folder;
                _items = folder.Items;
                bool flag = Settings.Default.token != null && Settings.Default.token != "";
                if (flag)
                {
                    _items.ItemAdd += new ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
                    _totalEventAdd++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot create event create new note after email sent: " + ex.Message);
            }
        }

        // Show form login if user not login | or show logout form if user logged
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
                MessageBox.Show("Error show form account: " + ex.Message);
            }
        }

        public void StartSyncEmailNow()
        {
            ThisAddIn.CheckProcessEmail();
        }

        // Remove event after email sent (prevent duplicate event)
        public void RemoveEventAfterEmailSent()
        {
            for (int i = 0; i < this._totalEventAdd; i++)
            {
                try
                {
                    _items.ItemAdd -= new ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error remove event after email sent: " + ex.Message);
                }
            }
            this._totalEventAdd = 0;
        }

        // Function show success message when loggin successfully
        public void ShowSuccessMessage(string message)
        {
            try
            {
                this._messageSuccess = new messageSuccess();
                this._messageSuccess.showMessage(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Thread (background task check sync process mail from note (server) to email sent (client)
        public void ThreadCallCheckSync()
        {
            try
            {
                ThisAddIn.timerSync = new System.Timers.Timer();
                ThisAddIn.timerSync.Interval = 180000.0; // Default 3 minutes
                ThisAddIn.timerSync.Elapsed += ThisAddIn.ThreadAutoSyncProcessMail;
                ThisAddIn.timerSync.AutoReset = true;
                ThisAddIn.timerSync.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error create thread check sync process email: " + ex.Message);
            }
        }

        // Thread auto sync process email:
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
                MessageBox.Show("Error start thread auto sync process email: " + ex.Message);
            }
        }

        // Get selected store (folder email) from email:
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
                MessageBox.Show("Error get selected store: " + exception.Message);
            }
            return this.selectedStore;
        }

        // Function action event when email sent:
        private void Items_ItemAdd(object item)
        {
            try
            {
                MailItem mailItem = item as MailItem;
                string[] reciptents = new string[0];
                reciptents = ThisAddIn.GetAllReciptents(mailItem);
                bool flag = (!mailItem.Subject.StartsWith("FW: ") || !mailItem.Subject.StartsWith("RE: ")) && reciptents.Length != 0;
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
                        MessageBox.Show("Error thread prepair data cannot start: " + ex.Message);
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error action event when email sent: " + exception.Message);
            }
        }

        // Function delete all temp data on local:
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
                MessageBox.Show("Error delete temp data: " + exception.Message);
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

        // Prepair data to send request
        public static void PrepairDataToRequest(object item)
        {
            try
            {
                Thread thread = new Thread(new ParameterizedThreadStart(ThisAddIn.UploadNewNoteToServer));
                thread.Start(item);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error prepair data: " + ex.Message);
            }
        }

        // Function create new note to server:
        public static void UploadNewNoteToServer(object objectMail)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                string host = Settings.Default.host;
                MailItem mailItem = objectMail as MailItem;
                object dataEmail = ThisAddIn.ParseDataFromEmail(mailItem);
                dynamic data = JObject.Parse(dataEmail.ToString());
                if (data.fileName != null) // Check if has attachment, upload it to server
                {
                    HttpClient uploadAttachmentClient = new HttpClient();
                    uploadAttachmentClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", "= " + Settings.Default.token + "&&&Email=" + Settings.Default.email);
                    uploadAttachmentClient.BaseAddress = new Uri(host);
                    uploadAttachmentClient.DefaultRequestHeaders.Accept.Clear();
                    uploadAttachmentClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    MultipartFormDataContent form = new MultipartFormDataContent();
                    HttpContent content = new StringContent("fileToUpload");
                    Dictionary<string, string> postDataUploadAttachment = new Dictionary<string, string>();
                    postDataUploadAttachment.Add("entry_id", data.entry_id.ToString());

                    HttpContent DictionaryItems = new FormUrlEncodedContent(postDataUploadAttachment);
                    form.Add(content, "fileToUpload");
                    form.Add(DictionaryItems, "medicineOrder");

                    var stream = new FileStream(data.pathAttachment.ToString(), FileMode.Open);
                    content = new StreamContent(stream);
                    content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
                    {
                        Name = "file",
                        FileName = data.fileName.ToString()
                    };
                    form.Add(content);

                    HttpResponseMessage httpResponseUploadMessage = null;

                    string urlUploadAttachment = host + "/api/mail/uploadAttachmentFromClient";
                    try
                    {
                        httpResponseUploadMessage = (uploadAttachmentClient.PostAsync(urlUploadAttachment, form)).Result;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Upload Attachments Error: " + ex.Message);
                    }
                    
                    string resultUpload = httpResponseUploadMessage.Content.ReadAsStringAsync().Result;
                }

                // Create new note:
                string newNoteApi = host + "/api/mail/saveNoteFromEmailDataClient";
                HttpClient createNewNoteHttpClient = new HttpClient();
                createNewNoteHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("BearerOutlook", "= " + Settings.Default.token + "&&&Email=" + Settings.Default.email);
                StringContent dataCreateNote = new StringContent(dataEmail.ToString(), Encoding.UTF8, "application/json");
                HttpResponseMessage httpResponseCreateNewMessage = null;
                try
                {
                    httpResponseCreateNewMessage = createNewNoteHttpClient.PostAsync(newNoteApi, dataCreateNote).Result;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Create Note Error: " + ex.Message);
                }
                string resultCreateNewNote = httpResponseCreateNewMessage.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error create new note: " + ex.Message);
            }
        }

        // Check dynamic/object has key/property
        public static bool IsPropertyExist(dynamic settings, string name)
        {
            if (settings is ExpandoObject)
                return ((IDictionary<string, object>)settings).ContainsKey(name);

            return settings.GetType().GetProperty(name) != null;
        }

        // Request to login server:
        public async void SendRequestLogin(JObject objData)
        {
            try
            {
                await Task.Run<string>(() => this.tmpResult = ThisAddIn.LoginAction(objData));
                if (this.tmpResult != "fail")
                {
                    this._formAccount.closeForm();
                    new messageSuccess().ShowDialog();
                }
                else
                {
                    this._formAccount.showError();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error send request login: " + ex.Message);
            }
        }

        // Reset default info
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
                MessageBox.Show("Reset default infor: " + ex.Message);
            }
        }

        // Get all reciptents|emails from mail item:
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
                        if (recipient.AddressEntry.Type == "EX")  // Check reciptent is exchange account???
                            list.Add(recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress);
                        else
                            list.Add(recipient.Address);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Get reciptent error: " + ex.Message);
                }
            }
            return list.ToArray();
        }

        // Get information account:
        public Accounts GetInformationAcounts()
        {
            return this.Application.Session.Accounts;
        }

        // Save new infor of current user
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
                MessageBox.Show("Error save new info: " + ex.Message);
            }
        }

        // Check default login: user logged or not???
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

        // Trigger event logout on btn logout
        public void TriggerLogoutBtn(string host, string email)
        {
            try
            {
                this.ResetDefaultInfo(host, email);
                this._formAccount = new accountForm();
                this._formAccount.ShowDialog();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error trigger event logout: " + exception.Message);
            }
        }

        // Login action:
        public static string LoginAction(JObject objData)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            string thisAddIn;
            dynamic data = JObject.Parse(objData.ToString());
            string str = (string)(data["host"].ToString() + "/api/mail/loginFromOutlook");
            JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
            jsonObject.Add("email", data["email"].ToString());
            jsonObject.Add("password", data["password"].ToString());
            HttpClient httpLoginClient = new HttpClient();
            StringContent stringContent = new StringContent(jsonObject.ToString(), Encoding.UTF8, "application/json");
            Task<HttpResponseMessage> httpResponseLogin = httpLoginClient.PostAsync(str, stringContent);
            HttpResponseMessage resultLoginResponse = httpResponseLogin.Result;
            AggregateException exception = httpResponseLogin.Exception;
            TaskStatus status = httpResponseLogin.Status;
            if (!resultLoginResponse.IsSuccessStatusCode)
            {
                thisAddIn = "fail";
            }
            else
            {
                Task<string> responseLogin = resultLoginResponse.Content.ReadAsStringAsync();
                if ((responseLogin.Result == null || !(responseLogin.Result != "") ? true : responseLogin.Result == "\"\""))
                {
                    thisAddIn = "fail";
                }
                else
                {
                    dynamic dataResponse = JObject.Parse(responseLogin.Result);
                    Globals.ThisAddIn.resultLogin = dataResponse.data;
                    thisAddIn = (string)((dynamic)Globals.ThisAddIn.resultLogin).ToString();
                }
            }
            if (thisAddIn == "fail")
            {
                Globals.ThisAddIn.ResetDefaultInfo(data["host"].ToString(), data["email"].ToString());
            }
            else
            {
                Globals.ThisAddIn.SaveYourInfo(data["host"].ToString(), data["email"].ToString(), ((dynamic)Globals.ThisAddIn.resultLogin).ToString());
                Globals.ThisAddIn.StartAutoThreads();
            }
            return thisAddIn;
        }

        // Parse data from email item:
        public static object ParseDataFromEmail(MailItem mail)
        {
            string now = DateTime.Now.ToString("yyyyMMddHHmmss");
            string folderRoot = string.Concat(System.Windows.Forms.Application.LocalUserAppDataPath, "\\contentEmailTemp\\mail_", now);
            if (!Directory.Exists(folderRoot))
            {
                Directory.CreateDirectory(folderRoot);
            }
            string[] strArrays = new string[] { ".png", ".jpg", ".gif", ".bmp" };
            string hTMLBody = null;
            string folderAttachments = string.Concat(new object[] { folderRoot, "\\", now, "-attachments" });
            string zipFileName = string.Concat(new object[] { now, "-attachments.zip" });
            string zipFilePath = string.Concat(folderRoot, "\\", zipFileName);
            string[] allReciptents = new string[0];
            allReciptents = ThisAddIn.GetAllReciptents(mail);
            hTMLBody = mail.HTMLBody;
            foreach (Outlook.Attachment attachment in mail.Attachments) // Convert media attachment to base 64
            {
                string base64String = null;
                string fileNameAttachment = string.Concat(now, "_", attachment.FileName);
                string fileExtAttachment = Path.GetExtension(string.Concat(folderRoot, "\\", fileNameAttachment)).Replace(".", "");
                string urlDefaultMicrosoft = "http://schemas.microsoft.com/mapi/proptag/0x3712001E";
                attachment.SaveAsFile(string.Concat(folderRoot, "\\", fileNameAttachment));
                string property = (string)((dynamic)attachment.PropertyAccessor.GetProperty(urlDefaultMicrosoft));
                try
                {
                    if (strArrays.Any<string>(new Func<string, bool>(attachment.FileName.Contains)))
                    {
                        using (Image image = Image.FromFile(string.Concat(folderRoot, "\\", fileNameAttachment)))
                        {
                            using (MemoryStream memoryStream = new MemoryStream())
                            {
                                image.Save(memoryStream, image.RawFormat);
                                base64String = Convert.ToBase64String(memoryStream.ToArray());
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Error convert media to base64: " + exception.Message);
                }
                try
                { // Format html body email
                    if (property == "")
                    {
                        if (!Directory.Exists(folderAttachments))
                        {
                            Directory.CreateDirectory(folderAttachments);
                        }
                        File.Move(string.Concat(folderRoot, "\\", fileNameAttachment), string.Concat(folderAttachments, "\\", fileNameAttachment));
                    }
                    else if (hTMLBody.ToLower().Contains(string.Concat("cid:", property.ToLower())))
                    {
                        hTMLBody = hTMLBody.Replace(string.Concat("\"cid:", property, "\""), string.Concat("data:image/", fileExtAttachment, ";base64,", base64String));
                    }
                }
                catch (Exception exception1)
                {
                    MessageBox.Show("Format HtmlBody Error: " + exception1.Message);
                }
            }
            JsonObject jsonObject = new JsonObject(new KeyValuePair<string, JsonValue>[0]);
            jsonObject.Add("contentMail", hTMLBody);
            jsonObject.Add("entry_id", mail.EntryID);
            jsonObject.Add("subject", mail.Subject);
            jsonObject.Add("fromEmail", Settings.Default.email);
            jsonObject.Add("addresses", JsonConvert.SerializeObject(allReciptents));
            if (Directory.Exists(folderAttachments)) // if has attachments:
            {
                ZipFile.CreateFromDirectory(folderAttachments, zipFilePath);
                jsonObject.Add("pathAttachment", zipFilePath);
                jsonObject.Add("fileName", zipFileName);
            }
            return jsonObject;
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
