using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using UserInformation;
using Apar_Stamping_AR.Model;
using MSC = Microsoft.SharePoint.Client;

namespace Apar_Stamping_AR.Model
{
    class CustomSharePointUtility
    {
        static UserOperation _UserOperation = new UserOperation();
        public static StreamWriter logFile;
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");
        public static string Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }
        public static MSC.ClientContext GetContext(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                var securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_AppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                WriteLog("Error in  CustomSharePointUtility.GetContext: " + ex.ToString());
                return null;
            }
        }


        public static List<ErrorModel> GetActiveErrorList(string siteUrl, string listName)
        {
            List<ErrorModel> errorModels = new List<ErrorModel>();
            try
            {
                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    if (context != null)
                    {
                        MSC.List list = context.Web.Lists.GetByTitle(listName);
                        MSC.ListItemCollectionPosition itemPosition = null;

                        while (true)
                        {
                            MSC.CamlQuery camlQuery = new MSC.CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;

                            camlQuery.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='SystemUpdated' /><Value Type='Choice'>No</Value></Eq></Where></Query><ViewFields><FieldRef Name='UserDisplayName' /><FieldRef Name='DisplayNo' /><FieldRef Name='URL' /><FieldRef Name='SystemUpdated' /><FieldRef Name='ID' /></ViewFields><QueryOptions /></View>";






                            MSC.ListItemCollection Items = list.GetItems(camlQuery);

                            context.Load(Items);
                            context.ExecuteQuery();
                            itemPosition = Items.ListItemCollectionPosition;
                            foreach (MSC.ListItem item in Items)
                            {
                                errorModels.Add(new ErrorModel
                                {
                                    ID = Convert.ToInt32(item["ID"]),
                                    URL = Convert.ToString(item["URL"]).Trim(),
                                    SystemUpdate = Convert.ToString(item["SystemUpdated"]).Trim(),

                                });
                            }
                            if (itemPosition == null)
                            {
                                break; // TODO: might not be correct. Was : Exit While
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {

            }
            return errorModels;
        }



        public static void UpdateURL(List<ErrorModel> errorModels, string siteUrl, string listName)
        {
            if (errorModels.Count > 0)
            {
                for (var i = 0; i < errorModels.Count; i++)
                {
                    var URL = errorModels[i].URL;

                    HttpWebRequest request = HttpWebRequest.CreateHttp(URL);
                    request.Accept = "application/json;odata=verbose";
                    Stream webStream = request.GetResponse().GetResponseStream();
                    StreamReader responseReader = new StreamReader(webStream);
                    string response = responseReader.ReadToEnd();

                    response = response.Replace("\"", "");

                    if (response == "done")
                    {
                        UpdateStatus(errorModels[i].ID, siteUrl, listName);
                        Console.WriteLine(errorModels[i].ID);
                    }

                }

            }

        }

        public static void UpdateStatus(int ID, string siteUrl, string ListName)
        {
            using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
            {
                if (context != null)
                {
                    MSC.List list = context.Web.Lists.GetByTitle(ListName);
                    MSC.ListItem listItem = null;
                    
                    MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
                    listItem = list.GetItemById(ID);

                    listItem["SystemUpdated"] = "Yes";
                    listItem.Update();
                    try
                    {
                        context.ExecuteQuery();


                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
        }


        public static void WriteLog(string logmsg)
        {
            // StreamWriter logFile;

            try
            {

                string LogString = DateTime.Now.ToString("dd/MM/yyyy HH:MM") + " " + logmsg.ToString();

                //  logFile.WriteLine(DateTime.Now);
                //  logFile.WriteLine(logmsg.ToString());
                //logFile.WriteLine(LogString);

                //logFile.Close();
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());

            }

        }

        public static AppConfiguration GetSharepointCredentials(string siteUrl)
        {
            AppConfiguration _AppConfiguration = new AppConfiguration();

            _AppConfiguration.ServiceSiteUrl = siteUrl;// _UserOperation.ReadValue("SP_Address");
            _AppConfiguration.ServiceUserName = _UserOperation.ReadValue("SP_USER_ID_Live");
            _AppConfiguration.ServicePassword = Decrypt(_UserOperation.ReadValue("SP_Password_Live"));

            return _AppConfiguration;
        }



    }
    public class AppConfiguration
    {
        public string ServiceSiteUrl;
        public string ServiceUserName;
        public string ServicePassword;
    }

}
