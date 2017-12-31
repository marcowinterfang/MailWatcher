using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace AddIn
{
    public partial class ThisAddIn
    {
        Outlook.Items _items;
        private string MAIL_PATH = @"C:\Outlook\Mails\";
        private string ATTATCHMENT_PATH = @"C:\Outlook\Attachments\";
        private string PDF_PATH = @"C:\Outlook\PDF\";
        private string TXT_PATH = @"C:\Outlook\TXT\";
        private string PDF_FILE_PATH = @"\\dlcfs0002\JAN\Janssen\Acceptance\OriginDocument\";
        private const string EDC_NAME = "specicname";
        private const string WEB_CONTRACT_NAME = "typename";
        private const string CONFIG_PATH = @"C:\Outlook\Log\Config.txt";
        //private const string RECIEVER_ADDRESS = @"RA-JANJP-PV-CAC@ITS.JNJ.com";
        private const string RECIEVER_ADDRESS = @"othername@gmail.com";
        private const string PVCAC = @"yourname@gmail.com";
        //private const string PVCAC = "RA-JANJP-PV-CAC@ITS.JNJ.com";
        //CHANGE IT TO SETTINGS
        //private const string CASE_SENDER_ADDRESS = "ra-janjp-pv-ae@its.jnj.com";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                LogHelper.Write(LogType.Debug, "start load");
                MAIL_PATH = System.Configuration.ConfigurationManager.AppSettings["mailpath"];
                ATTATCHMENT_PATH = System.Configuration.ConfigurationManager.AppSettings["attachmentpath"];
                #region Annotation
                //XmlDocument doc = new XmlDocument();
                //doc.Load("AddInConfig.xml");
                //XmlElement rootElement = doc.DocumentElement;
                //var pathElements = rootElement.GetElementsByTagName("Type");
                //foreach (XmlNode node in pathElements)
                //{
                //    string strName = ((XmlElement)node).GetAttribute("name");
                //    XmlNodeList subPathNodes = ((XmlElement)node).GetElementsByTagName("Path");
                //    if (subPathNodes.Count == 1)
                //    {
                //        if (strName == "Mail")
                //        {
                //            MAIL_PATH = subPathNodes[0].InnerText;
                //        }
                //        else if (strName == "Attachment")
                //        {
                //            ATTATCHMENT_PATH = subPathNodes[0].InnerText;
                //        }
                //    }
                //}
                #endregion
                //Undo VALIDATE the path
                CheckPath();
                PdfFilePath();

                //Outlook.NameSpace ns = Application.GetNamespace("MAPI");
                //string mailBoxName = PVCAC;

                //Outlook.Recipient objRecipient = ns.CreateRecipient(mailBoxName);

                //Outlook.MAPIFolder inbox = ns.GetSharedDefaultFolder(objRecipient, Outlook.OlDefaultFolders.olFolderInbox).Parent as Outlook.MAPIFolder;

                //Outlook.MAPIFolder pvInbox = inbox.Folders["受信トレイ"];
                //_items = pvInbox.Items;
                //_items.ItemAdd += MailArrival;
                //LogHelper.Write(LogType.Debug, "Loading complete");
                _items = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;
                _items.ItemAdd += MailArrival;
            }
            catch (Exception ex)
            {
                LogHelper.Write(LogType.Fatal, "Start failed");
                LogHelper.Write(ex);
            }
        }

        private void MailArrival(object Item)
        {
            try
            {
                LogHelper.Write(LogType.Debug, "Mail Arrived");
                Outlook.MailItem item = Item as Outlook.MailItem;
                if (item != null)
                {
                    string mailName = "";

                    if (item.Subject != null)
                    {
                        mailName = MAIL_PATH + item.Subject.Replace(":", "_") + DateTime.Now.ToString("yyyyMMddHHmmss") + ".msg";
                    }
                    else
                    {
                        mailName = MAIL_PATH + mailName + "NoSubject" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".msg";
                    }
                    LogHelper.Write(LogType.Debug, "Mail Name: " + mailName);

                    string fileName = "";
                    try
                    {
                        mailName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".msg";
                        LogHelper.Write(LogType.Debug, mailName);
                        item.SaveAs(MAIL_PATH + mailName, Outlook.OlSaveAsType.olMSG);
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Write(LogType.Error, "Error Start: Mail Save Error!");
                        LogHelper.Write(LogType.Error, ex.Message);
                        LogHelper.Write(LogType.Error, item.Subject == null ? string.Empty : item.Subject.ToString());
                        LogHelper.WriteCase(item.Subject == null ? string.Empty : item.Subject.ToString());
                        LogHelper.Write(LogType.Error, "End: Mail Save Error");
                    }
                    int pdfCount = 0;
                    LogHelper.Write(LogType.Debug, "Deal with attachments");

                    #region Save Attatchments
                    //if (item.Attachments.Count > 0)
                    //{
                    //    foreach (Outlook.Attachment attachment in item.Attachments)
                    //    {
                    //        string extension = Path.GetExtension(attachment.FileName);
                    //        if (extension != ".msg")
                    //        {
                    //            continue;
                    //        }
                    //        pdfCount++;
                    //        string attachmentName = ATTATCHMENT_PATH + attachment.FileName;
                    //        attachment.SaveAsFile(attachmentName);
                    //        attachmentName = PDF_FILE_PATH + attachment.FileName;
                    //        attachment.SaveAsFile(attachmentName);
                    //        fileName += attachment.FileName + "|,|";
                    //    }
                    //}
                    #endregion
                    fileName = SaveAttachment(item,ref pdfCount);
                    LogHelper.Write(LogType.Debug, "Get origindoctype");

                    //Judge the mail type (转信/返信/他番……)
                    //Emergency Status
                    int originDocType = (int)GetOriginDocType(item.Subject);
                    LogHelper.Write(LogType.Debug, "Get emergency value");

                    //string caseGuidanceType = SqlHelper.GetCaseGuidanceType(item.Subject);
                    int emergency = GetEmergency(item.Body, item.Subject);
                    if (CompareInfoDate(item.Body))
                    {
                        if (emergency > 2)
                            emergency = 2;
                    }
                    
                    LogHelper.Write(LogType.Debug, "Get sender address");
                    string senderAddress = GetSenderAddress(item);
                    LogHelper.Write(LogType.Debug, "Save in database");

                    //UNDO need to add file type(contract or edc)
                    string receiveTime = item.ReceivedTime == null ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"): item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss");
                    int recordId = SqlHelper.SaveMailRecord(item.EntryID, mailName, fileName, senderAddress, item.Body
                        , originDocType, item.Subject, emergency, pdfCount,receiveTime);
                    LogHelper.Write(LogType.Debug, "Save attachments into pdf folder");

                    if (originDocType == 0 && item.Attachments.Count > 0)
                    {
                        foreach (Outlook.Attachment att in item.Attachments)
                        {
                            if (Path.GetExtension(att.FileName) == ".pdf")
                            {
                                att.SaveAsFile(PDF_PATH + recordId + "#" + att.FileName);
                                break;
                                //SavePdfFile(att);
                            }
                        }
                    }
                    //else if(originDocType > 0 && item.Attachments.Count > 0)
                    //{
                    //    foreach (Outlook.Attachment it in item.Attachments)
                    //    {
                    //        if (Path.GetExtension(it.FileName) == ".pdf")
                    //        {
                    //            SavePdfFile(it);
                    //        }
                    //    }
                    //}
                    LogHelper.Write(LogType.Debug, "Mail arrive complete");
                }
            }
            catch (Exception ex)
            {
                Outlook.MailItem item = Item as Outlook.MailItem;
                if (item != null && item.Subject != null)
                {
                    LogHelper.WriteCase(item.Subject);
                }
                LogHelper.Write(LogType.Error, ex.Message);
                LogHelper.Write(ex);
            }
        }
        private string SaveAttachment(Outlook.MailItem mail,ref int pdfCount)
        {
            try
            {
                string fileName = string.Empty; 
                if (mail.Attachments.Count > 0)
                {
                    foreach (Outlook.Attachment attachment in mail.Attachments)
                    {
                        string extension = Path.GetExtension(attachment.FileName);
                        if (extension == ".gif" || extension == ".png" || extension == ".jpg")
                        {
                            continue;
                        }
                        pdfCount++;
                        string attachmentName = ATTATCHMENT_PATH + attachment.FileName;
                        attachment.SaveAsFile(attachmentName);
                        attachmentName = PDF_FILE_PATH + attachment.FileName;
                        attachment.SaveAsFile(attachmentName);
                        if (extension == ".sgm")
                        {
                            continue;
                        }
                        fileName += attachment.FileName + "|,|";
                    }
                    if (fileName.Length > 3 && fileName.Substring(fileName.Length-3) == "|,|")
                    {
                        fileName = fileName.Substring(0, fileName.Length - 3);
                    }
                }
                return fileName;
            }
            catch (Exception ex)
            {
                LogHelper.Write(LogType.Error, ex.Message);
                return string.Empty;
            }
        }
        private bool CompareInfoDate(string mailBody)
        {
            try
            {
                string flag = "情報入手日：";
                bool result = true;
                if (!string.IsNullOrEmpty(mailBody))
                {
                    string[] lines = mailBody.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    var infoLine = lines.FirstOrDefault(x => x.Contains(flag));
                    if (!string.IsNullOrEmpty(infoLine))
                    {
                        infoLine = infoLine.Substring(flag.Length - 1);
                    }
                    DateTime dt;
                    result = DateTime.TryParse(infoLine, out dt);
                    if (result && dt.Date == DateTime.Now.Date)
                    {
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write(LogType.Error, ex.Message);
            }
            return false;
        }
        private void SavePdfFile(Outlook.Attachment file)
        {
            try
            {
                file.SaveAsFile(PDF_FILE_PATH + file.FileName);
            }
            catch (Exception ex)
            {
                LogHelper.Write(ex);
            }
        }
        private void PdfFilePath()
        {
            try
            {
                if (File.Exists(CONFIG_PATH))
                {
                    string config = File.ReadAllText(CONFIG_PATH);
                    if (!string.IsNullOrEmpty(config))
                    {
                        PDF_FILE_PATH = config;
                    }
                }
                if (!Directory.Exists(PDF_FILE_PATH))
                {
                    Directory.CreateDirectory(PDF_FILE_PATH);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write(ex);
            }
        }
        private int GetEmergency(string mailbody, string subject)
        {
            int r = 9;
            string content;
            if (string.IsNullOrEmpty(mailbody) && string.IsNullOrEmpty(subject))
            {
                return r;
            }
            else if (string.IsNullOrEmpty(mailbody) && !string.IsNullOrEmpty(subject))
            {
                content = subject;
            }
            else if (string.IsNullOrEmpty(subject) && !string.IsNullOrEmpty(mailbody))
            {
                content = mailbody;
            }
            else
            {
                content = subject + mailbody;
            }
            bool isEmergency = GetIsTodayInfoDate(mailbody);
            //I类
            List<string> dltList = new List<string>() { "永眠","他界","生命を脅かす","死亡の恐れ","突然死",
                "死亡につながるおそれ","死因","剖検","検視","生命","逝去","心肺停止","葬式","死去","天国",
                "逝去","亡くなった","蘇生","Fatal","Death","Life threatening","Autopsy","生命を脅かす、死亡の恐れ"};
            //II类
            //III类
            string disablity = "障害、障害につながる";
            //IV类
            string hospital = "入院";
            //VI类
            string anomalies = "先天異常";
            foreach (var dlt in dltList)
            {
                if (content.Contains(dlt))
                {
                    r = 3;
                }
            }
            if (content.Contains(disablity) || content.Contains(hospital) || content.Contains(anomalies))
            {
                r = 5;
            }
            if (r==3 && isEmergency)
            {
                r = 1;
            }
            else if (r==4 && isEmergency)
            {
                r = 2;
            }
            else if (r==9 && isEmergency)
            {
                r = 4;
            }
            return r;
        }
        private bool GetIsTodayInfoDate(string mailBody)
        {
            string keyWord = "情報入手日：";
            DateTime infoDate;
            string strInfo = mailBody.Split(new string[] { "\r\n" }, StringSplitOptions.None).FirstOrDefault(x => x.Contains(keyWord));
            if (!string.IsNullOrEmpty(strInfo) && strInfo.Length > keyWord.Length)
            {
                strInfo = strInfo.Substring(keyWord.Length);
                if (DateTime.TryParse(strInfo, out infoDate))
                {
                    if (infoDate.Date != DateTime.Now.Date)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        private OriginDocType GetOriginDocType(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return OriginDocType.OTHER;
            }
            if (subject.Contains(WEB_CONTRACT_NAME))
            {
                return OriginDocType.Contract;
            }
            else if (subject.Contains(OriginDocType.IBR1L.ToString()))
            {
                return OriginDocType.IBR1L;
            }
            else if (subject.Contains(OriginDocType.IBR2L.ToString()))
            {
                return OriginDocType.IBR2L;
            }
            else if (subject.Contains(OriginDocType.RIS6U.ToString()))
            {
                return OriginDocType.RIS6U;
            }
            else if (subject.Contains(OriginDocType.TAP1U.ToString()))
            {
                return OriginDocType.TAP1U;
            }
            else if (subject.Contains(OriginDocType.ULT2C.ToString()))
            {
                return OriginDocType.ULT2C;
            }
            else if (subject.Contains(OriginDocType.XEP1L.ToString()))
            {
                return OriginDocType.XEP1L;
            }
            else if (subject.Contains(OriginDocType.ZYT1L.ToString()))
            {
                return OriginDocType.ZYT1L;
            }
            return OriginDocType.OTHER;
        }
        private void SetStatus(string subject) { }
        private void CheckPath()
        {
            if (!Directory.Exists(MAIL_PATH))
            {
                Directory.CreateDirectory(MAIL_PATH);
            }
            if (!Directory.Exists(ATTATCHMENT_PATH))
            {
                Directory.CreateDirectory(ATTATCHMENT_PATH);
            }
            if (!Directory.Exists(PDF_PATH))
            {
                Directory.CreateDirectory(PDF_PATH);
            }
            if (!Directory.Exists(TXT_PATH))
            {
                Directory.CreateDirectory(TXT_PATH);
            }
        }
        private string GetSenderAddress(Outlook.MailItem item)
        {
            string address = "";
            try
            {
                Outlook.Recipient recip;
                Outlook.ExchangeUser exUser;
                if (item.SenderEmailType.ToLowerInvariant() == "ex")
                {
                    recip = Globals.ThisAddIn.Application.GetNamespace("MAPI").CreateRecipient(item.SenderEmailAddress);
                    exUser = recip.AddressEntry.GetExchangeUser();
                    address = exUser.PrimarySmtpAddress;
                }
                else
                {
                    address = item.SenderEmailAddress.Replace("'", "");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write(ex);
                LogHelper.Write(LogType.Error, "Location:GetSenderAddress");
            }
            return address;
        }
        private void WebContrat(Outlook.MailItem item, ref string fileName)
        {
            if (item.Attachments.Count > 0)
            {
                foreach (Outlook.Attachment attachment in item.Attachments)
                {
                    string extension = Path.GetExtension(attachment.FileName);
                    if (extension == ".jpg" || extension == ".jpng" || extension == ".png")
                    {
                        continue;
                    }
                    if (extension == ".pdf")
                    {
                        attachment.SaveAsFile(PDF_PATH + attachment.FileName);
                    }
                    string attachmentName = ATTATCHMENT_PATH + attachment.FileName;
                    fileName += attachmentName + "|,|";
                    attachment.SaveAsFile(attachmentName);
                }
            }
        }
        private void EDC(Outlook.MailItem item)
        {
            string body = item.HTMLBody;
            string strRegexLink = @"(?is)<a .*?>";
            string regexHref = @"<a[^>]*href=([""'])?(?<href>[^'""]+)\1[^>]*>";
            MatchCollection matchList = Regex.Matches(body, strRegexLink, RegexOptions.IgnoreCase);
            List<string> linkList = new List<string>();
            foreach (Match link in matchList)
            {
                if (link.Value.Contains("http"))
                {
                    if (Regex.IsMatch(link.Value, regexHref))
                    {
                        var a = Regex.Match(link.Value, regexHref, RegexOptions.IgnoreCase);
                        linkList.Add(a.Groups["href"].Value);
                    }
                }
            }
            //Select a link and link to the website to download the file
            //if (linkList.Any())
            //{
            //Process.Start("iexplore.exe", linkList[0]);
            //}
            //read the file get the content,and download other files which mentioned by the file content
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 备注: Outlook 不会再遇到这种问题。如果具有
            //关闭 Outlook 时必须运行的代码，请参阅 http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
        public enum OriginDocType
        {
            Contract = 0,
            IBR1L = 1,
            IBR2L = 2,
            RIS6U = 3,
            TAP1U = 4,
            ULT2C = 5,
            XEP1L = 6,
            ZYT1L = 7,
            OTHER = 8
        }
        
    }
}

