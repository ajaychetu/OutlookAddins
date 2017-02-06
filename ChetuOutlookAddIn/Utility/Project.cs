using Excel;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ChetuOutlookAddIn.Utility
{
    public class Project
    {
        public int EmailCounter { get; set; }


        /// <summary>
        /// Getting data from Projects excel
        /// </summary>
        /// <returns></returns>
        public DataTable GetLiveProjectsDetailFromExcelFile()
        {
            DataSet result;
            DataTable dataTable = new DataTable();
            
            try
            {

                string path = ConfigurationManager.AppSettings["LiveProjectsExcel"].ToString();

                using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {

                    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                        result = excelReader.AsDataSet();
                        //4. DataSet - Create column names from first row
                        excelReader.IsFirstRowAsColumnNames = true;

                        result = excelReader.AsDataSet();
                        dataTable = result.Tables[0];
                        

                        return dataTable;
                    }
                }
            }

            catch 
            {
                throw;
            }
        }

        /// <summary>
        /// Programmatically Search Within a Specific Folder
        /// </summary>
        public bool SearchMorningSnap(DataRow dataRow)
        {
            try
            {
                Outlook.MAPIFolder inbox = Globals.ThisAddIn.Application.ActiveExplorer().Session.
                    GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                string folderName = ConfigurationManager.AppSettings["MailBoxFolderName"].ToString();
                Outlook.Items items = inbox.Folders[folderName].Items;
                Outlook.MailItem mailItem = null;
                

                object folderItem;
                string subjectName = string.Empty;

                string sFilter = " [ReceivedTime] >= '" + DateTime.Now.AddDays(Convert.ToInt32(ConfigurationManager.AppSettings["TestingDay"].ToString())).ToString("yyyy-MM-dd") + " 00:00' ";
                sFilter += " AND [ReceivedTime] < '" + DateTime.Now.AddDays(Convert.ToInt32(ConfigurationManager.AppSettings["TestingDay"].ToString()) + 1).ToString("yyyy-MM-dd") + " 00:00' ";

                //int a = items.Count;

                items = items.Restrict(sFilter);

                //int c = items.Count;

                string filter = ConfigurationManager.AppSettings["FilterKeyword"].ToString();

                folderItem = items.Find(filter);

                bool found = false;

                while (folderItem != null)
                {
                    mailItem = folderItem as Outlook.MailItem;

                    // Less than zero t1 is earlier than t2. Zero t1 is the same as t2. Greater than zero t1 is later than t2. 
                    if (mailItem != null)
                    {
                        if (mailItem.Subject.Contains(Convert.ToString(dataRow["Project"])))
                        {
                            found = true;
                            /*if (!(ReceivedDate.Hour <= Convert.ToInt32(ConfigurationManager.AppSettings["Hours"].ToString())
                                 && ReceivedDate.Minute <= Convert.ToInt32(ConfigurationManager.AppSettings["Mins"].ToString())))
                                {
                                }*/
                            break;
                        }
                        else
                        {
                            /// later will add

                        }
                    }
                    folderItem = items.FindNext();
                }

                //Check Morning snap status
                if (!found)
                {
                    Trace.trace.AppendLine("The Morning Sanp is missing of " + dataRow["Project"].ToString() + " project.</br>");
                    // The email send featur for TL and FM, We will implement later.
                    /*string subject = Convert.ToString(dataRow["Project"]) + " Project Morning Snap Missing";
                     string body = EmailTemplate("Ajay", Convert.ToString(dataRow["Project"]));
                     MailSendDelegate mailSend = new MailSendDelegate(SendNotification);
                     mailSend.BeginInvoke(Convert.ToString(dataRow["TL"]), subject, body, Convert.ToString(dataRow["FM"]),null, null);                    
                     //SendNotification(Convert.ToString(dataRow["TL"]), subject, body, Convert.ToString(dataRow["FM"]));
                     */
                }

                return found;
            }
            catch (Exception ex)
            {
                Trace.trace.AppendLine("SearchMorningSnap: Error" + ex.Message);
                throw;
            }
        }

        public void SendDetails(string to, string name, string details, string subject, string bcc)
        {
            try
            {
                 string body = EmailTemplate("", details);
                 MailSendDelegate mailSend = new MailSendDelegate(SendNotification);
                 mailSend.BeginInvoke(to, subject, body, "", bcc, null, null);                    
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Email Template 
        /// </summary>
        /// <param name="tLName"></param>
        /// <param name="projectName"></param>
        /// <returns></returns>
        public string EmailTemplate(string tLName, string projectName, bool type = false)
        {
            try
            {
                StringBuilder ObjBuilder = new StringBuilder();

                ObjBuilder.Append("<html><head></head><body>");

                //boday content

                ObjBuilder.Append("<table cellpadding='1' cellspacing='6' align='center' width='90%' style='background - color:LightYellow; font - family:Arial,Verdana; font - size:10pt; color: Maroon; border: 1 solid Orange;'>");
                ObjBuilder.Append("<tr><td colspan = '2'> Hi " + tLName + ",</td></tr>");
                if (type)
                {
                    ObjBuilder.Append("<tr><td colspan = '2'> The morning snap of the " + projectName + " project is missing on " + DateTime.Now.ToShortDateString() + ".</td></tr>");
                }
                else
                {
                    ObjBuilder.Append("<tr><td colspan = '2'> " + projectName + "</td></tr>");
                }
                ObjBuilder.Append("<tr><td colspan = '2'></td></tr>");
                ObjBuilder.Append("<tr><td align = 'right' style = 'color:Red; font-weight:bold;'>Note :</td><td style = 'color:Red;'>This is confidential mail.So Please, keep the privacy of this mail.</td></tr></table>");

                //footer of email
                ObjBuilder.Append("<table cellpadding='1' cellspacing='2' align='center' width='90%' style='font - family:Arial, Verdana; font - size:11pt;'>");
                ObjBuilder.Append("<tr><td align='left' valign = 'middle' style = 'Height:50px; font-weight:bold; text-decoration:underline;'>");
                ObjBuilder.Append("Thanks and Regards,</td></tr><tr><td style = 'font-weight:bold; color:Green;'>HR Team</td></tr><tr><td>hr@chetu.com</td></tr></table>");

                ObjBuilder.Append("</body></html>");

                return ObjBuilder.ToString();
            }
            catch
            {
                throw;
            }

        }

        /// <summary>
        /// Send Email Notification method
        /// </summary>
        /// <param name="to"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="cc"></param>
        public void SendNotification(string to, string subject, string body, string cc = "", string bcc  ="")
        {
       
            try
            {
                //--All mail setting initialized from app.config
                using (MailMessage mail = new MailMessage())
                {
                    //mail.To.Add("premk@chetu.com");
                    mail.To.Add(to);
                    if (!string.IsNullOrEmpty(cc))
                    {
                        mail.CC.Add(cc);
                    }
                    if (!string.IsNullOrEmpty(bcc))
                    {
                        mail.Bcc.Add(bcc);
                    }
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.IsBodyHtml = true;

                    using (SmtpClient smtp = new SmtpClient())
                    {
                        smtp.Send(mail);
                    }
                }

            }
            catch (Exception ex)
            {
                //---As this is mail send method so not throwing exception to avoid mail send break for futher emails
                //--Log Error those mails failed
                Trace.trace.AppendLine("<br/>SendNotification: Error " + ex.Message);
                Trace.trace.AppendLine("Mail Send member to & CC that failed " + to + " & " + cc);
                if (EmailCounter < 1)
                {
                    EmailCounter += 1;
                    MessageBox.Show("Error- [SendNotification Error] " + ConfigurationManager.AppSettings["MorningSnapSubject"].ToString()+ ": " + ex.Message);
                    SendNotification("ajays@chetu.com", "Error- [SendNotification Error] " + ConfigurationManager.AppSettings["MorningSnapSubject"].ToString(),
                    Trace.trace.ToString());
                }
 
            }
        }


        #region Delegates
        delegate void MailSendDelegate(string to, string subject, string body, string cc, string bcc);
        #endregion

    }
}
