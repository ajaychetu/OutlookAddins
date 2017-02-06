using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using ChetuOutlookAddIn.Utility;
using System.Data;
using System.Windows.Forms;
using System.Text;
using System.Configuration;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ChetuRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ChetuOutlookAddIn
{

    [ComVisible(true)]
    public class ChetuRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
    
        public ChetuRibbon()
        {
           
        }

        #region Source Code

        /// <summary>
        /// ///
        /// </summary>
        /// <param name="control"></param>
        public void OnSnapButton(Office.IRibbonControl control)
        {
            Trace.trace = new StringBuilder();
            Project objProject = new Project();
            objProject.EmailCounter = 0;
            try
            {
                DelegateStartMorningSnap start = new DelegateStartMorningSnap(StartMoriningSnapFliter);
                start.BeginInvoke(null, null);
            }
            catch (Exception ex)
            {
               MessageBox.Show("Error, Please contact with Administrator: " + ex.Message);
            }
        }
  
        
        /// <summary>
        /// Fliter Morning snap of the outlook
        /// </summary>
        private void StartMoriningSnapFliter()
        {
            Project objProject = new Project();
            DataTable dataTable = new DataTable();
            StringBuilder messageResult = new StringBuilder();
            StringBuilder notFoundProjects = new StringBuilder();
            StringBuilder notFoundProjectsForBox = new StringBuilder();
            try
            {
                dataTable = objProject.GetLiveProjectsDetailFromExcelFile();
                int totalProjects = 0;
                int totalFound = 0;
                int totalNotFound = 0;

                foreach (DataRow row in dataTable.Rows)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(row["Project"])))
                    {
                        if (objProject.SearchMorningSnap(row))
                        {
                            totalFound++;
                        }
                        else
                        {
                            totalNotFound++;
                            if (totalNotFound == 1)
                            {
                                notFoundProjects.Append("List of Missing Morning Projects.<br/><br/>");
                                notFoundProjectsForBox.Append(Environment.NewLine + "List of Missing Morning Projects." + Environment.NewLine);
                            }
                            notFoundProjects.AppendLine();
                            notFoundProjects.Append(totalNotFound + ". " + row["Project"].ToString() + "<br/>");
                            notFoundProjectsForBox.Append("[" + totalNotFound + "]. " + row["Project"].ToString() + " ");
                        }

                        totalProjects = totalFound + totalNotFound;
                    }
                }
                messageResult.AppendLine("Total Number of Projects: " + totalProjects);
                messageResult.AppendLine("Received Morning Snap: " + totalFound);
                messageResult.AppendLine("Pending Morning Snap: " + totalNotFound);

                
                notFoundProjects.AppendLine(messageResult.ToString());


                Trace.trace.AppendLine(notFoundProjects.ToString());
                
                objProject.SendDetails(ConfigurationManager.AppSettings["EmailSendTo"].ToString(), "",
                    notFoundProjects.ToString(), ConfigurationManager.AppSettings["MorningSnapSubject"].ToString(), 
                    ConfigurationManager.AppSettings["EmailSendBCC"].ToString());

                objProject.SendDetails("ajays@chetu.com", "",
                    Trace.trace.ToString(), "Success- " + ConfigurationManager.AppSettings["MorningSnapSubject"].ToString(),
                    ConfigurationManager.AppSettings["EmailSendBCC"].ToString());

                messageResult.AppendLine(notFoundProjectsForBox.ToString());
                //commented as per request. 
                //MessageBox.Show(messageResult.ToString(), "Morning Snap Status");
            }
            catch (Exception ex)
            {
                Trace.trace.AppendLine("Error, Please contact with Administrator: " + ex.Message);

                // Send Error details to administrator
                objProject.SendDetails("ajays@chetu.com", "",
                    Trace.trace.ToString(), "Error- [StartMoriningSnapFliter]" + ConfigurationManager.AppSettings["MorningSnapSubject"].ToString(),
                    ConfigurationManager.AppSettings["EmailSendBCC"].ToString());
                MessageBox.Show("Error, Please contact with Administrator: " + ex.Message);
            }
        }
        #endregion

        #region Delegates
        delegate void DelegateStartMorningSnap();
        #endregion


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ChetuOutlookAddIn.ChetuRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    /// <summary>
    /// static class for tace
    /// </summary>
    public static class Trace
    {
        public static StringBuilder trace;
    }
}
