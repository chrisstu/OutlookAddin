using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookExtras
{
    public partial class ThisAddIn
    {
        private string xml = @"";
        private Outlook.Explorer _activeExplorer;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _activeExplorer = Application.Explorers[1];
            _activeExplorer.SelectionChange += _activeExplorer_SelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void _activeExplorer_SelectionChange()
        {
            if (_activeExplorer.Selection.Count > 0)
            {
                Object selObject = _activeExplorer.Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = _activeExplorer.Selection[1];
                    string senderEmailAddress = mailItem.SenderEmailAddress;
                    if (Ribbon1.ribbon!=null)
                        Ribbon1.Reset(senderEmailAddress);
                }
            }
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
