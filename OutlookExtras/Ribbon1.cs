using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookExtras
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        public static Office.IRibbonUI ribbon;
        private static Ribbon1 _this;
        private Explorer _explorer;
        private MAPIFolder _personalInBox;
        private MAPIFolder _clearLifeInBox;
        public Ribbon1()
        {
            _this = this;
        }

        public static void Reset(string senderEmailAddress)
        {
            RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item.Label = "YOYOY";
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookExtras.Ribbon1.xml");
        }

        public bool GetVisible(Office.IRibbonControl control)
        {
            return true;
        }

        public string GetMenuContent(Office.IRibbonControl control)
        {
            StringBuilder xmlString = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" >");
            if (_explorer.Selection.Count > 0)
            {
                Object selObject = _explorer.Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = _explorer.Selection[1];
                    string senderEmailAddress = mailItem.SenderEmailAddress;
                    xmlString.Append(string.Format(@"<button id='button1' label='{0}' onAction='MoveMailItem' />", senderEmailAddress));
                }
            }
            xmlString.Append(@"</menu>");
            return xmlString.ToString();
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            Ribbon1.ribbon = ribbonUI;
            _explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            _personalInBox = _explorer.Session.Folders["chris@cstuart.com"];
            _clearLifeInBox = _explorer.Session.Folders["chris.stuart@clearlifeltd.com"];
        }

        public void OnMyButtonClick2(Office.IRibbonControl control)
        {
        }

        public void MoveMailItem(Office.IRibbonControl control)
        {
            MailItem item = _explorer.Selection[1];

            //foreach (MAPIFolder subFolder in _clearLifeInBox.Folders)
            //    ShowFoldersName("chris.stuart@clearlifeltd.com", subFolder);
            //foreach (MAPIFolder subFolder in _personalInBox.Folders)
            //    ShowFoldersName("chris@cstuart.com", subFolder);
            (MAPIFolder mapiFolder, string folderId) = GetFolderId(control.Id);
            MAPIFolder folder = _explorer.Session.GetFolderFromID(folderId);
            //item.Move(folder);
        }

        private void ShowFoldersName(string prefix, MAPIFolder mapiFolder)
        {
            Debug.WriteLine(prefix + "/" + mapiFolder.Name + ":" + mapiFolder.EntryID);
            foreach (MAPIFolder subFolder in mapiFolder.Folders)
            {
                ShowFoldersName(prefix + "/"+ mapiFolder.Name, subFolder);
            }
        }

        private (MAPIFolder, string) GetFolderId(string folderName)
        {
            switch (folderName)
            {
                case "iDealing":
                    return (_personalInBox, "000000000B7D2699948B5B48905619A4C72ACF480100FBD3FAE6B2C0BD4DB29C82AE3E2CA713000000D106720000");
            }
            return (null, string.Empty);
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
}
