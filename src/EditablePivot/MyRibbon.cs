using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EditablePivot.BaseClasses;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace EditablePivot
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MyRibbon()
        {
        }

        #region static methods exposed
        public static void InvalidateRibbon()
        {
            ThisAddIn._ribbon.Invalidate();
            PivotHelper.SetEditable(Globals.ThisAddIn.Application.Selection);
        }
        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("EditablePivot.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            ThisAddIn._ribbon = this.ribbon;
        }

        public void clickAction(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "about":
                    {
                        Forms.frmAbout frm = new Forms.frmAbout();
                        frm.ShowDialog();
                        frm.Close();
                    }
                    break;
                case "configure":
                    {
                        PivotHelper.ConfigEditable(Globals.ThisAddIn.Application.Selection);
                    }
                    break;
            }
        }

        public Bitmap getImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "about":
                    return new Bitmap(Properties.Resources.about);
                case "configure":
                    return new Bitmap(Properties.Resources.TableSummarizeWithPivot);
                default:
                    return null;
            }
        }

        public bool getEnabled(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "configure":
                    return PivotHelper.IsPivotCell(Globals.ThisAddIn.Application.Selection);
            }
            return false;
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
