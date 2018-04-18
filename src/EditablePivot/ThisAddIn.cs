using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EditablePivot.BaseClasses;

namespace EditablePivot
{
    public partial class ThisAddIn
    {
        // public variables
        public static Office.IRibbonUI _ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SheetChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler(ThisAddIn_SheetChange);
            this.Application.SheetSelectionChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler(ThisAddIn_SelectionChange);
            this.Application.SheetActivate += new Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(ThisAddIn_SheetActivate);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void ThisAddIn_SelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            //MessageBox.Show("ThisAddIn_SelectionChange");
            MyRibbon.InvalidateRibbon();
        }

        void ThisAddIn_SheetChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            //MessageBox.Show("ThisAddIn_SheetChange");
            MyRibbon.InvalidateRibbon();

            PivotHelper.PostSheetChange(Target);
        }

        void ThisAddIn_SheetActivate(object Sh)
        {
            //MessageBox.Show("ThisAddIn_SheetActivate");
            MyRibbon.InvalidateRibbon();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
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
