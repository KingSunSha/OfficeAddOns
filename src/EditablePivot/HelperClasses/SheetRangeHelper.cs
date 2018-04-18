using System;
using Microsoft.Office.Interop.Excel;

namespace EditablePivot.BaseClasses
{
    public class SheetRangeHelper
    {
        public const string tblCopyOptionsName = "TblCopyOptions";
        public const string worksheetPassword = "skates4u";

        public static void protectWorksheet(Worksheet sh, bool Protect, bool protectDrawingObjects = false)
        {
            if (Protect && !protectDrawingObjects)
            {
                sh.Protect(Password: worksheetPassword
                         , AllowUsingPivotTables: true
                         , AllowFiltering: true
                         , AllowSorting: true
                    //, DrawingObjects: false
                         , UserInterfaceOnly: true);
            }
            else if (Protect && protectDrawingObjects)
            {
                sh.Protect(Password: worksheetPassword
                         , AllowUsingPivotTables: true
                         , AllowFiltering: true
                         , AllowSorting: true
                         , DrawingObjects: false
                         , UserInterfaceOnly: true);
            }
            else
            {
                sh.Unprotect(worksheetPassword);
            }
        }

        public static void protectWorkbook(Workbook caller, bool Protect)
        {
            //loop through all sheets
            foreach (Worksheet s in caller.Worksheets)
            {
                if (Protect)
                {
                    // King Sun 2012-02-22 add parameters
                    // to enable/disable
                    //       AllowUsePivot: Use Pivot function
                    //       AllowFilter: use AutoFilter
                    s.Protect(Password: worksheetPassword
                            , AllowUsingPivotTables: true
                            , AllowFiltering: true
                            , UserInterfaceOnly: true);
                }
                else
                {
                    s.Unprotect(worksheetPassword);
                }

            }
        }

        public static void ResetAllSheets(Workbook caller)
        {
            foreach (Worksheet s in caller.Worksheets)
            {
                if (s.Visible != XlSheetVisibility.xlSheetVisible)
                {
                    s.Visible = XlSheetVisibility.xlSheetVisible;
                }
                s.Unprotect(worksheetPassword);
            }
        }

        public static Name getNamedRange(Workbook caller, Worksheet sh, string Name)
        {
            try
            {
                foreach (Name nm in sh.Names)
                {
                    if (nm.Name == sh.Name + "!" + Name)
                    {
                        return nm;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        public static Range getRange(Workbook caller, Worksheet sh, string Name)
        {
            Name nm = getNamedRange(caller: caller, sh: sh, Name: Name);
            if (nm != null)
            {
                try
                {
                    return nm.RefersToRange;
                }
                catch
                {
                    object obj = Globals.ThisAddIn.Application.ConvertFormula(nm.Name, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlR1C1);
                    return Globals.ThisAddIn.Application.Evaluate(obj);
                }
            }
            else
            {
                return null;
            }
        }

        public static Range getRange(Workbook caller, string Address)
        {
            object obj = Globals.ThisAddIn.Application.ConvertFormula(Address, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlR1C1);
            return Globals.ThisAddIn.Application.Evaluate(obj);
        }

        public static Worksheet GetWorksheet(string Name)
        {
            Worksheet ret = null;
            try
            {
                ret = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[Name];
            }
            catch { }
            return ret;
        }

        public static PivotTable getPivotTable(Worksheet sh, string Name)
        {
            PivotTable ret = null;
            foreach (PivotTable pt in sh.PivotTables())
            {
                if (pt.Name == Name) ret = pt;
            }

            return ret;
        }

        public static Shape GetShape(Worksheet sh, string Name)
        {
            Shape ret = null;
            foreach (Shape sha in sh.Shapes)
            {
                if (sha.Name == Name) ret = sha;
            }

            return ret;
        }

    }
}
