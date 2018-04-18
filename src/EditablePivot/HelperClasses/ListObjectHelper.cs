using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EditablePivot.BaseClasses
{
    public class ListObjectHelper
    {
        public const int rowPerProcess = 50000;

        public static ListObject CreateTable(Worksheet sh, string TableName, int Row, int Col)
        {
            ListObject tbl = getListObject(wb: Globals.ThisAddIn.Application.ActiveWorkbook, shName: sh.Name, tblName: TableName);
            if (tbl == null)
            {
                Range rng = sh.Cells[RowIndex: Row, ColumnIndex: Col];
                rng.Value = "column 1";
                tbl = sh.ListObjects.AddEx(SourceType: XlListObjectSourceType.xlSrcRange
                                         , Source: rng
                                         , XlListObjectHasHeaders: XlYesNoGuess.xlYes
                                         , TableStyleName: "TableStyleLight9");
                tbl.Name = TableName;
                tbl.ShowAutoFilter = false;
            }

            return tbl;
        }

        public static ListObject CreateTable(Worksheet sh
                                           , string TableName
                                           , int Row
                                           , int Col
                                           , string[] ColumnNames
                                           , string[] NumberFormat)
        {

            ListObject tbl = CreateTable(sh: sh, TableName: TableName, Row: Row, Col: Col);

            // create columns
            for (int i = ColumnNames.Length - tbl.ListColumns.Count; i > 0; i--)
            {
                tbl.ListColumns.Add();
            }

            // rename and format column
            for (int i = 0; i < ColumnNames.Length; i++)
            {
                tbl.ListColumns[i + 1].Name = ColumnNames[i];
                if (i < NumberFormat.Length && NumberFormat[i] != "")
                    tbl.ListColumns[i + 1].Range.EntireColumn.NumberFormat = NumberFormat[i];
            }

            return tbl;
        }

        public static void SortTable(ListObject tbl, List<string> columnNameList, List<bool> SortAcendingList = null)
        {
            if (columnNameList.Count == 0)
                return;

            if (SortAcendingList == null) SortAcendingList = new List<bool>();
            if (SortAcendingList.Count < columnNameList.Count)
            {
                for (int i = SortAcendingList.Count; i < columnNameList.Count; i++)
                    SortAcendingList.Add(true);
            }

            Sort s = tbl.Sort;
            s.SortFields.Clear();
            if (tbl.ListRows.Count > 0)
            {
                for (int i = 0; i < columnNameList.Count; i++)
                {
                    string colName = columnNameList[i];
                    int colIdx = GetColumnIndex(tbl, colName, true);
                    if (colIdx > 0)
                    {
                        s.SortFields.Add(Key: tbl.ListColumns[colIdx].DataBodyRange
                                       , SortOn: XlSortOn.xlSortOnValues
                                       , Order: SortAcendingList[i] ? XlSortOrder.xlAscending : XlSortOrder.xlDescending
                                       , DataOption: XlSortDataOption.xlSortNormal);
                    }
                    else
                    {
                        MessageBox.Show("Column " + colName + " does not exist in Table " + tbl.Name);
                    }
                }

                s.Header = XlYesNoGuess.xlYes;
                s.MatchCase = false;
                s.Orientation = XlSortOrientation.xlSortColumns;
                s.SortMethod = XlSortMethod.xlStroke;
                s.Apply();
            }
        }

        public static void SortTable(ListObject tbl, List<int> columnIdList)
        {
            if (columnIdList.Count == 0)
                return;

            Sort s = tbl.Sort;
            s.SortFields.Clear();
            if (tbl.ListRows.Count > 0)
            {
                foreach (int colId in columnIdList)
                {
                    if (colId > 0)
                    {
                        s.SortFields.Add(Key: tbl.ListColumns[colId].DataBodyRange
                                       , SortOn: XlSortOn.xlSortOnValues
                                       , Order: XlSortOrder.xlAscending
                                       , DataOption: XlSortDataOption.xlSortNormal);
                    }
                    else
                    {
                        MessageBox.Show("Column " + colId + " does not exist in Table " + tbl.Name);
                    }
                }

                s.Header = XlYesNoGuess.xlYes;
                s.MatchCase = false;
                s.Orientation = XlSortOrientation.xlSortColumns;
                s.SortMethod = XlSortMethod.xlStroke;
                s.Apply();
            }
        }

        public static void SortTable(ListObject tbl, int UpToColumnIndex)
        {
            if (UpToColumnIndex < 1)
                return;

            if (UpToColumnIndex > tbl.ListColumns.Count)
            {
                UpToColumnIndex = tbl.ListColumns.Count;
            }

            Sort s = tbl.Sort;
            s.SortFields.Clear();
            if (tbl.ListRows.Count > 0)
            {
                for (int i = 1; i <= UpToColumnIndex; i++)
                {
                    s.SortFields.Add(Key: tbl.ListColumns[i].DataBodyRange
                                   , SortOn: XlSortOn.xlSortOnValues
                                   , Order: XlSortOrder.xlAscending
                                   , DataOption: XlSortDataOption.xlSortNormal);
                }

                s.Header = XlYesNoGuess.xlYes;
                s.MatchCase = false;
                s.Orientation = XlSortOrientation.xlSortColumns;
                s.SortMethod = XlSortMethod.xlStroke;
                s.Apply();
            }
        }

        public static void ClearTable(ListObject table)
        {
            if (table.ListRows.Count == 0)
            {
                return;
            }

            // save protected state
            Worksheet sh = default(Worksheet);
            sh = table.Range.Worksheet;

            bool isProtected = false;
            bool showAutoFilter = false;
            isProtected = sh.ProtectContents;

            if (isProtected)
            {
                SheetRangeHelper.protectWorksheet(sh, false);
            }

            // this procedure clear all rows in a Table (ListObject) by deleting entirerows
            // to get better performance
            // King Sun 2011-11-14, if the list object AutoFilter is on, then turn it off before delete
            showAutoFilter = table.ShowAutoFilter;
            if (showAutoFilter)
                table.ShowAutoFilter = false;

            table.DataBodyRange.EntireRow.Delete();

            if (showAutoFilter)
                table.ShowAutoFilter = true;

            if (isProtected)
            {
                SheetRangeHelper.protectWorksheet(sh, true);
            }

            sh = null;

        }

        public static void ResetTable(ListObject dt)
        {
            // trigger Calculation
            ((_Worksheet)dt.Range.Worksheet).Calculate();
            // turn off AutoFilter
            if (dt.AutoFilter != null) dt.AutoFilter.ShowAllData();
            // turn off show total
            dt.ShowTotals = false;
        }

        public static bool SpecialCellsExists(ListObject table, XlCellType cellType)
        {
            try
            {
                long cnt = 0;
                cnt = table.DataBodyRange.SpecialCells(cellType).Rows.Row;
                return true;
            }
            catch
            {
                return false;
            }

        }

        public static int LookUpRowNum(ListObject table, string lookupColumn, object value, int matchType = 0)
        {
            // it is important to know the lookup column data type, wrong data type, for instance,
            // lookup integer in text column will result in NOT FOUND
            try
            {
                Range rng = table.ListColumns[lookupColumn].DataBodyRange;
                return (int)table.Application.WorksheetFunction.Match(value, rng, matchType);
            }
            catch
            {
                return -1;
            }
        }

        public static object LookupValue(ListObject table, string lookupColumn, object value, string returnColumn)
        {
            // it is important to know the lookup column data type, wrong data type, for instance,
            // lookup integer in text column will result in NOT FOUND
            long rowNum = LookUpRowNum(table, lookupColumn, value);

            if (rowNum > 0)
            {
                Range rng = table.ListColumns[returnColumn].DataBodyRange[rowNum, 1];
                object ret = rng.Value;
                //rng.Address(ReferenceStyle:=XlReferenceStyle.xlA1)

                return ret;
            }
            else
            {
                return null;
            }

        }

        public static long LookupFilteredRowNum(ListObject table, string filterColumn, object filterValue, string lookupColumn, object value)
        {
            // this function is to lookup RowNum from a table on the filtered result
            // filterColumn must be sorted
            long startRow = 0;
            long endRow = 0;
            Range rng = default(Range);
            startRow = LookUpRowNum(table, filterColumn, filterValue);
            if (startRow <= 0)
            {
                return -1;
            }
            endRow = LookUpRowNum(table, filterColumn, filterValue);
            rng = table.Application.Range[table.ListColumns[lookupColumn].DataBodyRange[startRow, 1], table.ListColumns[lookupColumn].DataBodyRange[endRow, 1]];

            try
            {
                long rowNum = (long)table.Application.WorksheetFunction.Match(value, rng, 0);
                if (rowNum > 0)
                {
                    return startRow + rowNum - 1;
                }
                else
                {
                    return -1;
                }
            }
            catch
            {
                return -1;
            }
        }

        public static object LookupFilteredValue(ListObject table, string filterColumn, object filterValue, string lookupColumn, object value, string returnColumn)
        {
            // this function is to lookup value from a table on the filtered result
            // filterColumn must be sorted
            long rowNum = 0;
            rowNum = LookupFilteredRowNum(table, filterColumn, filterValue, lookupColumn, value);
            if (rowNum > 0)
                return table.ListColumns[returnColumn].DataBodyRange[rowNum, 1].Value;
            return "-1";

        }

        public static int GetColumnIndex(ListObject table, string columnName, bool ignoreNotFound = false)
        {
            foreach (ListColumn col in table.ListColumns)
            {
                if (col.Name.ToUpper() == columnName.ToUpper()) return col.Index;
            }

            if (ignoreNotFound) return -1;
            else
            {
                string errMsg = "Cannot find Column " + columnName + " in table " + table.Name;
                System.Windows.Forms.MessageBox.Show(errMsg);
                throw new Exception(errMsg);
            }

            //try {
            //    // this function returns index of specified column in the table
            //    return table.ListColumns[columnName].Index;
            //} catch (Exception ex) {
            //    if (ignoreNotFound) {
            //        return -1;
            //    } else {
            //        System.Windows.Forms.MessageBox.Show("Cannot find Column " + columnName + " in table " + table.Name);
            //        throw ex;
            //    }
            //}
        }

        private static int GetColumnIndexOld(ListObject table, string columnName, bool ignoreNotFound = false)
        {
            try
            {
                // this function returns index of specified column in the table
                string rangeStr = null;
                rangeStr = "#tableName[[#Headers],[#columnName]]";

                // add check on the listSeparator, if not equal to ",", then change
                //    Dim listSeparator As String
                //    listSeparator = Application.International(xlListSeparator)
                //    If listSeparator <> "," Then
                //        rangeStr = Replace(rangeStr, ",", listSeparator)
                //    End If

                rangeStr = rangeStr.Replace("#tableName", table.Name);
                rangeStr = rangeStr.Replace("#columnName", columnName);

                return table.Range.Worksheet.Range[rangeStr].Column - table.Range.Column + 1;

            }
            catch (Exception ex)
            {
                if (ignoreNotFound)
                {
                    return -1;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Cannot find Column " + columnName + " in table " + table.Name);
                    throw ex;
                }
            }
        }

        public static void SetTableColumn(ListObject destTable, string destColumn, string formulaText)
        {
            if (destTable.ListRows.Count > 0)
            {
                // set the column format to General to enable formulaText
                string saveFmt = null;
                saveFmt = destTable.ListColumns[destColumn].DataBodyRange.NumberFormat.ToString();
                destTable.ListColumns[destColumn].DataBodyRange.NumberFormat = "General";

                Range dr = destTable.DataBodyRange;
                int colDest = 0;
                colDest = GetColumnIndex(destTable, destColumn);
                // King 2012-06-24 performance tuning
                dr.Columns[colDest].FormulaR1C1 = formulaText;
                dr.Columns[colDest] = dr.Columns[colDest].Value;

                // restore column format
                // King 2012-06-24 performance tuning
                if (saveFmt != "General")
                    destTable.ListColumns[destColumn].DataBodyRange.NumberFormat = saveFmt;
            }
        }

        public static void copyTableColumn(ListObject destTable
                                         , string destColumn
                                         , string lookupTable
                                         , string returnColumn
                                         , string matchColumn
                                         , string matchToColumn
                                         , int ignoreNA = 0)
        {
            if (destTable.ListRows.Count > 0)
            {
                string formulaTemp = null;
                string formulaText = null;
                if (ignoreNA > 0)
                {
                    formulaTemp = "=IFERROR(INDEX($LookupTable[$ReturnColumn]," + "MATCH($DestTable[[#This Row]," + "[$MatchColumn]],$LookupTable[$MatchToColumn],0)),\"\")";
                }
                else
                {
                    formulaTemp = "=INDEX($LookupTable[$ReturnColumn]," + "MATCH($DestTable[[#This Row]," + "[$MatchColumn]],$LookupTable[$MatchToColumn],0))";
                }

                formulaText = formulaTemp;
                formulaText = formulaText.Replace("$LookupTable", lookupTable);
                formulaText = formulaText.Replace("$ReturnColumn", returnColumn);
                formulaText = formulaText.Replace("$DestTable", destTable.Name);
                formulaText = formulaText.Replace("$MatchColumn", matchColumn);
                formulaText = formulaText.Replace("$MatchToColumn", matchToColumn);

                SetTableColumn(destTable, destColumn, formulaText);
            }
        }

        public static void copyTableColumns(Workbook caller, int stepNo = 0)
        {
            int colStepNo = 0;
            int colDestTable = 0;
            int colDestColumn = 0;
            int colMatchColumn = 0;
            int colLookupTable = 0;
            int colMatchToColumn = 0;
            int colReturnColumn = 0;
            int colIgnoreNA = 0;

            ListObject tblCO = default(ListObject);
            tblCO = getListObject(caller, SheetRangeHelper.tblCopyOptionsName);

            if (stepNo > 0)
            {
                colStepNo = GetColumnIndex(tblCO, "StepNo");
            }
            else
            {
                colStepNo = 1;
            }
            colDestTable = GetColumnIndex(tblCO, "DestTable");
            colDestColumn = GetColumnIndex(tblCO, "DestColumn");
            colMatchColumn = GetColumnIndex(tblCO, "MatchColumn");
            colLookupTable = GetColumnIndex(tblCO, "LookupTable");
            colMatchToColumn = GetColumnIndex(tblCO, "MatchToColumn");
            colReturnColumn = GetColumnIndex(tblCO, "ReturnColumn");
            colIgnoreNA = GetColumnIndex(tblCO, "IgnoreNA", true);

            ListObject destTable = default(ListObject);
            string destTableName = null;
            string destColumn = null;
            string lookupTableName = null;
            int ignoreNA = 0;
            foreach (ListRow r in tblCO.ListRows)
            {
                // check if the step no matches
                if (stepNo == 0 || (r.Range.Cells[1, colStepNo].Value == stepNo.ToString()))
                {
                    destTableName = r.Range.Cells[1, colDestTable].Text;
                    destTable = getListObject(caller, destTableName);
                    destColumn = r.Range.Cells.Cells[1, colDestColumn].Text;
                    lookupTableName = r.Range.Cells.Cells[1, colLookupTable].Text;
                    if (colIgnoreNA > 0)
                    {
                        ignoreNA = r.Range.Cells.Cells[1, colIgnoreNA].Text;
                    }
                    else
                    {
                        ignoreNA = 0;
                    }
                    if (!string.IsNullOrEmpty(lookupTableName))
                    {
                        copyTableColumn(destTable: destTable
                                      , destColumn: destColumn
                                      , lookupTable: r.Range.Cells.Cells[1, colLookupTable].Text
                                      , returnColumn: r.Range.Cells[1, colReturnColumn].Text
                                      , matchColumn: r.Range.Cells.Cells[1, colMatchColumn].Text
                                      , matchToColumn: r.Range.Cells.Cells[1, colMatchToColumn].text
                                      , ignoreNA: ignoreNA);
                    }

                }
            }
        }

        public static void addListRow(ListObject tbl)
        {
            // King Sun 2012-16-11
            // this procedure is created as a replacement of ListRows.Add function to improve performance
            bool showTotals = false;
            showTotals = tbl.ShowTotals;
            if (showTotals)
                tbl.ShowTotals = false;
            tbl.Range[tbl.ListRows.Count + 2, 1] = " ";
            if (showTotals)
                tbl.ShowTotals = true;
        }

        public static ListObject getListObject(Workbook wb, string tblName)
        {
            ListObject ret = null;
            int i = 0;
            Worksheet sh = default(Worksheet);
            for (i = 1; i <= wb.Worksheets.Count; i++)
            {
                try
                {
                    sh = wb.Worksheets[i];
                    ret = sh.ListObjects[tblName];
                    if (ret != null)
                    {
                        break; // TODO: might not be correct. Was : Exit For
                    }
                }
                catch
                {
                }
            }

            return ret;
        }

        public static ListObject getListObject(Workbook wb, string shName, string tblName)
        {
            ListObject ret = null;
            Worksheet sh = default(Worksheet);
            try
            {
                sh = wb.Worksheets[shName];
            }
            catch
            {
                return ret;
            }

            try
            {
                ret = sh.ListObjects[tblName];

            }
            catch
            {
            }

            return ret;
        }

        public static bool LoadTextToTable(string txt, Microsoft.Office.Interop.Excel.ListObject target)
        {
            // King Sun 2014-09-13 when data is big, it's not possible to separate the first line (OutOfMemory issue), so start with the first line
            // jump out if empty string
            if (string.IsNullOrEmpty(txt))
            {
                return true;
            }

            // this function is to load text into a target range instead of using clipboard
            // to avoid problems in PutInClipboard() and Paste()

            // save current application settings
            //XlCalculation saveCalc;
            //bool saveDisplayAlerts = false;
            //bool saveScreenUpdating = false;
            //bool saveEnableEvents = false;

            //saveCalc = target.Application.Calculation;
            //saveDisplayAlerts = target.Application.DisplayAlerts;
            //saveScreenUpdating = target.Application.ScreenUpdating;
            //saveEnableEvents = target.Application.EnableEvents;

            //// disable application auto settings to improve performance
            //target.Application.DisplayAlerts = false;
            //target.Application.Calculation = XlCalculation.xlCalculationManual;
            //target.Application.ScreenUpdating = false;
            //target.Application.EnableEvents = false;


            // King Sun 2012-06-25 performance tuning, use Variant to store the data then push back to target data range directly
            // reference http://msdn.microsoft.com/en-us/library/ff726673.aspx#xlFasterVBA
            //Dim sh As Worksheet
            //Dim r As Long, c As Long
            //Set sh = target.Worksheet
            //r = target.Row()
            //c = target.Column()

            // split to lines
            // King Sun 2013-02-28 bug fix: remove Cr from the end of the line
            // lines = txt.Split("\n".ToCharArray());
            string[] lines = System.Text.RegularExpressions.Regex.Split(txt, "\n");
            int rowCnt = lines.Length - 1;
            if (rowCnt < 1) return true;

            int maxRows = target.Range.Worksheet.Rows.Count - target.Range.Row;
            // Handles situations when the number of rows exceeds the Excel Limit (1048576 for Excel 2013)
            if (rowCnt > maxRows)
            {
                MessageBox.Show("Number of rows " + rowCnt.ToString() + " exceeds the Excel limit. Please be aware that data is incomplete.");
                rowCnt = maxRows;
            }

            // cols = lines[0].Split('\t');
            int colCnt = target.ListColumns.Count;

            //// resize the table to fit rows and columns
            if (rowCnt == 1)
            {
                target.ListRows.AddEx();
            }
            else
            {
                //Range rng = target.Range.Resize[rowCnt + 1, target.Range.Columns.Count];
                Range cell1 = target.Range.Worksheet.Cells[RowIndex: target.Range.Row, ColumnIndex: target.Range.Column];
                Range cell2 = target.Range.Worksheet.Cells[RowIndex: target.Range.Row + rowCnt, ColumnIndex: target.Range.Column + colCnt];
                Range rng = target.Range.Worksheet.Range[cell1, cell2];
                rng.ClearFormats();
                target.Resize(rng);
            }

            char[] trimChars = new char[] { '\r', '\n' };
            int rowPerProcess = ListObjectHelper.rowPerProcess;
            int loopCount = (int)Math.Ceiling((double)(rowCnt) / rowPerProcess) - 1;

            for (int l = 0; l <= loopCount; l++)
            {
                int StartRow = rowPerProcess * l + 1;
                int EndRow = StartRow + rowPerProcess - 1;
                if (EndRow > rowCnt) EndRow = rowCnt;
                int processRowCnt = EndRow - StartRow + 1;

                object[,] data = new object[processRowCnt, colCnt];
                // loop through lines
                for (int r = StartRow; r <= EndRow; r++)
                {
                    string[] d = lines[r].TrimEnd(trimChars).Split("\t".ToCharArray());
                    for (int c = 0; c < colCnt && c < d.Length; c++)
                    {
                        try
                        {
                            data[r - StartRow, c] = d[c];
                        }
                        catch
                        {
                            Logging.debugLog("r:" + r.ToString() + " c:" + c.ToString());
                        }
                    }
                }

                Range cell1 = target.DataBodyRange[RowIndex: StartRow, ColumnIndex: 1];
                Range cell2 = target.DataBodyRange[RowIndex: EndRow, ColumnIndex: colCnt];
                Range rng = target.Range.Worksheet.Range[cell1, cell2];

                if (processRowCnt == 1 && colCnt == 1) rng.Value = data[1, 1];
                else rng.Value = data;
            }

            //object[,] d = new object[lineCnt, colCnt];
            ////ReDim d()

            //long i = 0;
            //int j = 0;
            //for (i = 0; i <= lineCnt - 1; i++) {
            //    cols = lines[i].Split('\t');
            //    for (j = 0; j <= cols.Length - 1; j++) {
            //        d[i, j] = cols[j];
            //    }
            //}

            //// write back data to the target range
            //target.DataBodyRange.Value = d;

            //// restore application auto settings
            //target.Application.DisplayAlerts = saveDisplayAlerts;
            //target.Application.Calculation = saveCalc;
            //target.Application.ScreenUpdating = saveScreenUpdating;
            //target.Application.EnableEvents = saveEnableEvents;

            return false;
        }

        public static bool LoadTextToRange(string txt, Range target)
        {
            // jump out if empty string
            if (string.IsNullOrEmpty(txt))
            {
                return true;
            }

            // this function is to load text into a target range instead of using clipboard
            // to avoid problems in PutInClipboard() and Paste()

            // save current application settings
            XlCalculation saveCalc;
            bool saveDisplayAlerts = false;
            bool saveScreenUpdating = false;
            bool saveEnableEvents = false;

            saveCalc = target.Application.Calculation;
            saveDisplayAlerts = target.Application.DisplayAlerts;
            saveScreenUpdating = target.Application.ScreenUpdating;
            saveEnableEvents = target.Application.EnableEvents;

            // disable application auto settings to improve performance
            target.Application.DisplayAlerts = false;
            target.Application.Calculation = XlCalculation.xlCalculationManual;
            target.Application.ScreenUpdating = false;
            target.Application.EnableEvents = false;

            string[] lines = null;
            long lineCnt = 0;
            string[] cols = null;
            int colCnt = 0;

            Microsoft.Office.Interop.Excel.Worksheet sh = default(Microsoft.Office.Interop.Excel.Worksheet);
            long r = 0;
            long c = 0;
            sh = target.Worksheet;
            r = target.Row;
            c = target.Column;

            // split to lines
            lines = txt.Split("\n".ToCharArray());
            lineCnt = lines.Length;

            // Handles situations when the number of rows exceeds the Excel Limit (1048576 for Excel 2013)
            long maxrows = sh.Rows.Count - r;
            if (lineCnt > maxrows)
            {
                MessageBox.Show("Number of rows " + lineCnt.ToString() + " exceeds the Excel limit. Please be aware that data is incomplete.");
                lineCnt = maxrows;
            }

            // split the first line to columns
            cols = lines[0].Split('\t');
            colCnt = cols.Length;

            long i = 0;
            int j = 0;
            for (i = 0; i <= lineCnt - 1; i++)
            {
                cols = lines[i].Split('\t');
                c = target.Column;
                for (j = 0; j <= cols.Length - 1; j++)
                {
                    sh.Cells[r, c] = cols[j];
                    c = c + 1;
                }
                r = r + 1;
            }

            // restore application auto settings
            target.Application.DisplayAlerts = saveDisplayAlerts;
            target.Application.Calculation = saveCalc;
            target.Application.ScreenUpdating = saveScreenUpdating;
            target.Application.EnableEvents = saveEnableEvents;

            return false;
        }

        public static bool CopyPasteText(string txt, Range target)
        {
            bool ret = false;

            // save current application settings
            XlCalculation saveCalc = default(XlCalculation);
            bool saveDisplayAlerts = false;
            bool saveScreenUpdating = false;
            bool saveEnableEvents = false;

            saveCalc = target.Application.Calculation;
            saveDisplayAlerts = target.Application.DisplayAlerts;
            saveScreenUpdating = target.Application.ScreenUpdating;
            saveEnableEvents = target.Application.EnableEvents;

            // disable application auto settings to improve performance
            target.Application.DisplayAlerts = false;
            target.Application.Calculation = XlCalculation.xlCalculationManual;
            target.Application.ScreenUpdating = false;
            target.Application.EnableEvents = false;

            //Dim lineCnt = Split(txt, vbLf).Length
            // resize the table to fit rows and columns
            target.Application.EnableAutoComplete = false;
            //Dim rng As Range = target.ListObject.Range.Resize(lineCnt + 1, target.ListObject.ListColumns.Count)
            //target.ListObject.Resize(rng)

            try
            {
                Clipboard.SetText(txt);
                target.Worksheet.Paste(target);
                Clipboard.Clear();
                ret = true;
            }
            catch
            {
                Clipboard.Clear();
                ret = false;
            }

            target.Application.EnableAutoComplete = true;
            target.Application.DisplayAlerts = saveDisplayAlerts;
            target.Application.Calculation = saveCalc;
            target.Application.ScreenUpdating = saveScreenUpdating;
            target.Application.EnableEvents = saveEnableEvents;

            return ret;
        }

        // King Sun new function to read/write table range
        public static List<object[,]> ReadTableColumn(ListObject MyTable, string ColumnName)
        {
            List<object[,]> ret = new List<object[,]>();
            Worksheet sh = MyTable.Range.Worksheet;
            int colCnt = MyTable.ListColumns.Count;
            int rowCnt = MyTable.ListRows.Count;
            if (rowCnt == 0) return ret;

            int rowPerProcess = ListObjectHelper.rowPerProcess;
            int loopCount = (int)Math.Ceiling((double)rowCnt / rowPerProcess) - 1;

            for (int l = 0; l <= loopCount; l++)
            {
                int StartRow = rowPerProcess * l + 1;
                int EndRow = StartRow + rowPerProcess - 1;
                if (EndRow > rowCnt)
                    EndRow = rowCnt;
                int processRowCnt = EndRow - StartRow + 1;

                object[,] obj = new object[processRowCnt + 1, 2];

                Range cell1 = MyTable.ListColumns[ColumnName].DataBodyRange[RowIndex: StartRow];
                Range cell2 = MyTable.ListColumns[ColumnName].DataBodyRange[RowIndex: EndRow];
                Range rng = sh.Range[cell1, cell2];
                if (processRowCnt == 1) obj[1, 1] = rng.Value;
                else obj = rng.Value;

                ret.Add(obj);
            }
            return ret;
        }

        // King Sun new function to read/write table range
        public static List<object[,]> ReadTable(ListObject MyTable)
        {
            List<object[,]> ret = new List<object[,]>();
            Worksheet sh = MyTable.Range.Worksheet;
            int colCnt = MyTable.ListColumns.Count;
            int rowCnt = MyTable.ListRows.Count;
            if (rowCnt == 0) return ret;

            int rowPerProcess = ListObjectHelper.rowPerProcess;
            int loopCount = (int)Math.Ceiling((double)rowCnt / rowPerProcess) - 1;

            for (int l = 0; l <= loopCount; l++)
            {
                int StartRow = rowPerProcess * l + 1;
                int EndRow = StartRow + rowPerProcess - 1;
                if (EndRow > rowCnt)
                    EndRow = rowCnt;
                int processRowCnt = EndRow - StartRow + 1;

                object[,] obj = new object[processRowCnt + 1, colCnt];

                Range cell1 = MyTable.DataBodyRange[RowIndex: StartRow, ColumnIndex: 1];
                Range cell2 = MyTable.DataBodyRange[RowIndex: EndRow, ColumnIndex: colCnt];
                Range rng = sh.Range[cell1, cell2];
                if (processRowCnt == 1 && colCnt == 1) obj[1, 1] = rng.Value;
                else obj = rng.Value;

                ret.Add(obj);
            }
            return ret;
        }

        // King Sun new function to read/write table range
        public static bool WriteTable(ListObject MyTable, List<object[,]> Value)
        {
            Worksheet sh = MyTable.Range.Worksheet;
            int colCnt = MyTable.ListColumns.Count;
            int rowCnt = MyTable.ListRows.Count;
            if (rowCnt == 0) return true;

            int rowPerProcess = ListObjectHelper.rowPerProcess;
            int loopCount = (int)Math.Ceiling((double)rowCnt / rowPerProcess) - 1;

            for (int l = 0; l <= loopCount; l++)
            {
                int StartRow = rowPerProcess * l + 1;
                int EndRow = StartRow + rowPerProcess - 1;
                if (EndRow > rowCnt)
                    EndRow = rowCnt;
                int processRowCnt = EndRow - StartRow + 1;

                Range cell1 = MyTable.DataBodyRange[RowIndex: StartRow, ColumnIndex: 1];
                Range cell2 = MyTable.DataBodyRange[RowIndex: EndRow, ColumnIndex: colCnt];
                Range rng = sh.Range[cell1, cell2];

                if (processRowCnt == 1 && colCnt == 1) rng.Value = Value[l][1, 1];
                else rng.Value = Value[l];
            }
            return true;
        }

        // King Sun new function to append lines to a table
        public static bool AppendTable(ListObject MyTable, object[,] Value)
        {
            Worksheet sh = MyTable.Range.Worksheet;
            int colCnt = MyTable.ListColumns.Count;
            int originalRowCnt = MyTable.ListRows.Count;

            int rowCnt = Value.GetUpperBound(0);
            if (rowCnt == 0) return true;

            if (originalRowCnt == 0) MyTable.ListRows.AddEx();

            if (rowCnt > 0)
            {
                Range rng = MyTable.Range.Resize[originalRowCnt + rowCnt + 1, colCnt];
                MyTable.Resize(rng);
            }

            Range cell1 = MyTable.DataBodyRange[RowIndex: originalRowCnt + 1, ColumnIndex: 1];
            Range cell2 = MyTable.DataBodyRange[RowIndex: originalRowCnt + rowCnt, ColumnIndex: colCnt];
            {
                Range rng = sh.Range[cell1, cell2];

                if (rowCnt == 1 && colCnt == 1) rng.Value = Value[1, 1];
                else rng.Value = Value;
            }
            //}
            return true;
        }

        // King Sun new function to read/write table range
        public static bool WriteTableColumn(ListObject MyTable, string ColumnName, List<object[,]> Value)
        {
            Worksheet sh = MyTable.Range.Worksheet;
            int colCnt = MyTable.ListColumns.Count;
            int rowCnt = MyTable.ListRows.Count;
            if (rowCnt == 0) return true;

            int rowPerProcess = ListObjectHelper.rowPerProcess;
            int loopCount = (int)Math.Ceiling((double)rowCnt / rowPerProcess) - 1;

            for (int l = 0; l <= loopCount; l++)
            {
                int StartRow = rowPerProcess * l + 1;
                int EndRow = StartRow + rowPerProcess - 1;
                if (EndRow > rowCnt)
                    EndRow = rowCnt;
                int processRowCnt = EndRow - StartRow + 1;

                Range cell1 = MyTable.ListColumns[ColumnName].DataBodyRange[RowIndex: StartRow];
                Range cell2 = MyTable.ListColumns[ColumnName].DataBodyRange[RowIndex: EndRow];
                Range rng = sh.Range[cell1, cell2];

                if (processRowCnt == 1) rng.Value = Value[l][1, 1];
                else rng.Value = Value[l];
            }
            return true;
        }
    }
}
