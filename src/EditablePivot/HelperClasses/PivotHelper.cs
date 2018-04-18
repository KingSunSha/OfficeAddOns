using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EditablePivot.BaseClasses
{
    public class PivotHelper
    {
        public const string EditEnabledFlag = "EditEnabled";
        public const string EditDisabledFlag = "EditDisabled";

        public static bool IsPivotCell(Range range)
        {
            try
            {
                PivotCell pc = range.PivotCell;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool TryGetPivotTable(Range range, out PivotTable pt)
        {
            pt = null;
            try
            {
                pt = range.PivotTable;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void ConfigEditable(Range range) {
            PivotTable pt;
            if (TryGetPivotTable(range, out pt)) {
                Forms.frmConfig frm = new Forms.frmConfig();

                frm.Init(pt);
                frm.ShowDialog();

                if (frm.OK) {
                    pt.Tag = frm.EditEnabled ? EditEnabledFlag : EditDisabledFlag;
                    SetEditable(range);
                }

                frm.Close();
            }
        }

        public static void SetEditable(Range range)
        {
            PivotTable pt;
            if (TryGetPivotTable(range, out pt))
            {
                if (IsPivotValueCell(range) && pt.Tag == EditEnabledFlag)
                {
                    pt.EnableDataValueEditing = true;
                }
                else
                {
                    pt.EnableDataValueEditing = false;
                }
            }
        }

        public static void PostSheetChange(Range range) { 
            PivotTable pt;
            if (!TryGetPivotTable(range, out pt))
                return;

            if (range.Address == pt.TableRange2.Address)
            {
                // when pivot get refreshed
                try
                {
                    // pt.PivotCache().Refresh();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Fail to refresh the pivot table. " + Environment.NewLine
                                                       + ex.Message + Environment.NewLine + ex.StackTrace);
                }
            }
            else
            {
                Range effectiveRange = Globals.ThisAddIn.Application.Intersect(range, pt.DataBodyRange);
                List<PivotCell> cells = new List<PivotCell>();
                List<string> values = new List<string>();
                for (int i = 1; i <= effectiveRange.Count; i++)
                {
                    Range cell = effectiveRange[i];
                    if (cell.PivotCell != null)
                    {
                        cells.Add(cell.PivotCell);
                        if (cell.Value == null)
                        {
                            values.Add("");
                        }
                        else
                        {
                            values.Add(cell.Value.ToString());
                        }
                    }
                }
                ChangePivotCells(pt: pt, cells: cells , values: values);
            }
        }
        
        private static bool IsPivotValueCell(Range range)
        {
            try
            {
                PivotCell pc = range.PivotCell;
                if (pc.PivotCellType == XlPivotCellType.xlPivotCellValue)
                    return true;
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static void ChangePivotCells(PivotTable pt, List<PivotCell> cells, List<string> values)
        {
            Worksheet sh = pt.TableRange1.Worksheet;

            string sourceString = pt.PivotCache().SourceData;
            ListObject dt = ListObjectHelper.getListObject(Globals.ThisAddIn.Application.ActiveWorkbook, sourceString);

            if (dt == null) {
                MessageBox.Show("Datas source of the pivot table must be a Table");
                pt.PivotCache().Refresh();
                return;
                // Range sourceRange = SheetRangeHelper.getRange(Globals.ThisAddIn.Application.ActiveWorkbook, sourceString);
            }

            // loop through changed cells, apply changes
            for (int i = 0; i <= cells.Count - 1; i++)
            {
                if (!GeneralSettings.IsNumeric(values[i]) && !string.IsNullOrEmpty(values[i]))
                {
                    System.Windows.Forms.MessageBox.Show("Only numeric values are allowed.");
                }
                else
                {
                    ChangePivotCell(cells[i], values[i], dt);
                }
            }

            ListObjectHelper.ResetTable(dt);

        }

        private static void ChangePivotCell(PivotCell Cell, string Value, ListObject sourceDataTable)
        {
            // get source range for the pivot table
            string dataCol = null;
            int colIndex = 0;
            int writableColIndex = 0;
            dataCol = Cell.DataField.SourceName;
            colIndex = ListObjectHelper.GetColumnIndex(sourceDataTable, dataCol);

            int baseColIndex = colIndex;

            // go through the pivot items on the changed target
            List<string> filterColNames = new List<string>();
            List<int> filterColIndexes = new List<int>();
            List<object> filterValues = new List<object>();

            // look for Pivot Items the cell is linked to
            string pfIds = null;
            // a list of Pivot Field Ids have gone through, to be used when check additional filters
            pfIds = ":";
            
            // look in Column Items
            foreach (PivotItem pi in Cell.ColumnItems)
            {
                PivotField pf = pi.Parent;
                pfIds += pf.SourceName + ":";
                filterColNames.Add(pf.SourceName);
                filterColIndexes.Add(ListObjectHelper.GetColumnIndex(sourceDataTable, pf.SourceName));
                if (pi.SourceNameStandard == "(blank)")
                {
                    filterValues.Add("=");
                }
                else
                {
                    string val = pi.Name;
                    if (GeneralSettings.IsNumeric(val))
                    {
                        val = Math.Round(Convert.ToDouble(val), 8).ToString();
                    }
                    filterValues.Add(val);
                }
            }

            // look in Row Items
            foreach (PivotItem pi in Cell.RowItems)
            {
                PivotField pf = pi.Parent;
                pfIds += pf.SourceName + ":";
                filterColNames.Add(pf.SourceName);
                filterColIndexes.Add(ListObjectHelper.GetColumnIndex(sourceDataTable, pf.SourceName));
                if (pi.SourceNameStandard == "(blank)")
                {
                    filterValues.Add("=");
                }
                else
                {
                    string val = pi.Name;
                    if (GeneralSettings.IsNumeric(val))
                    {
                        val = Math.Round(Convert.ToDouble(val), 8).ToString();
                    }

                    filterValues.Add(val);
                }
            }

            // apply filters on the data source
            int i = 0;

            if (sourceDataTable.AutoFilter != null)
                sourceDataTable.AutoFilter.ShowAllData();

            for (i = 0; i <= filterColNames.Count - 1; i++)
            {
                sourceDataTable.Range.AutoFilter(Field: filterColIndexes[i], Criteria1: filterValues[i]);
            }

            // if nothing after filter, then exit
            if (!ListObjectHelper.SpecialCellsExists(sourceDataTable, XlCellType.xlCellTypeVisible))
            {
                ListObjectHelper.ResetTable(sourceDataTable);
                return;
            }
            else
            {
                //try to empty the field
                if (string.IsNullOrEmpty(Value))
                {
                    // loop through all visible rows and set value to ""
                    foreach (Range rowX in sourceDataTable.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows)
                    {
                        rowX.Cells[1, colIndex].Value = "";
                    }
                }
                else if (Value == "0")
                {
                    // loop through all visible rows and set value to ""
                    foreach (Range rowX in sourceDataTable.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows)
                    {
                        rowX.Cells[1, colIndex].Value = 0;
                    }
                }
                else
                {
                    // check if Total Range exist for the table
                    if (!sourceDataTable.ShowTotals)
                        sourceDataTable.ShowTotals = true;

                    double newTotal = 0;
                    newTotal = Math.Round(Convert.ToDouble(Value), 6);
                    // max decimal 6

                    // check total row count, if total row count is 1 then update directly
                    int ttlRowCnt = 0;
                    sourceDataTable.ListColumns[1].TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationCount;
                    ttlRowCnt = Convert.ToInt32(sourceDataTable.TotalsRowRange.Cells[1, 1].Value);

                    if (ttlRowCnt == 1)
                    {
                        // loop through all visible rows and update value
                        foreach (Range rowX in sourceDataTable.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows)
                        {
                            rowX.Cells[1, colIndex].Value = newTotal;
                        }
                    }
                    else
                    {
                        // sum up data source value, divided by new value, calculate change factor
                        double previousTotal = 0;
                        double multipleFactor = 0;
                        double addFactor = 0;
                        double baseTotal = 0;

                        // set calculation to SUM
                        sourceDataTable.ListColumns[colIndex].TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationSum;
                        sourceDataTable.ListColumns[baseColIndex].TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationSum;

                        // trigger Calculation
                        ((_Worksheet)sourceDataTable.Range.Worksheet).Calculate();
                        previousTotal = Math.Round(sourceDataTable.TotalsRowRange.Cells[1, colIndex].Value, 6);
                        baseTotal = Math.Round(sourceDataTable.TotalsRowRange.Cells[1, baseColIndex].Value, 6);

                        bool addToAll = true;
                        if (baseTotal > 0)
                        {
                            multipleFactor = newTotal / baseTotal;
                            addFactor = 0;
                            // when Base Total value = 0 then get item counts to split evenly
                        }
                        else
                        {
                            multipleFactor = 0;
                            sourceDataTable.ListColumns[colIndex].TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationCount;
                            sourceDataTable.ListColumns[baseColIndex].TotalsCalculation = XlTotalsCalculation.xlTotalsCalculationCount;
                            // trigger Calculation
                            ((_Worksheet)sourceDataTable.Range.Worksheet).Calculate();
                            previousTotal = sourceDataTable.TotalsRowRange.Cells[1, colIndex].Value;
                            baseTotal = sourceDataTable.TotalsRowRange.Cells[1, baseColIndex].Value;
                            if (baseTotal == 0)
                            {
                                // King Sun 2013-03-22  if all rows are empty, then take FCST_UNIT count to disaggregate
                                baseTotal = sourceDataTable.TotalsRowRange.Cells[1, 1].Value;
                            }
                            else
                            {
                                addToAll = false;
                            }
                            addFactor = newTotal / baseTotal;
                        }

                        double sumValue = 0.0;
                        // loop through all visible rows and update value
                        foreach (Range rowX in sourceDataTable.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows)
                        {
                            double cellValue = 0.0;
                            if (rowX.Cells[1, baseColIndex].Value != null) cellValue = rowX.Cells[1, baseColIndex].Value;

                            if (addToAll || rowX.Cells[1, colIndex].Value != null)
                            {
                                cellValue = Math.Round(cellValue * multipleFactor + addFactor, 6);
                                rowX.Cells[1, colIndex].Value = cellValue;
                                sumValue = sumValue + cellValue;
                            }
                        }

                        // adjust rounding difference
                        double diff = 0;
                        diff = sumValue - newTotal;
                        // if rounding causes mismatch of value, add the difference on the first possible cell
                        if (diff != 0)
                        {
                            foreach (Range rowX in sourceDataTable.DataBodyRange.SpecialCells(XlCellType.xlCellTypeVisible).Rows)
                            {
                                if (!string.IsNullOrEmpty(rowX.Cells[1, colIndex].Value.ToString()))
                                {
                                    if (rowX.Cells[1, colIndex].Value > diff)
                                    {
                                        rowX.Cells[1, colIndex].Value = rowX.Cells[1, colIndex].Value - diff;
                                        break; // TODO: might not be correct. Was : Exit For
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return;
        }


    }
}
