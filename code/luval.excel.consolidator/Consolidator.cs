﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.excel.consolidator
{
    public class Consolidator
    {
        public void Execute(Options options, FileInfo outputFile, FileInfo[] fileNames)
        {
            using (var consolidatedPackage = new ExcelPackage(outputFile))
            {
                using (var consolidatedSheet = consolidatedPackage.Workbook.Worksheets.Add("CONSOLIDATED"))
                {
                    var consolidatedRow = options.DataStartRow + 1;
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        using (var excelFilePackage = new ExcelPackage(fileNames[i]))
                        {
                            using (var excelSheet = excelFilePackage.Workbook.Worksheets[1])
                            {
                                var eofCriteria = false;
                                var emptyRowCount = 1;
                                var excelRow = options.DataStartRow;
                                while(!eofCriteria)
                                {
                                    var isNull = IsNull(excelSheet.Cells[excelRow, options.DataStartColumn].Value);
                                    if (!isNull)
                                    {
                                        for (int col = options.DataStartColumn; col <= options.DataEndColumn; col++)
                                        {
                                            var original = excelSheet.Cells[excelRow, col];
                                            CopyCell(excelSheet.Cells[excelRow, col], consolidatedSheet.Cells[consolidatedRow, col]);
                                        }
                                        consolidatedRow++;
                                        excelRow++;
                                    }
                                    else
                                    {
                                        emptyRowCount++;
                                        eofCriteria = emptyRowCount > 4 && isNull;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private bool IsNull(object val)
        {
            if (val is string)
                return string.IsNullOrEmpty(Convert.ToString(val));
            return val == null || DBNull.Value.Equals(val);
        }

        private void CopyHeader(Options options, ExcelWorksheet destination, ExcelWorksheet original)
        {
            for (int row = options.HeaderStartRow; row <= options.HeaderEndRow; row++)
            {
                for (int col = options.DataStartColumn; col <= options.DataEndColumn; col++)
                {
                    CopyCell(original.Cells[row, col], destination.Cells[row, col]);
                }
            }
        }

        private void CopyCell(ExcelRange originalCell, ExcelRange destinationCell)
        {
            destinationCell.Value = originalCell.Value;

            destinationCell.AutoFilter = originalCell.AutoFilter;
            destinationCell.Address = originalCell.Address;
            destinationCell.Hyperlink = originalCell.Hyperlink;
            destinationCell.Merge = originalCell.Merge;
            destinationCell.StyleName = originalCell.StyleName;

            destinationCell.Style.Border = originalCell.Style.Border;
            destinationCell.Style.Fill = originalCell.Style.Fill;
            destinationCell.Style.Font = originalCell.Style.Font;
            destinationCell.Style.Hidden = originalCell.Style.Hidden;
            destinationCell.Style.HorizontalAlignment = originalCell.Style.HorizontalAlignment;
            destinationCell.Style.Indent = originalCell.Style.Indent;
            destinationCell.Style.Numberformat = originalCell.Style.Numberformat;
            destinationCell.Style.QuotePrefix = originalCell.Style.QuotePrefix;
            destinationCell.Style.ReadingOrder = originalCell.Style.ReadingOrder;
            destinationCell.Style.ShrinkToFit = originalCell.Style.ShrinkToFit;
            destinationCell.Style.TextRotation = originalCell.Style.TextRotation;
            destinationCell.Style.VerticalAlignment = originalCell.Style.VerticalAlignment;
            destinationCell.Style.WrapText = originalCell.Style.WrapText;
            destinationCell.Style.XfId = originalCell.Style.XfId;

            destinationCell.Formula = originalCell.Formula;
            destinationCell.FormulaR1C1 = originalCell.FormulaR1C1;
        }
    }
}