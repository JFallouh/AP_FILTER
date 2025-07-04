// Created by James Fallouh
// File: Program.cs
// Date: 2025-07-04
// Purpose: Split RL_NEW_PAYABLES_TLC_TM.XLS into P1 (digits-only IDINVC) and P2 (others),
//          preserving ALL sheets (filtering only Invoices & Invoice_Details) and keeping formatting.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace PayablesSplitter
{
    class Program
    {
        private const string SourcePath = @"\\fs01\Accounting\AP\RL_NEW_PAYABLES_TLC_TM.XLS";
        private const string DestFolder = @"\\fs01\Accounting\AP\TM_AP_EXPORT\";

        static void Main()
        {
            Directory.CreateDirectory(DestFolder);

            using var srcStream = File.OpenRead(SourcePath);
            var srcWb = new HSSFWorkbook(srcStream);

            var invSheet = srcWb.GetSheet("Invoices");
            var detSheet = srcWb.GetSheet("Invoice_Details");
            var invHdr   = invSheet.GetRow(0);
            var detHdr   = detSheet.GetRow(0);

            // Split invoices
            var invP1 = ExtractRows(invSheet, invHdr, numericOnly: true,  out var cntP1);
            var invP2 = ExtractRows(invSheet, invHdr, numericOnly: false, out var cntP2);

            // Filter details
            var detP1 = FilterDetails(detSheet, detHdr, cntP1);
            var detP2 = FilterDetails(detSheet, detHdr, cntP2);

            // Build & save
            var wbP1 = CreateWorkbook(srcWb, invP1, invHdr, detP1, detHdr);
            SaveWorkbookWithRetry(wbP1, Path.Combine(DestFolder, "RL_NEW_PAYABLES_TLC_TM_P1.xls"));

            var wbP2 = CreateWorkbook(srcWb, invP2, invHdr, detP2, detHdr);
            SaveWorkbookWithRetry(wbP2, Path.Combine(DestFolder, "RL_NEW_PAYABLES_TLC_TM_P2.xls"));
        }

        /* ------------------- helpers ------------------- */

        static List<IRow> ExtractRows(ISheet sheet, IRow header, bool numericOnly, out HashSet<string> outCnt)
        {
            int idCol  = header.Cells.Select((c,i)=>(c.StringCellValue,i))
                                     .First(p => p.StringCellValue == "IDINVC").i;
            int cntCol = header.Cells.Select((c,i)=>(c.StringCellValue,i))
                                     .First(p => p.StringCellValue == "CNTITEM").i;

            var rows   = new List<IRow> { header };
            outCnt     = new HashSet<string>();

            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                var idVal    = row.GetCell(idCol)?.ToString() ?? "";
                bool isDigit = idVal.All(char.IsDigit);

                if ((numericOnly && isDigit) || (!numericOnly && !isDigit))
                {
                    rows.Add(row);
                    outCnt.Add(row.GetCell(cntCol)?.ToString() ?? "");
                }
            }
            return rows;
        }

        static List<IRow> FilterDetails(ISheet sheet, IRow header, HashSet<string> allowedCnt)
        {
            int cntCol = header.Cells.Select((c,i)=>(c.StringCellValue,i))
                                     .First(p => p.StringCellValue == "CNTITEM").i;

            var rows = new List<IRow> { header };
            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                var cntVal = row.GetCell(cntCol)?.ToString() ?? "";
                if (allowedCnt.Contains(cntVal))
                    rows.Add(row);
            }
            return rows;
        }

        static HSSFWorkbook CreateWorkbook(
            HSSFWorkbook srcWb,
            List<IRow> invRows, IRow invHdr,
            List<IRow> detRows, IRow detHdr)
        {
            var wb = new HSSFWorkbook();

            CopySheet(wb, srcWb, "Invoices",        invRows, invHdr);
            CopySheet(wb, srcWb, "Invoice_Details", detRows, detHdr);

            // Copy other sheets unchanged
            foreach (ISheet src in srcWb)
            {
                if (src.SheetName is "Invoices" or "Invoice_Details") continue;

                var allRows = new List<IRow>();
                for (int r = 0; r <= src.LastRowNum; r++)
                    if (src.GetRow(r) is IRow row) allRows.Add(row);

                CopySheet(wb, srcWb, src.SheetName, allRows, allRows[0]);
            }
            return wb;
        }

        /* -----------  COPY SHEET WITH FORMATTING  ----------- */

        static void CopySheet(
            HSSFWorkbook destWb,
            HSSFWorkbook srcWb,
            string name,
            List<IRow> srcRows,
            IRow srcHeader)
        {
            var sheet = destWb.CreateSheet(name);

            // style & font caches
            var styleMap = new Dictionary<short, ICellStyle>();
            var fontMap  = new Dictionary<short, short>();

            // column widths
            for (int c = 0; c < srcHeader.LastCellNum; c++)
                sheet.SetColumnWidth(c, srcHeader.Sheet.GetColumnWidth(c));

            for (int r = 0; r < srcRows.Count; r++)
            {
                var srcRow = srcRows[r];
                var dstRow = sheet.CreateRow(r);
                dstRow.Height = srcRow.Height;

                for (int c = 0; c < srcHeader.LastCellNum; c++)
                {
                    var sc = srcRow.GetCell(c);
                    if (sc == null) continue;
                    var dc = dstRow.CreateCell(c, sc.CellType);

                    // copy value
                    switch (sc.CellType)
                    {
                        case CellType.String:  dc.SetCellValue(sc.StringCellValue);  break;
                        case CellType.Numeric: dc.SetCellValue(sc.NumericCellValue); break;
                        case CellType.Boolean: dc.SetCellValue(sc.BooleanCellValue); break;
                        case CellType.Formula: dc.SetCellFormula(sc.CellFormula);    break;
                        default:               dc.SetCellValue(sc.ToString());      break;
                    }

                    // copy style
                    var srcStyleIdx = sc.CellStyle.Index;
                    if (!styleMap.TryGetValue(srcStyleIdx, out var dstStyle))
                    {
                        dstStyle = destWb.CreateCellStyle();
                        dstStyle.CloneStyleFrom(sc.CellStyle);

                        // clone font
                        short srcFontIdx = sc.CellStyle.FontIndex;
                        if (!fontMap.TryGetValue(srcFontIdx, out var dstFontIdx))
                        {
                            var srcFont = srcWb.GetFontAt(srcFontIdx);
                            var dstFont = destWb.CreateFont();
                            dstFont.Boldweight       = srcFont.Boldweight;
                            dstFont.Color            = srcFont.Color;
                            dstFont.FontHeight       = srcFont.FontHeight;
                            dstFont.FontName         = srcFont.FontName;
                            dstFont.IsItalic         = srcFont.IsItalic;
                            dstFont.Underline        = srcFont.Underline;
                            dstFontIdx               = dstFont.Index;
                            fontMap[srcFontIdx]      = dstFontIdx;
                        }
                        dstStyle.SetFont(destWb.GetFontAt(dstFontIdx));
                        styleMap[srcStyleIdx] = dstStyle;
                    }
                    dc.CellStyle = dstStyle;
                }
            }
        }

        /// <summary>
        /// Delete & rewrite; retry up to 5 × (1 s) if file is locked.
        /// </summary>
        static void SaveWorkbookWithRetry(HSSFWorkbook wb, string path)
        {
            const int maxAttempts = 5;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                    using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
                    wb.Write(fs);
                    return;
                }
                catch (IOException) when (attempt < maxAttempts)
                {
                    Thread.Sleep(1000);
                }
            }
        }
    }

    /* ------------ extension ------------- */
    static class RowExtensions
    {
        public static ICell? GetCell(this IRow row, string columnName)
        {
            var header = row.Sheet.GetRow(0);
            for (int c = 0; c < header.LastCellNum; c++)
                if (header.GetCell(c)?.StringCellValue == columnName)
                    return row.GetCell(c);
            return null;
        }
    }
}
