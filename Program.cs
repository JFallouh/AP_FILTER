// Created by James Fallouh
// File: Program.cs
// Date: 2025-07-04
// Purpose: Split RL_NEW_PAYABLES_TLC_TM.XLS into P1 (digits-only IDINVC) and P2 (others),
//          preserving ALL sheets (filtering only Invoices & Invoice_Details), and emit a log.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;                       // [Change]
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace PayablesSplitter
{
    class Program
    {
        private const string SourcePath = @"\\fs01\Accounting\AP\RL_NEW_PAYABLES_TLC_TM.XLS";
        private const string DestFolder  = @"\\fs01\Accounting\AP\TM_AP_EXPORT\";
        private const string LogPath     = DestFolder + "PayablesSplitter.log";

        static void Main()
        {
            Directory.CreateDirectory(DestFolder);
            using var log = new StreamWriter(LogPath, append: true);
            log.WriteLine($"=== Run at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
            try
            {
                log.WriteLine("Loading source workbook...");
                using var srcStream = File.OpenRead(SourcePath);
                var srcWb = new HSSFWorkbook(srcStream);

                log.WriteLine("Reading Invoices & Invoice_Details sheets...");
                var invSheet = srcWb.GetSheet("Invoices");
                var detSheet = srcWb.GetSheet("Invoice_Details");
                var invHdr   = invSheet.GetRow(0);
                var detHdr   = detSheet.GetRow(0);

                // Split Invoices
                var invP1    = ExtractRows(invSheet, invHdr, numericOnly: true,  out var cntP1);
                var invP2    = ExtractRows(invSheet, invHdr, numericOnly: false, out var cntP2);
                log.WriteLine($" P1 invoices: {invP1.Count-1} rows, P2 invoices: {invP2.Count-1} rows");

                // Filter Details
                var detP1    = FilterDetails(detSheet, detHdr, cntP1);
                var detP2    = FilterDetails(detSheet, detHdr, cntP2);
                log.WriteLine($" P1 details: {detP1.Count-1} rows, P2 details: {detP2.Count-1} rows");

                // Build & save
                var wbP1 = CreateWorkbook(srcWb, invP1, invHdr, detP1, detHdr);
                SaveWorkbookWithRetry(wbP1, Path.Combine(DestFolder, "RL_NEW_PAYABLES_TLC_TM_P1.xls"), log);

                var wbP2 = CreateWorkbook(srcWb, invP2, invHdr, detP2, detHdr);
                SaveWorkbookWithRetry(wbP2, Path.Combine(DestFolder, "RL_NEW_PAYABLES_TLC_TM_P2.xls"), log);

                log.WriteLine("Processing complete.\n");
            }
            catch (Exception ex)
            {
                log.WriteLine("ERROR: " + ex);
                throw;
            }
        }

        static List<IRow> ExtractRows(ISheet sheet, IRow header, bool numericOnly, out HashSet<string> outCnt)
        {
            int idCol  = header.Cells.Select((c,i)=>(c.StringCellValue,i))
                                     .First(p=>p.StringCellValue=="IDINVC").i;
            int cntCol = header.Cells.Select((c,i)=>(c.StringCellValue,i))
                                     .First(p=>p.StringCellValue=="CNTITEM").i;

            var rows = new List<IRow>{ header };
            outCnt = new HashSet<string>();

            for (int r=1; r<=sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row==null) continue;
                var idVal = row.GetCell(idCol)?.ToString() ?? "";
                bool isDigits = idVal.All(char.IsDigit);
                if ((numericOnly&&isDigits) || (!numericOnly&&!isDigits))
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
                                     .First(p=>p.StringCellValue=="CNTITEM").i;
            var rows = new List<IRow>{ header };
            for (int r=1; r<=sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row==null) continue;
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

            CopySheet(wb, "Invoices",        invRows, invHdr);
            CopySheet(wb, "Invoice_Details", detRows, detHdr);

            // copy all other sheets unchanged
            foreach (ISheet src in srcWb)
            {
                if (src.SheetName=="Invoices" || src.SheetName=="Invoice_Details") continue;

                var allRows = new List<IRow>();
                for (int r=0; r<=src.LastRowNum; r++)
                    if (src.GetRow(r) is IRow row)
                        allRows.Add(row);

                CopySheet(wb, src.SheetName, allRows, allRows[0]);
            }

            return wb;
        }

        static void CopySheet(HSSFWorkbook wb, string name, List<IRow> rows, IRow header)
        {
            var sheet = wb.CreateSheet(name);
            for (int i=0; i<rows.Count; i++)
            {
                var srcCellRow = rows[i];
                var dstRow     = sheet.CreateRow(i);
                for (int c=0; c<header.LastCellNum; c++)
                {
                    var srcCell = srcCellRow.GetCell(c);
                    if (srcCell==null) continue;
                    var dstCell = dstRow.CreateCell(c);
                    switch (srcCell.CellType)
                    {
                        case CellType.String:  dstCell.SetCellValue(srcCell.StringCellValue);  break;
                        case CellType.Numeric: dstCell.SetCellValue(srcCell.NumericCellValue); break;
                        case CellType.Boolean: dstCell.SetCellValue(srcCell.BooleanCellValue); break;
                        default:               dstCell.SetCellValue(srcCell.ToString());      break;
                    }
                }
            }
        }

        /// <summary>
        /// Tries up to 5 times to delete & rewrite the file, waiting 1s between attempts.
        /// </summary>
        static void SaveWorkbookWithRetry(HSSFWorkbook wb, string path, StreamWriter log)
        {
            const int maxAttempts = 5;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        log.WriteLine($"Attempt {attempt}: deleting existing file...");
                        File.Delete(path);
                    }

                    using var fs = new FileStream(
                        path,
                        FileMode.Create,
                        FileAccess.Write,
                        FileShare.Read
                    );
                    wb.Write(fs);
                    log.WriteLine($"Saved → {path}");
                    return;
                }
                catch (IOException ex)
                {
                    log.WriteLine($"Attempt {attempt} failed: {ex.Message}");
                    if (attempt == maxAttempts)
                        throw;               // rethrow after final failure
                    Thread.Sleep(1000);     // wait then retry
                }
            }
        }
    }

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
