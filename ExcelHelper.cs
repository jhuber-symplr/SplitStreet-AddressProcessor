using System.IO;
using ClosedXML.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Linq;
using System.Collections.Generic;

public static class ExcelHelper
{
    public static void ProcessXlsx(string filePath, IEnumerable<string> possibleHeaders, string splitAddressPattern)
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                var address1Column = FindAddress1Column(worksheet.Row(1), possibleHeaders);

                if (address1Column.HasValue)
                {
                    int lastColumn = worksheet.Row(1).LastCellUsed().Address.ColumnNumber + 1;
                    worksheet.Cell(1, lastColumn).Value = "ADDRESS 2";

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        string address1 = row.Cell(address1Column.Value).GetValue<string>();
                        var splitAddress = AddressSplitter.SplitAddress(address1, splitAddressPattern);

                        row.Cell(address1Column.Value).Value = splitAddress.Address1;
                        row.Cell(lastColumn).Value = splitAddress.Address2;
                    }
                }
            }
            workbook.Save();
        }
    }

    public static void ProcessXls(string filePath, IEnumerable<string> possibleHeaders, string splitAddressPattern)
    {
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
        {
            var workbook = new HSSFWorkbook(fs);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                var headerRow = sheet.GetRow(0);
                int address1ColumnIndex = FindAddress1Column(headerRow, possibleHeaders);

                if (address1ColumnIndex >= 0)
                {
                    int lastColumnIndex = headerRow.LastCellNum;
                    headerRow.CreateCell(lastColumnIndex).SetCellValue("ADDRESS 2");

                    for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                        var addressCell = row.GetCell(address1ColumnIndex);

                        if (addressCell != null && addressCell.CellType == CellType.String)
                        {
                            string address1 = addressCell.StringCellValue;
                            var splitAddress = AddressSplitter.SplitAddress(address1, splitAddressPattern);

                            addressCell.SetCellValue(splitAddress.Address1);
                            row.CreateCell(lastColumnIndex).SetCellValue(splitAddress.Address2);
                        }
                    }
                }
            }

            using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(outputStream);
            }
        }
    }

    private static int? FindAddress1Column(IXLRow headerRow, IEnumerable<string> possibleHeaders)
    {
        foreach (var cell in headerRow.CellsUsed())
        {
            if (possibleHeaders.Any(header => string.Equals(cell.GetString(), header, StringComparison.OrdinalIgnoreCase)))
            {
                return cell.Address.ColumnNumber;
            }
        }
        return null;
    }

    private static int FindAddress1Column(IRow headerRow, IEnumerable<string> possibleHeaders)
    {
        for (int j = 0; j < headerRow.LastCellNum; j++)
        {
            var cell = headerRow.GetCell(j);
            if (cell != null && cell.CellType == CellType.String)
            {
                if (possibleHeaders.Any(header => string.Equals(cell.StringCellValue, header, StringComparison.OrdinalIgnoreCase)))
                {
                    return j;
                }
            }
        }
        return -1;
    }
}
