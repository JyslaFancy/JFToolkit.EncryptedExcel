using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// Extension methods for working with Excel sheets and cells
/// </summary>
public static class ExcelExtensions
{
    /// <summary>
    /// Gets a cell value as string, handling different cell types
    /// </summary>
    /// <param name="cell">The cell to read</param>
    /// <returns>String representation of the cell value</returns>
    public static string GetStringValue(this ICell? cell)
    {
        if (cell == null)
            return string.Empty;

        return cell.CellType switch
        {
            CellType.String => cell.StringCellValue ?? string.Empty,
            CellType.Numeric => DateUtil.IsCellDateFormatted(cell) 
                ? (cell.DateCellValue?.ToString() ?? string.Empty)
                : cell.NumericCellValue.ToString(),
            CellType.Boolean => cell.BooleanCellValue.ToString(),
            CellType.Formula => GetFormulaResultAsString(cell),
            CellType.Blank => string.Empty,
            _ => cell.ToString() ?? string.Empty
        };
    }

    /// <summary>
    /// Gets a cell value as double, returning 0 if not numeric
    /// </summary>
    /// <param name="cell">The cell to read</param>
    /// <returns>Numeric value of the cell</returns>
    public static double GetNumericValue(this ICell? cell)
    {
        if (cell == null)
            return 0;

        return cell.CellType switch
        {
            CellType.Numeric => cell.NumericCellValue,
            CellType.Formula when cell.CachedFormulaResultType == CellType.Numeric => cell.NumericCellValue,
            CellType.String when double.TryParse(cell.StringCellValue, out double result) => result,
            _ => 0
        };
    }

    /// <summary>
    /// Gets a cell value as DateTime, returning DateTime.MinValue if not a date
    /// </summary>
    /// <param name="cell">The cell to read</param>
    /// <returns>DateTime value of the cell</returns>
    public static DateTime GetDateTimeValue(this ICell? cell)
    {
        if (cell == null)
            return DateTime.MinValue;

        if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
        {
            return cell.DateCellValue ?? DateTime.MinValue;
        }

        if (cell.CellType == CellType.String && DateTime.TryParse(cell.StringCellValue, out DateTime result))
        {
            return result;
        }

        return DateTime.MinValue;
    }

    /// <summary>
    /// Gets a cell value as boolean
    /// </summary>
    /// <param name="cell">The cell to read</param>
    /// <returns>Boolean value of the cell</returns>
    public static bool GetBooleanValue(this ICell? cell)
    {
        if (cell == null)
            return false;

        return cell.CellType switch
        {
            CellType.Boolean => cell.BooleanCellValue,
            CellType.String => bool.TryParse(cell.StringCellValue, out bool result) && result,
            CellType.Numeric => cell.NumericCellValue != 0,
            CellType.Formula when cell.CachedFormulaResultType == CellType.Boolean => cell.BooleanCellValue,
            _ => false
        };
    }

    /// <summary>
    /// Gets a cell at the specified row and column coordinates
    /// </summary>
    /// <param name="sheet">The sheet to read from</param>
    /// <param name="rowIndex">Zero-based row index</param>
    /// <param name="columnIndex">Zero-based column index</param>
    /// <returns>ICell instance or null if not found</returns>
    public static ICell? GetCell(this ISheet sheet, int rowIndex, int columnIndex)
    {
        var row = sheet.GetRow(rowIndex);
        return row?.GetCell(columnIndex);
    }

    /// <summary>
    /// Gets a cell value as string at the specified coordinates
    /// </summary>
    /// <param name="sheet">The sheet to read from</param>
    /// <param name="rowIndex">Zero-based row index</param>
    /// <param name="columnIndex">Zero-based column index</param>
    /// <returns>String value of the cell</returns>
    public static string GetCellValue(this ISheet sheet, int rowIndex, int columnIndex)
    {
        return sheet.GetCell(rowIndex, columnIndex)?.GetStringValue() ?? string.Empty;
    }

    /// <summary>
    /// Gets all values from a row as string array
    /// </summary>
    /// <param name="row">The row to read</param>
    /// <returns>Array of string values</returns>
    public static string[] GetRowValues(this IRow? row)
    {
        if (row == null)
            return Array.Empty<string>();

        var values = new List<string>();
        for (int i = row.FirstCellNum; i < row.LastCellNum; i++)
        {
            var cell = row.GetCell(i);
            values.Add(cell?.GetStringValue() ?? string.Empty);
        }
        return values.ToArray();
    }

    /// <summary>
    /// Gets all values from a column as string array
    /// </summary>
    /// <param name="sheet">The sheet to read from</param>
    /// <param name="columnIndex">Zero-based column index</param>
    /// <param name="startRow">Starting row index (default: 0)</param>
    /// <param name="endRow">Ending row index (default: last row)</param>
    /// <returns>Array of string values</returns>
    public static string[] GetColumnValues(this ISheet sheet, int columnIndex, int startRow = 0, int? endRow = null)
    {
        var lastRowNum = endRow ?? sheet.LastRowNum;
        var values = new List<string>();

        for (int i = startRow; i <= lastRowNum; i++)
        {
            var cell = sheet.GetCell(i, columnIndex);
            values.Add(cell?.GetStringValue() ?? string.Empty);
        }

        return values.ToArray();
    }

    /// <summary>
    /// Converts the entire sheet to a 2D string array
    /// </summary>
    /// <param name="sheet">The sheet to convert</param>
    /// <returns>2D array of string values</returns>
    public static string[,] ToArray(this ISheet sheet)
    {
        if (sheet.LastRowNum == -1)
            return new string[0, 0];

        // Find the maximum column count
        int maxColumns = 0;
        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            if (row != null && row.LastCellNum > maxColumns)
                maxColumns = row.LastCellNum;
        }

        var result = new string[sheet.LastRowNum + 1, maxColumns];

        for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            for (int colIndex = 0; colIndex < maxColumns; colIndex++)
            {
                var cell = row?.GetCell(colIndex);
                result[rowIndex, colIndex] = cell?.GetStringValue() ?? string.Empty;
            }
        }

        return result;
    }

    /// <summary>
    /// Sets a cell value in the sheet, creating the row and cell if they don't exist
    /// </summary>
    /// <param name="sheet">The sheet to modify</param>
    /// <param name="rowIndex">Zero-based row index</param>
    /// <param name="columnIndex">Zero-based column index</param>
    /// <param name="value">Value to set</param>
    public static void SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, object? value)
    {
        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        
        switch (value)
        {
            case null:
                cell.SetCellType(CellType.Blank);
                break;
            case string stringValue:
                cell.SetCellValue(stringValue);
                break;
            case int intValue:
                cell.SetCellValue(intValue);
                break;
            case double doubleValue:
                cell.SetCellValue(doubleValue);
                break;
            case float floatValue:
                cell.SetCellValue(floatValue);
                break;
            case decimal decimalValue:
                cell.SetCellValue((double)decimalValue);
                break;
            case DateTime dateValue:
                cell.SetCellValue(dateValue);
                break;
            case bool boolValue:
                cell.SetCellValue(boolValue);
                break;
            default:
                cell.SetCellValue(value.ToString());
                break;
        }
    }

    /// <summary>
    /// Sets values for an entire row
    /// </summary>
    /// <param name="sheet">The sheet to modify</param>
    /// <param name="rowIndex">Zero-based row index</param>
    /// <param name="values">Values to set in the row</param>
    public static void SetRowValues(this ISheet sheet, int rowIndex, params object?[] values)
    {
        for (int i = 0; i < values.Length; i++)
        {
            sheet.SetCellValue(rowIndex, i, values[i]);
        }
    }

    /// <summary>
    /// Adds a new row with the specified values
    /// </summary>
    /// <param name="sheet">The sheet to modify</param>
    /// <param name="values">Values to add in the new row</param>
    /// <returns>The index of the created row</returns>
    public static int AddRow(this ISheet sheet, params object?[] values)
    {
        int newRowIndex = sheet.LastRowNum + 1;
        sheet.SetRowValues(newRowIndex, values);
        return newRowIndex;
    }

    /// <summary>
    /// Saves the workbook to a file without encryption
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="filePath">Path where to save the file</param>
    public static void SaveToFile(this IWorkbook workbook, string filePath)
    {
        EncryptedExcelWriter.SaveToFile(workbook, filePath);
    }

    /// <summary>
    /// Saves the workbook as encrypted Excel file
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="filePath">Path where to save the file</param>
    /// <param name="password">Password to encrypt the file with</param>
    public static void SaveEncryptedToFile(this IWorkbook workbook, string filePath, string password)
    {
        EncryptedExcelWriter.SaveEncryptedToFile(workbook, filePath, password);
    }

    /// <summary>
    /// Gets the workbook as a byte array without encryption
    /// </summary>
    /// <param name="workbook">The workbook to convert</param>
    /// <returns>Byte array containing the Excel file</returns>
    public static byte[] ToByteArray(this IWorkbook workbook)
    {
        return EncryptedExcelWriter.ToByteArray(workbook);
    }

    /// <summary>
    /// Gets the workbook as an encrypted byte array
    /// </summary>
    /// <param name="workbook">The workbook to convert</param>
    /// <param name="password">Password to encrypt with</param>
    /// <returns>Byte array containing the encrypted Excel file</returns>
    public static byte[] ToEncryptedByteArray(this IWorkbook workbook, string password)
    {
        return EncryptedExcelWriter.ToEncryptedByteArray(workbook, password);
    }

    private static string GetFormulaResultAsString(ICell cell)
    {
        try
        {
            return cell.CachedFormulaResultType switch
            {
                CellType.String => cell.StringCellValue ?? string.Empty,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell) 
                    ? (cell.DateCellValue?.ToString() ?? string.Empty)
                    : cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                _ => cell.ToString() ?? string.Empty
            };
        }
        catch
        {
            return cell.ToString() ?? string.Empty;
        }
    }
}
