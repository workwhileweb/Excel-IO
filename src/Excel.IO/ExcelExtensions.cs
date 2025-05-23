﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Reflection;

namespace Excel.IO;

public static class ExcelExtensions
{
    private const char FirstLetter = 'A';

    public static string GetCellReference(this Row row, int columnIndex)
    {            
        var cellReference = string.Empty;

        while (columnIndex > 0)
        {
            var remainder = (columnIndex - 1) % 26;
            var letter = (char)(FirstLetter + remainder);
            cellReference = letter + cellReference;
            columnIndex = (columnIndex - 1) / 26;
        }

        return $"{cellReference}{row.RowIndex}";
    }

    /// <summary>
    /// Finds the column identifier for a given cell, ie: A
    /// </summary>
    /// <param name="cell">The Cell to find the column for</param>
    /// <returns>The column name</returns>
    public static string GetColumn(this Cell cell)
    {
        if (cell.CellReference == null) return string.Empty;
        if (cell.CellReference.Value==null) return string.Empty;

        var endIndex = 0;

        for (var i = 0; i < cell.CellReference.Value.Length; i++)
            if (char.IsLetter(cell.CellReference.Value[i])) endIndex = i + 1;

        return cell.CellReference.Value[..endIndex];
    }

    /// <summary>
    /// Returns the value for a given Cell, taking into account the shared string table
    /// </summary>
    /// <param name="cell">The Cell to get the value for</param>
    /// <returns>The value of the Cell or null</returns>
    public static object GetCellValue(this Cell cell)
    {
        if (cell == null) return null;

        var worksheet = cell.FindParentWorksheet();

        if (string.IsNullOrWhiteSpace(cell.DataType))
        {
            if (cell.StyleIndex == null || !cell.StyleIndex.HasValue)
            {
                // General
                return cell.CellFormula != null ? cell.CellValue?.Text.ReplaceDecimalSeparator() : cell.InnerText.ReplaceDecimalSeparator();
            }

            var document = worksheet.WorksheetPart?.OpenXmlPackage as SpreadsheetDocument;
            var styleSheet = document?.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
            var cellStyle = styleSheet?.CellFormats?.ChildElements[(int)cell.StyleIndex.Value];
            var formatId = ((CellFormat)cellStyle)?.NumberFormatId;

            if (cell.CellValue == null || formatId==null) return null;

            switch ((int)formatId.Value)
            {
                // Linked Cell
                case 0:
                    return cell.CellValue == null ? string.Empty : cell.CellValue.Text;

                // Numbers
                // TODO: Find out if only integers fall into this case, or if all numeric data types do as well
                case 1:
                    return cell.CellFormula != null ? cell.CellValue.Text.ReplaceDecimalSeparator() : cell.InnerText.ReplaceDecimalSeparator();

                // Percentage
                case 9:

                // Scientific Notation
                case 11:

                // Fraction
                case 10:
                case 12:
                    return float.Parse(cell.CellFormula != null ? cell.CellValue.Text.ReplaceDecimalSeparator() : cell.InnerText.ReplaceDecimalSeparator());

                // General
                case 44:
                    return cell.CellFormula != null ? cell.CellValue.Text.ReplaceDecimalSeparator() : cell.InnerText.ReplaceDecimalSeparator();

                // Text
                case 49:
                    return cell.CellFormula != null ? cell.CellValue.Text : cell.InnerText;

                // Date
                case 14:
                case 15:
                case 16:
                case 17:
                case 18:
                case 19:
                case 20:
                case 21:
                case 22:
                case 164:
                case 165:
                case 166:
                case 169:
                    cell.TryParseDate(out var date);
                    return date;

                // Phone Number
                // TODO: Format Phone Numbers
                case 168:
                    return cell.CellValue.Text;

                // Currency
                case 167:
                    return decimal.Parse(cell.CellFormula != null ? cell.CellValue.Text : cell.InnerText);
                default:
                    throw new NotImplementedException($"Format with ID {(int)formatId.Value} and value {cell.CellValue?.InnerText ?? cell.InnerText} wasn't handled and needs to be parsed to the right format!");
            }
        }

        if (cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStringTablePart = worksheet.FindSharedStringTablePart();

            if (sharedStringTablePart is not null)
            {
                return sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
            }
        }
        else if (cell.DataType.Value == CellValues.Boolean)
        {
            return cell.InnerText != "0";
        }

        return cell.CellFormula != null ? cell.CellValue?.Text : cell.InnerText;
    }

    public static bool TryParseDate(this Cell cell, out DateTime? date)
    {
        date = null;

        if (cell.StyleIndex == null || !cell.StyleIndex.HasValue) return false;

        // See SpreadsheetML Reference at 18.8.30 numFmt(Number Format) for more detail: www.ecma-international.org/publications/standards/Ecma-376.htm
        // Also note some Excel specific variations            
        // The standard defines built-in format ID 14: "mm-dd-yy"; 22: "m/d/yy h:mm"; 37: "#,##0 ;(#,##0)"; 38: "#,##0 ;[Red](#,##0)"; 39: "#,##0.00;(#,##0.00)"; 40: "#,##0.00;[Red](#,##0.00)"; 47: "mmss.0"; KOR fmt 55: "yyyy-mm-dd".
        // Excel defines built-in format ID 14: "m/d/yyyy"; 22: "m/d/yyyy h:mm"; 37: "#,##0_);(#,##0)"; 38: "#,##0_);[Red](#,##0)"; 39: "#,##0.00_);(#,##0.00)"; 40: "#,##0.00_);[Red](#,##0.00)"; 47: "mm:ss.0"; KOR fmt 55: "yyyy/mm/dd".

        if (cell.CellFormula != null && cell.CellValue!=null)
        {
            date = DateTime.FromOADate(double.Parse(cell.CellValue.Text.ReplaceDecimalSeparator()));
        }
        else if (string.IsNullOrWhiteSpace(cell.InnerText))
        {
            date = DateTime.FromOADate(2);
            return true;
        }
        else
        {
            date = DateTime.FromOADate(double.Parse(cell.InnerText.ReplaceDecimalSeparator()));
        }
        return true;
    }

    public static Worksheet FindParentWorksheet(this Cell cell)
    {
        var parent = cell.Parent;

        while (parent is { Parent: not null } &&
               parent.Parent != parent &&
               !parent.LocalName.Equals("worksheet", StringComparison.InvariantCultureIgnoreCase))
        {
            parent = parent.Parent;
        }

        if(parent==null) throw new Exception("Worksheet invalid");

        if (!parent.LocalName.Equals("worksheet", StringComparison.InvariantCultureIgnoreCase))
            throw new Exception("Worksheet invalid");

        return parent as Worksheet;
    }

    public static SharedStringTablePart FindSharedStringTablePart(this Worksheet worksheet)
    {
        var document = worksheet.WorksheetPart?.OpenXmlPackage as SpreadsheetDocument;

        return document?.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
    }

    public static string ResolveToNameOrDisplayName(this PropertyInfo item)
    {
        var displayNameAttr = item.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), true).Cast<System.ComponentModel.DisplayNameAttribute>().FirstOrDefault();

        return displayNameAttr != null ? displayNameAttr.DisplayName : item.Name;
    }
}