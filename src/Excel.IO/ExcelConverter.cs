// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
// ReSharper disable UnusedMember.Global

namespace Excel.IO;

/// <summary>
/// Converter that allows implementations of <see cref="IExcelRow "/> to be exported.
/// </summary>
public class ExcelConverter : IExcelConverter
{
    private static SpreadsheetDocument _GetDocument(Stream stream)
    {
        var spreadsheetDocument = SpreadsheetDocument.Open(stream, isEditable: true);

        return spreadsheetDocument.WorkbookPart == null ? SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook) : spreadsheetDocument;
    }

    public void Append(IExcelRow row, Stream outputStream)
    {
        using var spreadsheetDocument = _GetDocument(outputStream);
        Write([row], spreadsheetDocument);
    }

    /// <summary>
    /// Exports the given rows to an Excel workbook
    /// </summary>
    /// <param name="rows">The rows to write to the workbook. Each property will be written as a cell in the row.</param>
    /// <param name="outputStream">The stream to write the workbook to</param>
    public void Write(IEnumerable<IExcelRow> rows, Stream outputStream)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook);
        Write(rows, spreadsheetDocument);
    }

    /// <summary>
    /// Exports the given rows to an Excel workbook
    /// </summary>
    /// <param name="rows">The rows to write to the workbook. Each property will be written as a cell in the row.</param>
    /// <param name="path">The path to write the workbook to</param>
    public void Write(IEnumerable<IExcelRow> rows, string path)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        Write(rows, spreadsheetDocument);
    }

    public class Grouping<TKey, TElement> : List<TElement>, IGrouping<TKey, TElement>
    {
        public Grouping(TKey key) => Key = key;
        public Grouping(TKey key, int capacity) : base(capacity) => Key = key;
        public Grouping(TKey key, IEnumerable<TElement> collection) : base(collection) => Key = key;
        public TKey Key { get; }
    }

    public void Write(IEnumerable<object> rows, string path, string sheet)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var group = new Grouping<string, object>(sheet, rows);
        Write([group], spreadsheetDocument, []);
    }

    private static void Write(IEnumerable<IExcelRow> rows, SpreadsheetDocument spreadsheetDocument)
    {
        var rowsGroupedBySheet = rows.GroupBy(r => r.SheetName);
        Write(rowsGroupedBySheet, spreadsheetDocument, typeof(IExcelRow).GetProperties());
    }

    private static void Write(IEnumerable<IGrouping<string, object>> rowsGroupedBySheet, SpreadsheetDocument spreadsheetDocument , PropertyInfo[] propertiesToIgnore)
    {
        if (spreadsheetDocument.WorkbookPart == null)
        {
            var workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
        }

        var sheets = spreadsheetDocument.WorkbookPart!.Workbook.Sheets ?? spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

        

        uint sheetId = 1;

        foreach (var rowGroup in rowsGroupedBySheet)
        {
            SheetData sheetData;
            var headerWritten = false;
            uint rowIndex = 1;

            var existingSheet = sheets.ChildElements.OfType<Sheet>().FirstOrDefault(s => s.Name == rowGroup.Key);

            if (existingSheet == null)
            {
                var worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var relationshipIdPart = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart);
                var sheet = new Sheet { Id = relationshipIdPart, SheetId = sheetId, Name = rowGroup.Key };

                sheets.Append(sheet);
                sheetId++;

                sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            }
            else
            {
                var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(existingSheet.Id ?? throw new InvalidOperationException());
                sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();                                       

                // get the correct row to write to
                var lastSheetRow = sheetData!.ChildElements.OfType<Row>().Last();
                rowIndex = lastSheetRow.RowIndex + 1;
                headerWritten = true;
            }

            foreach (var row in rowGroup)
            {
                var sheetRow = new Row { RowIndex = new UInt32Value(rowIndex) };
                sheetData?.Append(sheetRow);

                var properties = row.GetType()
                    .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.CanWrite)
                    //.Concat(row.GetType()
                    //    .GetFields(BindingFlags.Public | BindingFlags.Instance)
                    //    .Cast<MemberInfo>())
                    .ToArray();

                //var properties = row.GetType().GetProperties();

                var validProperties = properties.Except(propertiesToIgnore, SimpleComparer.Instance).ToList();

                if (!headerWritten)
                {
                    WriteHeader(validProperties, sheetRow, row);

                    headerWritten = true;
                    rowIndex++;

                    sheetRow = new Row { RowIndex = new UInt32Value(rowIndex) };
                    sheetData?.Append(sheetRow);
                }

                WriteCells(validProperties, sheetRow, row);

                rowIndex++;
            }
        }
    }

    public IEnumerable<T> Read<T>(string path, string sheet) where T : new()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(path, false);
        return Read<T>(spreadsheetDocument, sheet);
    }

    /// <summary>
    /// Reads a known workbook format into a collection of IExcelRow implementations
    /// </summary>
    /// <typeparam name="T">Implementation of IExcelRow that specifies the sheet to read and the row headings to include</typeparam>
    /// <param name="path">Path on disk of the workbook</param>
    /// <returns>A collection of <typeparamref name="T"/></returns>
    public IEnumerable<T> Read<T>(string path) where T : IExcelRow, new()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(path, false);
        return Read<T>(spreadsheetDocument, new T().SheetName);
    }

    /// <summary>
    /// Reads a known workbook format into a collection of IExcelRow implementations
    /// </summary>
    /// <typeparam name="T">Implementation of IExcelRow that specifies the sheet to read and the row headings to include</typeparam>
    /// <param name="stream">Stream that represents the workbook</param>
    /// <returns>A collection of <typeparamref name="T"/></returns>
    public IEnumerable<T> Read<T>(Stream stream) where T : IExcelRow, new()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        return Read<T>(spreadsheetDocument, new T().SheetName);
    }

    private static IEnumerable<T> Read<T>(SpreadsheetDocument spreadsheetDocument, string sheetName) where T : new()
    {
        var toReturn = new List<T>();
        var workBookPart = spreadsheetDocument.WorkbookPart;

        if (workBookPart == null) return toReturn;

        foreach (var sheet in workBookPart.Workbook.Descendants<Sheet>())
        {
            if (workBookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart)
            {
                // the part was supposed to be here, but wasn't found :/
                continue;
            }

            if (sheet.Name!.HasValue && sheet.Name.Value!.Equals(sheetName))
            {
                toReturn.AddRange(ReadSheet<T>(worksheetPart));
            }
        }

        return toReturn;
    }

    private static List<T> ReadSheet<T>(WorksheetPart wsPart) where T : new()
    {
        var toReturn = new List<T>();

        // assume the first row contains column names
        var headerRow = true;
        var headers = new Dictionary<string, object>();

        foreach (var row in wsPart.Worksheet.Descendants<Row>())
        {
            // one instance of T per row
            var obj = new T();
            var properties = obj.GetType().GetProperties();

            foreach (var c in row.Elements<Cell>())
            {
                var column = c.GetColumn();
                var value = c.GetCellValue();

                if (headerRow)
                {
                    headers.Add(column, value);
                }
                else
                {
                    // look for a property on the T that matches the name (ignore SheetName)

                    if (!headers.TryGetValue(column, out var columnHeader)) continue;
                    var propertyInfo = properties.FirstOrDefault(p => p.ResolveToNameOrDisplayName().Equals(columnHeader.ToString(), StringComparison.OrdinalIgnoreCase));

                    if (propertyInfo == null) continue;
                    var t = propertyInfo.PropertyType;
                    t = Nullable.GetUnderlyingType(t) ?? t;

                    propertyInfo.SetValue(obj, t.IsEnum ? Enum.Parse(t, (string)value) : Convert.ChangeType(value, t));
                }
            }

            if (!headerRow) toReturn.Add(obj);

            headerRow = false;
        }

        return toReturn;
    }

    private static void WriteCells(IEnumerable<PropertyInfo> properties, Row sheetRow, object userRow)
    {
        var columnIndex = 0;

        foreach (var item in properties)
        {
            var result = _TryInsertExcelColumn(sheetRow, userRow, columnIndex, item, isHeader: false);

            if (result.IsExcelColumn)
            {
                columnIndex = result.ColumnIndex;
                continue;
            }

            var cellValue = item.GetValue(userRow);

            sheetRow.InsertAt(
                new Cell
                {
                    CellReference = sheetRow.GetCellReference(columnIndex + 1),
                    CellValue = new CellValue(cellValue == null ? string.Empty : $"{cellValue}"),
                    DataType = new EnumValue<CellValues>(ResolveCellType(item.PropertyType))
                },
                columnIndex);

            columnIndex++;
        }
    }

    private static CellValues ResolveCellType(Type propertyType)
    {
        var nullableType = Nullable.GetUnderlyingType(propertyType);

        if (nullableType != null) propertyType = Nullable.GetUnderlyingType(propertyType);

        // TODO: Support date? 
        return Type.GetTypeCode(propertyType) switch
        {
            TypeCode.Decimal or TypeCode.Double or TypeCode.Int16 or TypeCode.Int32 or TypeCode.Int64
                or TypeCode.UInt16 or TypeCode.UInt32 or TypeCode.UInt64 => CellValues.Number,
            _ => CellValues.String
        };
    }

    private static void WriteHeader(IEnumerable<PropertyInfo> properties, Row sheetRow, object userRow)
    {
        var columnIndex = 0;

        foreach (var item in properties)
        {
            var result = _TryInsertExcelColumn(sheetRow, userRow, columnIndex, item, isHeader: true);

            if (result.IsExcelColumn)
            {
                columnIndex = result.ColumnIndex;
                continue;
            }

            var headerName = item.Name;

            var displayNameAttr = item.GetCustomAttribute<System.ComponentModel.DisplayNameAttribute>(true);

            if (displayNameAttr != null) headerName = displayNameAttr.DisplayName;

            sheetRow.InsertAt(
                new Cell
                {
                    CellReference = sheetRow.GetCellReference(columnIndex + 1),
                    CellValue = new CellValue(headerName),
                    DataType = new EnumValue<CellValues>(CellValues.String)
                },
                columnIndex);

            columnIndex++;
        }
    }

    private static InsertExcelColumnResult _TryInsertExcelColumn(Row sheetRow, object row, int columnIndex, PropertyInfo item, bool isHeader)
    {
        var excelColumnsAttr = item.GetCustomAttribute<ExcelColumnsAttribute>(true);

        if (excelColumnsAttr == null) return InsertExcelColumnResult.IsNotExcelColumn;
        var dict = (IDictionary<string, string>)item.GetValue(row);

        if (dict == null) return InsertExcelColumnResult.IsNotExcelColumn;
        foreach (var kvp in dict)
        {
            sheetRow.InsertAt(
                new Cell
                {
                    CellReference = sheetRow.GetCellReference(columnIndex + 1),
                    CellValue = new CellValue(isHeader ?
                        kvp.Key :
                        kvp.Value ?? string.Empty),
                    DataType = new EnumValue<CellValues>(isHeader ?
                        CellValues.String :
                        ResolveCellType(item.PropertyType))
                },
                columnIndex);

            columnIndex++;
        }

        return new InsertExcelColumnResult { IsExcelColumn = true, ColumnIndex = columnIndex };

    }

    private class InsertExcelColumnResult
    {
        public static InsertExcelColumnResult IsNotExcelColumn { get; } = new() { IsExcelColumn = false };

        public int ColumnIndex { get; init; }

        public bool IsExcelColumn { get; init; }
    }

    private class SimpleComparer : IEqualityComparer<PropertyInfo>
    {
        static SimpleComparer()
        {
            Instance = new SimpleComparer();
        }

        public static SimpleComparer Instance { get; }

        public bool Equals(PropertyInfo x, PropertyInfo y)
        {
            if(x==null) return y== null;
            if (y == null) return false;
            return x.Name == y.Name;
        }

        public int GetHashCode(PropertyInfo obj)
        {
            // only care if the name of the property info matches
            return obj.Name.GetHashCode();
        }
    }
}