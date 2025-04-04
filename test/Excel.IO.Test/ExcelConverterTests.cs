// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using DocumentFormat.OpenXml.Spreadsheet;
using Excel.IO.Test.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace Excel.IO.Test;

public class ExcelConverterTests : IDisposable
{
    private readonly Stream _xlsxTestResource;

    public ExcelConverterTests()
    {
        var res =
            Assembly.GetExecutingAssembly().GetManifestResourceStream("Excel.IO.Test.Resources.test.xlsx");

        var sr = new StreamReader(res ?? throw new InvalidOperationException());
        _xlsxTestResource = sr.BaseStream;
    }

    public void Dispose()
    {
        _xlsxTestResource.Close();
        _xlsxTestResource.Dispose();
    }

    [Fact]
    public void ExcelConverter_Can_Write_A_Single_Sheet_Workbook()
    {
        var excelConverter = new ExcelConverter();

        const string sheetName = "Sheet1";
        var rows = new List<IExcelRow>();

        for (var i = 0; i < 100; i++)
        {
            var mockRow = new MockExcelRow
            {
                SheetName = sheetName
            };

            rows.Add(mockRow);
        }

        using var result = new MemoryStream();
        excelConverter.Write(rows, result);
        Assert.True(result.Length > 0);
    }

    [Fact]
    public void ExcelConverter_Can_Write_A_MultiSheet_Workbook()
    {
        var excelConverter = new ExcelConverter();
        var rows = new List<IExcelRow>();

        for (var i = 0; i < 100; i++)
        {
            var mockRow = new MockExcelRow
            {
                SheetName = $"Sheet{i}"
            };

            rows.Add(mockRow);
        }

        using var result = new MemoryStream();
        excelConverter.Write(rows, result);
        Assert.True(result.Length > 0);
    }

    [Fact]
    public void ExcelConverter_Can_Read_A_Single_Sheet_From_A_Workbook()
    {
        var excelConverter = new ExcelConverter();
        var result = excelConverter.Read<MockExcelRow3>(_xlsxTestResource);

        Assert.Equal(10, result.Count());
    }

    [Fact]
    public void ExcelConverter_Can_Read_Multiple_Sheets_From_A_Workbook()
    {
        var excelConverter = new ExcelConverter();
        var result1 = excelConverter.Read<MockExcelRow3>(_xlsxTestResource);
        var result2 = excelConverter.Read<MockExcelRow4>(_xlsxTestResource);
        var result3 = excelConverter.Read<MockExcelRow5>(_xlsxTestResource);

        Assert.Equal(10, result1.Count());
        Assert.Equal(4, result2.Count());
        Assert.Equal(10, result3.Count());
    }

    [Fact]
    public void ExcelConverter_Can_Read_A_Single_Row_From_A_Sheet()
    {
        var excelConverter = new ExcelConverter();
        var rows = (List<MockExcelRow5>)excelConverter.Read<MockExcelRow5>(_xlsxTestResource);

        Assert.NotEmpty(rows);
        rows[0].GetType().GetProperties().ToList().ForEach(property =>
        {
            Assert.NotNull(property.GetValue(rows[0]));
        });
    }

    [Fact]
    public void ExcelConverter_Can_Read_Multiple_Rows_From_Multiple_Sheet()
    {
        var excelConverter = new ExcelConverter();
        var rows = (List<MockExcelRow5>)excelConverter.Read<MockExcelRow5>(_xlsxTestResource);

        Assert.NotEmpty(rows);

        rows.ForEach(row =>
        {
            row.GetType().GetProperties().ToList().ForEach(property =>
            {
                Assert.NotNull(property.GetValue(row));
            });
        });
    }

    [Fact]
    public void ExcelConverter_Can_Read_Multiple_Rows_From_A_Sheet()
    {
        var excelConverter = new ExcelConverter();
        var sheets = new List<List<IExcelRow>>
        {
            excelConverter.Read<MockExcelRow3>(_xlsxTestResource).ToList<IExcelRow>(),
            excelConverter.Read<MockExcelRow4>(_xlsxTestResource).ToList<IExcelRow>(),
            excelConverter.Read<MockExcelRow5>(_xlsxTestResource).ToList<IExcelRow>()
        };

        sheets.ForEach(sheet =>
        {
            Assert.NotEmpty(sheet);

            sheet.ForEach(row =>
            {
                row.GetType().GetProperties().ToList().ForEach(property =>
                {
                    Assert.NotNull(property.GetValue(row));
                });
            });
        });
    }

    [Fact]
    public void Cell_References_Correctly_Increment_Column_Letters()
    {
        var row = new Row
        {
            RowIndex = 1
        };

        var expectedCells = new[] { "A1", "B1", "C1", "D1" };

        var actualCells = new List<string>();

        for (var i = 1; i < 5; i++)
        {
            var cellRef = row.GetCellReference(i);
            actualCells.Add(cellRef);
        }

        foreach (var expectedCell in expectedCells)
        {
            Assert.Equal(expectedCell, actualCells[Array.IndexOf(expectedCells, expectedCell)]);
        }
    }

    [Fact]
    public void Columns_27_And_28_Are_Handled_Correctly()
    {
        var row = new Row
        {
            RowIndex = 1
        };

        var cellRef = row.GetCellReference(27);
        Assert.Equal("AA1", cellRef);

        var cellRef2 = row.GetCellReference(28);
        Assert.Equal("AB1", cellRef2);
    }

    [Fact]
    public void Columns_53_And_54_Are_Handled_Correctly()
    {
        var row = new Row
        {
            RowIndex = 1
        };

        var cellRef = row.GetCellReference(53);
        Assert.Equal("BA1", cellRef);

        var cellRef2 = row.GetCellReference(54);
        Assert.Equal("BB1", cellRef2);
    }

    [Fact]
    public void Cell_References_Correct_Row_Number()
    {
        var row = new Row
        {
            RowIndex = 4
        };

        var cellRef = row.GetCellReference(1);

        Assert.Equal("A4", cellRef);
    }

    [Fact]
    public void Sheets_Written_Can_Be_Read()
    {
        var excelConverter = new ExcelConverter();
        var written = new[] 
        {
            new MockExcelRow3
            {
                Address = "123 Fake",
                FirstName = "John",
                LastName = "Doe",
                LastContact = DateTime.Now,
                CustomerId = 1,
                IsActive = true,
                Balance = 100.00m,
                Category = Category.CategoryA
            }
        };

        var tmpFile = Path.GetTempFileName();

        try
        {
            excelConverter.Write(written, tmpFile);

            var read = excelConverter.Read<MockExcelRow3>(tmpFile).ToList();
				
            Assert.Equal(written.Length, read.Count);
            Assert.Equal(written.First().Address, read.First().Address);
        }
        finally
        {
            File.Delete(tmpFile);
        }
    }

    [Fact]
    public void ExcelColumnsAttribute_Correctly_WriteColumns_For_Dictionary_Keys_And_Row_Values_For_Dictionary_Value()
    {
        var excelConverter = new ExcelConverter();
        var written = new[]
        {
            new MockExcelRow6
            {
                CustomProperties = new Dictionary<string, string>
                {
                    { "Key1", "Value1" },
                    { "Key2", "Value2" },
                    { "Key3", "Value3" }
                }
            }
        };

        var tmpFile = Path.GetTempFileName();

        try
        {
            excelConverter.Write(written, tmpFile);
                
            var read = excelConverter.Read<MockExcelRow6ExplicitProperties>(tmpFile).ToList();

            Assert.Equal(written.Length, read.Count);

            //it's implied that the header is being written correctly as the Key1, Key2, Key3 can only be read if the header is written correctly
            Assert.Equal(written.First().CustomProperties.First().Value, read.First().Key1);
            Assert.Equal(written.First().CustomProperties.Skip(1).First().Value, read.First().Key2);
            Assert.Equal(written.First().CustomProperties.Skip(2).First().Value, read.First().Key3);
        }
        finally
        {
            File.Delete(tmpFile);
        }
    }

    [Fact]
    public void Appending_Is_Successful_When_The_File_To_Append_To_Does_Not_Already_Exist()
    {
        var excelConverter = new ExcelConverter();
        var expected = new MockExcelRow6ExplicitProperties { Key1 = "a", Key2 = "b", Key3 = "c" };

        var tmpFile = Path.GetTempFileName();

        using (var fileStreamWrite = new FileStream(tmpFile, FileMode.Create))
        {
            excelConverter.Append(expected, fileStreamWrite);
        }

        using (var fileStreamRead = new FileStream(tmpFile, FileMode.Open))
        {
            var actual = excelConverter.Read<MockExcelRow6ExplicitProperties>(fileStreamRead).ToList();

            Assert.Equal(expected.Key1, actual.First().Key1);
            Assert.Equal(expected.Key2, actual.First().Key2);
            Assert.Equal(expected.Key3, actual.First().Key3);
        }
    }

    [Fact]
    public void Appending_Is_Successful_When_The_File_To_Append_To_Does_Already_Exist()
    {
        var excelConverter = new ExcelConverter();
        var expectedRow1 = new MockExcelRow6ExplicitProperties { Key1 = "a", Key2 = "b", Key3 = "c" };
        var expectedRow2 = new MockExcelRow6ExplicitProperties { Key1 = "d", Key2 = "e", Key3 = "f" };

        var tmpFile = Path.GetTempFileName();

        using (var fileStreamWrite1 = new FileStream(tmpFile, FileMode.Create))
        {
            excelConverter.Append(expectedRow1, fileStreamWrite1);
        }

        using (var fileStreamWrite2 = new FileStream(tmpFile, FileMode.Open))
        {
            excelConverter.Append(expectedRow2, fileStreamWrite2);
        }

        using (var fileStreamRead = new FileStream(tmpFile, FileMode.Open))
        {
            var rows = excelConverter.Read<MockExcelRow6ExplicitProperties>(fileStreamRead).ToList();
            var actualRow1 = rows.First();
            var actualRow2 = rows.Skip(1).First();

            Assert.Equal(expectedRow1.Key1, actualRow1.Key1);
            Assert.Equal(expectedRow1.Key2, actualRow1.Key2);
            Assert.Equal(expectedRow1.Key3, actualRow1.Key3);

            Assert.Equal(expectedRow2.Key1, actualRow2.Key1);
            Assert.Equal(expectedRow2.Key2, actualRow2.Key2);
            Assert.Equal(expectedRow2.Key3, actualRow2.Key3);
        }
    }
}