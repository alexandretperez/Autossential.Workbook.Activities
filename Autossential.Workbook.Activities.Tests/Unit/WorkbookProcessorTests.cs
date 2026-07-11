using Autossential.Workbook.Activities.Core;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;

namespace Autossential.Workbook.Activities.Tests.Unit
{
    public class WorkbookProcessorTests : BaseTests
    {
        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task Dispose_CreateFile_WhenContainsWriteOperations(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.WriteCell("Sheet1", "A1", "Hello");
            processor.Dispose();
            await Assert.That(File.Exists(filePath)).IsTrue();
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task Dispose_DoesNotCreateFile_WhenNoWriteOperations(string extension)
        {
            var (processor, filePath) = NewFile(extension);
            processor.Dispose();
            await Assert.That(File.Exists(filePath)).IsFalse();
        }

        [Test]
        [Arguments(0, 0, "A1", 0)]
        [Arguments(10, 1, "A1", 9)]
        [Arguments(12, 2, "A2", 12)]
        [Arguments(14, 3, "C3", 12)]
        public async Task GetColumnCount_ReturnsExpectedCount_BasedOnStartingCell(int cols, int rows, string startingCell, int expectedCols)
        {
            var data = TableUtils.Build(cols, rows, (col, row) =>
            {
                var value = $"C{col}R{row}";

                if (col <= 3)
                    return value;

                if (col % 3 == 0)
                    return value;

                return col % 2 == 0 ? null : "";
            });

            var (processor, _) = NewFile(".xls");
            processor.WriteRange("Sheet1", data, startingCell, false);

            var count = processor.GetColumnCount("Sheet1", startingCell);
            await Assert.That(count).IsEqualTo(expectedCols);
        }

        [Test]

        [Arguments(".xls", "A1", "A1", 8)]
        [Arguments(".xls", "A1", "A2", 8)]
        [Arguments(".xls", "A1", "A1:J1", 5)]
        [Arguments(".xls", "A1", "B2:J2", 5)]
        [Arguments(".xls", "A1", "D1:G3", 4)]
        [Arguments(".xls", "A1", "H3", 1)]
        [Arguments(".xls", "A1", "I3", 0)]
        [Arguments(".xls", "A2", "A1", 8)]

        [Arguments(".xlsx", "A1", "A1", 8)]
        [Arguments(".xlsx", "A1", "A2", 8)]
        [Arguments(".xlsx", "A1", "A1:J1", 5)]
        [Arguments(".xlsx", "A1", "B2:J2", 5)]
        [Arguments(".xlsx", "A1", "D1:G3", 4)]
        [Arguments(".xlsx", "A1", "H3", 1)]
        [Arguments(".xlsx", "A1", "I3", 0)]
        [Arguments(".xlsx", "A2", "A1", 8)]
        public async Task GetColumnCount_ReturnsExpectedCount_BasedOnRange(string extension, string writeCell, string range, int expectedCount)
        {
            var data = TableUtils.Build(10, 3, (col, row) =>
            {
                var value = $"C{col}R{row}";

                return row switch
                {
                    1 => col == 1 || col == 3 || col == 5 ? value : DBNull.Value,
                    2 => col == 6 ? value : string.Empty,
                    3 => col == 7 || col == 8 ? value : null,
                    _ => null
                };
            });

            var (processor, _) = NewFile(extension);
            processor.WriteRange("Sheet1", data, writeCell, false);
            await Assert.That(processor.GetColumnCount("Sheet1", range)).IsEqualTo(expectedCount);
        }


        [Test]
        [Arguments(0, 0, "A1", 0)]
        [Arguments(1, 10, "A1", 9)]
        [Arguments(1, 12, "A2", 12)]
        [Arguments(1, 14, "C3", 12)]
        public async Task GetRowCount_ReturnsExpectedCount_BasedOnStartingCell(int cols, int rows, string startingCell, int expectedRows)
        {
            var data = TableUtils.Build(cols, rows, (col, row) =>
            {
                var value = $"C{col}R{row}";

                if (row <= 3)
                    return value;

                if (row % 3 == 0)
                    return value;

                return row % 2 == 0 ? null : "";
            });

            var (processor, _) = NewFile(".xls");
            processor.WriteRange("Sheet1", data, startingCell, false);
            var count = processor.GetRowCount("Sheet1", startingCell);
            await Assert.That(count).IsEqualTo(expectedRows);
        }

        [Test]

        [Arguments(".xls", "A1", 8)]
        [Arguments(".xls", "A1:A10", 5)]
        [Arguments(".xls", "B1:C4", 0)]
        [Arguments(".xls", "B4:B6", 3)]

        [Arguments(".xlsx", "A1", 8)]
        [Arguments(".xlsx", "A1:A10", 5)]
        [Arguments(".xlsx", "B1:C4", 0)]
        [Arguments(".xlsx", "B4:B6", 3)]
        public async Task GetRowCount_ReturnsExpectedCount_BasedOnRange(string extension, string range, int expectedCount)
        {
            var data = TableUtils.Build(3, 10, (col, row) =>
            {
                var value = $"C{col}R{row}";

                return col switch
                {
                    1 => row == 1 || row == 3 || row == 5 ? value : DBNull.Value,
                    2 => row == 6 ? value : string.Empty,
                    3 => row == 7 || row == 8 ? value : null,
                    _ => null
                };
            });

            var (processor, _) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            await Assert.That(processor.GetRowCount("Sheet1", range)).IsEqualTo(expectedCount);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task GetSheetNames_ContainsOnlyTargetSheet_AfterWriteCell(string extension)
        {
            var (processor, _) = NewFile(extension);
            processor.WriteCell("Data", "A1", 1);
            await Assert.That(processor.GetSheetNames()).IsEquivalentTo(["Data"]);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task GetSheetNames_ContainsOnlyTargetSheet_AfterWriteRange(string extension)
        {
            var (processor, _) = NewFile(extension);
            processor.WriteRange("Table", TableUtils.Build(3, 2), "A1", true);
            await Assert.That(processor.GetSheetNames()).IsEquivalentTo(["Table"]);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task GetSheetNames_ReturnsDefaultSheet_WhenNoWriteOperations(string extension)
        {
            var (processor, _) = NewFile(extension);
            var sheetNames = processor.GetSheetNames();
            await Assert.That(sheetNames).IsEquivalentTo(["Sheet1"]);
        }

        [Test]
        public async Task ReadRange_ReturnsEmptyTable_WhenSheetDoesNotExist()
        {
            var (processor, _) = NewFile(".xlsx");
            var readData = processor.ReadRange("Missing", "A1", true);
            await Assert.That(readData.Columns.Count).IsEqualTo(0);
            await Assert.That(readData.Rows.Count).IsEqualTo(0);
        }

        [Test]
        public async Task ReadRange_RemovesEmptyRowsAndColumns_AfterTrim()
        {
            var (processor, filePath) = NewFile(".xlsx");
            var data = TableUtils.Generate(5, 8, 2);
            for (int i = 0; i < 3; i++)
                data.Columns.Add($"NewCol{i}");

            for (int i = 0; i < 2; i++)
            {
                var dr = data.NewRow();
                dr[data.Columns.Count - 1] = "*";
                data.Rows.Add(dr);
            }

            processor.WriteRange("Sheet1", data, "A1", true);
            processor.Dispose();

            Corrupt(filePath);
            processor = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var readData = processor.ReadRange("Sheet1", "A1", true);

            await Assert.That(readData.Columns.Count).IsEqualTo(data.Columns.Count - 3);
            await Assert.That(readData.Rows.Count).IsEqualTo(data.Rows.Count - 2);
        }

        [Test]

        [Arguments(".xls", "A1", true, 1, 1, 10, 10)]
        [Arguments(".xls", "A2", false, 1, 1, 10, 10)]
        [Arguments(".xls", "A2:D9", false, 1, 1, 4, 8)]
        [Arguments(".xls", "B3:G8", false, 1, 1, 6, 6)]
        [Arguments(".xls", "C5:E7", false, 1, 1, 3, 3)]
        [Arguments(".xls", "F1", true, 3, 4, 5, 2)]
        [Arguments(".xls", "H1:J11", true, 1, 1, 3, 10)]
        [Arguments(".xls", "H2:I17", true, 1, 1, 2, 15)]

        [Arguments(".xlsx", "A1", true, 1, 1, 10, 10)]
        [Arguments(".xlsx", "A2", false, 1, 1, 10, 10)]
        [Arguments(".xlsx", "A2:D9", false, 1, 1, 4, 8)]
        [Arguments(".xlsx", "B3:G8", false, 1, 1, 6, 6)]
        [Arguments(".xlsx", "C5:E7", false, 1, 1, 3, 3)]
        [Arguments(".xlsx", "F1", true, 3, 4, 5, 2)]
        [Arguments(".xlsx", "H1:J11", true, 1, 1, 3, 10)]
        [Arguments(".xlsx", "H2:I17", true, 1, 1, 2, 15)]
        public async Task ReadRange_ReturnsExpectedData_BasedOnArguments(string extension, string range, bool hasHeaders, int headerRows, int rowsPerRecord, int expectedCols, int expectedRows)
        {
            var data = TableUtils.Generate(10, 10);
            var (processor, _) = NewFile(extension);

            processor.WriteRange("Sheet1", data, "A1", true);
            var readData = processor.ReadRange("Sheet1", range, hasHeaders, headerRows, rowsPerRecord);
            await Assert.That(readData.Columns.Count).IsEqualTo(expectedCols);
            await Assert.That(readData.Rows.Count).IsEqualTo(expectedRows);
        }

        [Test]

        [Arguments(".xls", "A1:C18", true, 2, 3, 3, 6)]
        [Arguments(".xls", "B2:C18", true, 2, 3, 2, 5)]
        [Arguments(".xls", "I3:L12", true, 3, 2, 4, 4)]

        [Arguments(".xlsx", "A1:C18", true, 2, 3, 3, 6)]
        [Arguments(".xlsx", "B2:C18", true, 2, 3, 2, 5)]
        [Arguments(".xlsx", "I3:L12", true, 3, 2, 4, 4)]
        public async Task ReadRange_ReturnsExpectedRange_WhenHeaderRowsAndRowsPerRecord(string extension, string range, bool hasHeaders, int headerRows, int rowsPerRecord, int expectedCols, int expectedRows)
        {
            var data = TableUtils.Build(10, 10);
            var (processor, _) = NewFile(extension);

            processor.WriteRange("Sheet1", data, "A1", true);

            var readData = processor.ReadRange("Sheet1", range, hasHeaders, headerRows, rowsPerRecord);
            await Assert.That(readData.Columns.Count).IsEqualTo(expectedCols);
            await Assert.That(readData.Rows.Count).IsEqualTo(expectedRows);
        }


        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task Save_WhenMultipleCalls_IgnoresSubsequent(string extension)
        {
            var (processor, filePath) = NewFile(extension);

            processor.Save();
            var first = File.GetLastWriteTime(filePath);
            await Task.Delay(50);

            processor.Save();
            var second = File.GetLastWriteTime(filePath);
            await Task.Delay(50);

            processor.Dispose(); // also calls Save
            var third = File.GetLastWriteTime(filePath);

            await Assert.That(first).IsEqualTo(second).And.IsEqualTo(third);
        }

        [Test]
        [Arguments(".xlsx", "A1", "Hello")]
        [Arguments(".xls", "A1", "Hello")]
        [Arguments(".xlsx", "C9", 100d)]
        [Arguments(".xls", "C9", 100d)]
        [Arguments(".xlsx", "B15", null)]
        [Arguments(".xls", "B15", null)]
        [Arguments(".xlsx", "A2", false)]
        [Arguments(".xls", "A2", false)]
        [Arguments(".xlsx", "E4", true)]
        [Arguments(".xls", "E4", true)]
        [Arguments(".xlsx", "H9", "")]
        [Arguments(".xls", "E1", "")]
        public async Task WriteAndReadCell_ValueMatches_AfterWriteAndRead(string extension, string address, object? value)
        {
            var (processor, _) = NewFile(extension);

            processor.WriteCell("shEET1", address, value);

            var readValue = processor.ReadCell("Sheet1", address);
            if (extension == ".xls" && value?.ToString() == string.Empty)
                await Assert.That(readValue).IsNull();
            else
                await Assert.That(readValue).IsEqualTo(value);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task WriteRange_ReturnsExpectedData_AfterWriteRange(string extension)
        {
            var (processor, _) = NewFile(extension);
            var data = TableUtils.Build(3, 3);

            processor.WriteRange("Data", data, "B2", true);

            var readData = processor.ReadRange("Data", "B2", true);
            var areEqual = TableUtils.AreTablesEqual(data, readData);
            await Assert.That(areEqual).IsTrue();
        }

        [Test]

        [Arguments(".xls", "A1", 0, 10)]
        [Arguments(".xls", "B4", 0, 6)]
        [Arguments(".xls", "C4", 0, 5)]
        [Arguments(".xls", "D1", 0, 0)]
        [Arguments(".xls", "B7", 0, 3)]

        [Arguments(".xls", "A1", 5, 5)]
        [Arguments(".xls", "B1", 5, 3)]

        [Arguments(".xlsx", "A1", 0, 10)]
        [Arguments(".xlsx", "B4", 0, 6)]
        [Arguments(".xlsx", "C4", 0, 5)]
        [Arguments(".xlsx", "D1", 0, 0)]
        [Arguments(".xlsx", "B7", 0, 3)]

        [Arguments(".xlsx", "A1", 5, 5)]
        [Arguments(".xlsx", "B1", 5, 3)]
        public async Task ReadColumn_ReturnsExpectedData_BaseOnSpecificStartingCell(string extension, string startingCell, int limit, int expectedCount)
        {
            var data = TableUtils.Build(4, 10, (col, row) =>
            {
                var value = $"C{col}R{row}";

                return col switch
                {
                    1 => value,
                    2 => row % 3 == 0 ? value : null,
                    3 => row % 4 == 0 ? value : null,
                    _ => null
                };
            });

            var (processor, _) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            var values = processor.ReadColumn("Sheet1", startingCell, limit);
            await Assert.That(values.Length).IsEqualTo(expectedCount);
        }

        [Test]

        [Arguments(".xls", "A1", 0, 10)]
        [Arguments(".xls", "A2", 0, 9)]
        [Arguments(".xls", "A3", 0, 8)]
        [Arguments(".xls", "C3", 0, 6)]
        [Arguments(".xls", "E2", 4, 2)]
        [Arguments(".xls", "E3", 0, 4)]
        [Arguments(".xls", "B3", 6, 3)]

        [Arguments(".xlsx", "A1", 0, 10)]
        [Arguments(".xlsx", "A2", 0, 9)]
        [Arguments(".xlsx", "A3", 0, 8)]
        [Arguments(".xlsx", "C3", 0, 6)]
        [Arguments(".xlsx", "E2", 4, 2)]
        [Arguments(".xlsx", "E3", 0, 4)]
        [Arguments(".xlsx", "B3", 6, 3)]
        public async Task ReadRow_ReturnsExpectedData_BasedOnSpecificStartingCell(string extension, string startingCell, int limit, int expectedCount)
        {
            var data = TableUtils.Build(10, 4, (col, row) =>
            {
                var value = $"C{col}R{row}";

                return row switch
                {
                    1 => value,
                    2 => col % 3 == 0 ? value : null,
                    3 => col % 4 == 0 ? value : null,
                    _ => null
                };
            });

            var (processor, _) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            var values = processor.ReadRow("Sheet1", startingCell, limit);
            await Assert.That(values.Length).IsEqualTo(expectedCount);
        }

        [Test]

        [Arguments(".xls", "A1", 32, "F3", 6, 3)]
        [Arguments(".xls", "E2", 32, "F3", 6, 3)]
        [Arguments(".xls", "H1", null, "I2", 9, 2)]
        [Arguments(".xls", "H1", "", "I2", 9, 2)]
        [Arguments(".xls", "A1", "Col7", "G1", 7, 1)]
        [Arguments(".xls", "A1", "C10R8", "J9", 10, 9)]
        [Arguments(".xls", "A1", "IamNotThere", "", -1, -1)]

        [Arguments(".xlsx", "A1", 32, "F3", 6, 3)]
        [Arguments(".xlsx", "E2", 32, "F3", 6, 3)]
        [Arguments(".xlsx", "H1", null, "I2", 9, 2)]
        [Arguments(".xlsx", "H1", "", "I2", 9, 2)]
        [Arguments(".xlsx", "A1", "Col7", "G1", 7, 1)]
        [Arguments(".xlsx", "A1", "C10R8", "J9", 10, 9)]
        [Arguments(".xlsx", "A1", "IamNotThere", "", -1, -1)]
        public async Task FindValue_ReturnsAddress_WhenValueExists(string extension, string range, object? value, string expectedAddress, int expectedCol, int expectedRow)
        {
            var data = TableUtils.Generate(10, 10, 42);
            var (processor, _) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", true);
            var (address, col, row) = processor.FindValue("Sheet1", range, value);
            await Assert.That(address).IsEqualTo(expectedAddress);
            await Assert.That(col).IsEqualTo(expectedCol);
            await Assert.That(row).IsEqualTo(expectedRow);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task InsertSheet_Inserts_WhenNameIsUnique(string extension)
        {
            var (processor, _) = NewFile(extension);
            processor.InsertSheet("Sheet2");
            processor.InsertSheet("Sheet3");
            processor.InsertSheet("Sheet4");

            var sheetNames = processor.GetSheetNames();
            await Assert.That(sheetNames).IsEquivalentTo(["Sheet1", "Sheet2", "Sheet3", "Sheet4"]);
        }

        [Test]
        [Arguments(".xlsx", 1)]
        [Arguments(".xlsx", 2)]
        [Arguments(".xlsx", 3)]

        [Arguments(".xls", 1)]
        [Arguments(".xls", 2)]
        [Arguments(".xls", 3)]
        public async Task InsertSheet_InsertsInSpecificPosition_WhenPositionIsInformed(string extension, int position)
        {
            var (processor, _) = NewFile(extension);
            processor.InsertSheet("Sheet2");
            processor.InsertSheet("Sheet3");
            processor.InsertSheet("Sheet4");

            processor.InsertSheet("Data", position);

            var sheetNames = processor.GetSheetNames();
            await Assert.That(sheetNames).Count().IsEqualTo(5);
            await Assert.That(sheetNames[position - 1]).IsEqualTo("Data");
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public void InsertSheet_ThrownsException_WhenNameAlreadyExists(string extension)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                var (processor, _) = NewFile(extension);
                processor.InsertSheet("Sheet2");
                processor.InsertSheet("Sheet2");
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task DeleteSheet_SheetIsDeleted_WhenExists(string extension)
        {
            var (processor, _) = NewFile(extension);
            processor.InsertSheet("Sheet2");
            processor.InsertSheet("Sheet3");
            processor.InsertSheet("Sheet4");
            processor.DeleteSheet("Sheet3");

            await Assert.That(processor.GetSheetNames()).IsEquivalentTo(["Sheet1", "Sheet2", "Sheet4"]);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public void DeleteSheet_ThrowsException_WhenOnlyOneSheetExists(string extension)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                var (processor, _) = NewFile(extension);
                processor.DeleteSheet("Sheet1");
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task RenameSheet_Renames_WhenSheetIsFound(string extension)
        {
            var (processor, _) = NewFile(extension);
            processor.RenameSheet("Sheet1", "Data");
            await Assert.That(processor.GetSheetNames()).IsEquivalentTo(["Data"]);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task RenameSheet_Renames_WhenNoOtherSheetExistsWithSameName(string extension)
        {
            var (processor, _) = NewFile(extension);
            processor.RenameSheet("Sheet1", "sheet1"); // lowercase
            await Assert.That(processor.GetSheetNames()).IsEquivalentTo(["sheet1"]);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public void RenameSheet_ThrowsException_WhenAnotherSheetExistsWithSameName(string extension)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                var (processor, _) = NewFile(extension);
                processor.InsertSheet("Sheet2");
                processor.InsertSheet("Sheet3");
                processor.RenameSheet("Sheet2", "Sheet1");
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public void RenameSheet_ThrowsException_WhenSheetDoesNotExists(string extension)
        {
            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                var (processor, f) = NewFile(extension);
                processor.RenameSheet("DoesNotExist", "Data");
            });
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task ReadRange_ReturnsCorrectHeaderAndRows_WhenHeaderRowsAndRowsPerRecord(string extension)
        {
            var data = TableUtils.Build(4, 12, (col, row) =>
            {
                if (row == 1)
                    return $"C{col}";

                if (row == 2)
                    return col == 2 || col == 4 ? $"C{col}" : "";

                if (row == 3)
                    return col == 3 || col == 4 ? $"C{col}" : null;

                if (row % 3 == 0)
                {
                    if (col == 2)
                        return $"C{col}R{row}";

                    return "";
                }

                if (col == 2 && (row == 5 || row == 10))
                    return "";

                return $"C{col}R{row}";
            });

            var (processor, f) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", false);
            var readData = processor.ReadRange("Sheet1", "A1", true, 3, 3);
            var cols = readData.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            await Assert.That(cols).IsEquivalentTo(["C1", "C2 C2", "C3 C3", "C4 C4 C4"]);

            await Assert.That(readData.Rows[0].ItemArray.Cast<string>())
                .IsEquivalentTo(["C1R4 C1R5", "C2R4 C2R6", "C3R4 C3R5", "C4R4 C4R5"]);

            await Assert.That(readData.Rows[1].ItemArray.Cast<string>())
                .IsEquivalentTo(["C1R7 C1R8", "C2R7 C2R8 C2R9", "C3R7 C3R8", "C4R7 C4R8"]);

            await Assert.That(readData.Rows[2].ItemArray.Cast<string>())
                .IsEquivalentTo(["C1R10 C1R11", "C2R11 C2R12", "C3R10 C3R11", "C4R10 C4R11"]);
        }

        [Test]
        [Arguments(".xlsx")]
        [Arguments(".xls")]
        public async Task HideUnhide_Sheets_Verify(string extension)
        {
            var (processor, file) = NewFile(extension);
            processor.InsertSheet("Sheet2");
            processor.InsertSheet("Sheet3");
            processor.HideSheet("Sheet2");
            processor.Save();
            
            var info = WorkbookInspector.Inspect(file, "Sheet2");
            await Assert.That(info.IsVisible).IsFalse();

            processor.UnhideSheet("Sheet2");
            processor.Save();

            info = WorkbookInspector.Inspect(file, "Sheet2");
            await Assert.That(info.IsVisible).IsTrue();
        }


        [Test]
        public async Task FreezePanes_Xlsx_Verify()
        {
            var (processor, file) = NewFile(".xlsx");
            var table = TableUtils.Generate(10, 10);
            processor.WriteRange("Sheet1", table, "A1", true);

            processor.FreezePanes("Sheet1", 2, 2);
            processor.Save();

            var info = WorkbookInspector.Inspect(file, "Sheet1");
            await Assert.That(info.IsFrozen).IsTrue();
            await Assert.That(info.ColsFrozen).IsEqualTo(2);
            await Assert.That(info.RowsFrozen).IsEqualTo(2);

            processor.FreezePanes("Sheet1", 0, 0);
            processor.Save();

            info = WorkbookInspector.Inspect(file, "Sheet1");
            await Assert.That(info.IsFrozen).IsFalse();
            await Assert.That(info.ColsFrozen).IsEqualTo(0);
            await Assert.That(info.RowsFrozen).IsEqualTo(0);
        }

        [Test]
        public async Task FreezePanes_Xls_Verify()
        {
            var (processor, file) = NewFile(".xls");
            var table = TableUtils.Generate(10, 10);
            processor.WriteRange("Sheet1", table, "A1", true);

            processor.FreezePanes("Sheet1", 2, 2);
            processor.Save();

            var info = WorkbookInspector.Inspect(file, "Sheet1");
            await Assert.That(info.IsFrozen).IsTrue();
            await Assert.That(info.ColsFrozen).IsEqualTo(2);
            await Assert.That(info.RowsFrozen).IsEqualTo(2);

            processor.FreezePanes("Sheet1", 0, 0);
            processor.Save();

            info = WorkbookInspector.Inspect(file, "Sheet1");
            await Assert.That(info.IsFrozen).IsFalse();
            await Assert.That(info.ColsFrozen).IsEqualTo(0);
            await Assert.That(info.RowsFrozen).IsEqualTo(0);
        }
    }
}