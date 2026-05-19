using NPOI.SS.UserModel;

namespace Autossential.Workbook.Activities.Tests
{
    public class WorkbookFixture : IDisposable
    {
        public WorkbookFixture()
        {
            static void FillAll(ISheet sheet)
            {
                for (int r = 0; r < 10; r++)
                {
                    var row = sheet.CreateRow(r);
                    for (int c = 0; c < 10; c++)
                    {
                        row.CreateCell(c).SetCellValue($"{r}_{c}");
                    }
                }
            }

            static void DiagonalFill(ISheet sheet)
            {
                for (int r = 0; r < 10; r++)
                {
                    var row = sheet.CreateRow(r);
                    for (int c = 0; c < 10; c++)
                    {
                        var cell = row.CreateCell(c);
                        if (r == c)
                            cell.SetCellValue($"{r}_{c}");
                    }
                }
            }

            static void PartialFill(ISheet sheet)
            {
                for (int r = 0; r < 10; r++)
                {
                    var row = sheet.CreateRow(r);
                    if ((r + 1) % 3 == 0)
                        continue;

                    for (int c = 0; c < 10; c++)
                    {
                        var cell = row.CreateCell(c);
                        if ((c + 1) % 2 == 0)
                            cell.SetCellValue($"{r}_{c}");
                    }
                }
            }

            static void Table(ISheet sheet)
            {
                var recordIndex = 0;
                var rowsPerRecord = 3;
                for (int r = 0; r <= 30; r++)
                {
                    if (r < 2)
                        continue;

                    var row = sheet.CreateRow(r);
                    for (int c = 0; c < 5; c++)
                    {
                        row.CreateCell(c).SetCellValue($"{r}_{c}");
                    }

                    if ((r + 1) % rowsPerRecord == 0)
                    {
                        recordIndex++;
                    }
                }
            }

            OpenXMLFilePath = WorkbookGenerator.CreateWorkbookFile(true, FillAll, DiagonalFill, PartialFill, Table);
            BinaryFilePath = WorkbookGenerator.CreateWorkbookFile(false, FillAll, DiagonalFill, PartialFill, Table);
        }

        public void Dispose()
        {
            File.Delete(OpenXMLFilePath);
            File.Delete(BinaryFilePath);
        }

        public string OpenXMLFilePath { get; set; }
        public string BinaryFilePath { get; set; }
    }
}
