using OfficeOpenXml;

namespace FilesHelper
{
    public class ExcelFactory
    {
        public static ExcelFactory Instance { get { return new ExcelFactory(); } }

        public ExcelFactory() { }

        public static byte[] GenerateEmpty(string filename)
        {
            filename = string.IsNullOrEmpty(filename) ? "default" : filename;
            filename += ".xlsx";

            var downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + $"\\Downloads\\{filename}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                using (var worksheet = package.Workbook.Worksheets.Add(filename))

                using (var memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream);
                    memoryStream.Position = 0;
                }

                return package.GetAsByteArray();
            }
        }

        public static byte[] GenerateEmpty(List<string> columns, string? filename)
        {
            var firstRow = 1;
            filename = string.IsNullOrEmpty(filename) ? "default" : filename;
            filename += ".xlsx";

            var downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + $"\\Downloads\\{filename}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                using (var worksheet = package.Workbook.Worksheets.Add(filename))
                {
                    columns.AsParallel().ForAll(column =>
                    {
                        worksheet.Cells[firstRow, columns.IndexOf(column) + firstRow].Value = columns[columns.IndexOf(column)];
                        worksheet.Column(columns.IndexOf(column) + firstRow).AutoFit();
                    });
                }

                using (var memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream);
                    memoryStream.Position = 0;
                }

                return package.GetAsByteArray();
            }
        }

        public static void GenerateEmptyAndWrite(List<string> columns, string? filename)
        {
            var firstRow = 1;
            filename = string.IsNullOrEmpty(filename) ? "default" : filename;
            filename += ".xlsx";

            var downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + $"\\Downloads\\{filename}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                using (var worksheet = package.Workbook.Worksheets.Add(filename))
                {
                    columns.AsParallel().ForAll(column =>
                    {
                        worksheet.Cells[firstRow, columns.IndexOf(column) + firstRow].Value = columns[columns.IndexOf(column)];
                        worksheet.Column(columns.IndexOf(column) + firstRow).AutoFit();
                    });
                }

                using (var memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream);
                    File.WriteAllBytes(downloadsFolderPath, memoryStream.ToArray());
                    memoryStream.Position = 0;
                }
            }
        }

        public static void GenerateEmptyAndWrite(string filename)
        {
            filename = string.IsNullOrEmpty(filename) ? "default" : filename;
            filename += ".xlsx";

            var downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + $"\\Downloads\\{filename}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                using (var worksheet = package.Workbook.Worksheets.Add(filename))

                using (var memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream);
                    File.WriteAllBytes($"{downloadsFolderPath}", memoryStream.ToArray());
                    memoryStream.Position = 0;
                }
            }
        }
    }
}