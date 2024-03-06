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

        public static byte[] GenerateEmpty<T>(T t, string? filename)
        {
            filename = string.IsNullOrEmpty(filename) ? "default" : filename;
            filename += ".xlsx";

            var downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + $"\\Downloads\\{filename}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                using (var worksheet = package.Workbook.Worksheets.Add(filename))
                {
                    var properties = typeof(T).GetProperties().ToArray();

                    var columns = properties.Select(p => p.Name).ToList();

                    columns.ForEach(column =>
                    {
                        worksheet.Cells[1, columns.IndexOf(column) + 1].Value = columns[columns.IndexOf(column)];
                        worksheet.Column(columns.IndexOf(column) + 1).AutoFit();
                    });

                    using (var memoryStream = new MemoryStream())
                    {
                        package.SaveAs(memoryStream);
                        memoryStream.Position = 0;
                    }
                }

                return package.GetAsByteArray();
            }
        }

        public static void GenerateEmptyAndWrite<T>(T t, string? filename)
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
                    var properties = typeof(T).GetProperties().ToArray();

                    var columns = properties.Select(p => p.Name).ToList();

                    columns.ForEach(column =>
                    {
                        worksheet.Cells[firstRow, columns.IndexOf(column) + firstRow].Value = columns[columns.IndexOf(column)];
                        worksheet.Column(columns.IndexOf(column) + firstRow).AutoFit();
                    });

                    using (var memoryStream = new MemoryStream())
                    {
                        package.SaveAs(memoryStream);
                        File.WriteAllBytes(downloadsFolderPath, memoryStream.ToArray());
                        memoryStream.Position = 0;
                    }
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

        public static void GenerateFromObjectAndWrite<T>(List<T> t, string? filename)
        {
            filename = string.IsNullOrEmpty(filename) ? "default" : filename;
            filename += ".xlsx";

            if (t is null || t.Count == 0)
            {
                GenerateEmpty(filename);
            }

            var downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + $"\\Downloads\\{filename}";

            var properties = typeof(T).GetProperties().ToArray();

            var columns = properties.Select(p => p.Name).ToList();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                if (columns.Any())
                {
                    using (var worksheet = package.Workbook.Worksheets.Add(filename))
                    {
                        columns.ForEach(column =>
                        {
                            worksheet.Cells[1, columns.IndexOf(column) + 1].Value = columns[columns.IndexOf(column)];
                            worksheet.Column(columns.IndexOf(column) + 1).AutoFit();
                        });

                        t.AsParallel().ForAll(p =>
                        {
                            var i = 1;
                            var x = 0;
                            while (x < t.Count)
                            {
                                worksheet.Cells[x + 2, i].Value = properties[i - 1].GetValue(p);

                                if (i == columns.Count)
                                {
                                    x += 1;
                                    i = 0;
                                }
                                i++;
                            }
                        });

                        using (var memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);
                            File.WriteAllBytes(downloadsFolderPath, memoryStream.ToArray());
                            memoryStream.Position = 0;
                        }
                    }
                }
            }
        }

        public static List<T> ReadFile<T>(Stream stream)
        {
            if (!stream.CanRead)
            {
                throw new Exception("The file isn't readable");
            }

            using (var package = new ExcelPackage(stream))
            {
                var excel = package.Workbook.Worksheets.FirstOrDefault();
                if (excel == null)
                {
                    throw new Exception("No worksheet found in the Excel file.");
                }

                var properties = typeof(T).GetProperties();
                var result = new List<T>();

                for (int i = 2; i <= excel.Dimension.End.Row; i++)
                {
                    var instance = Activator.CreateInstance<T>();

                    for (int col = 1; col <= excel.Dimension.End.Column; col++)
                    {
                        var cellValue = excel.Cells[i, col].Text;
                        var property = properties[col - 1];

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            var convertedValue = Convert.ChangeType(cellValue, property.PropertyType);
                            property.SetValue(instance, convertedValue);
                        }
                    }
                    result.Add(instance);
                }

                return result;
            }
        }
    }
}