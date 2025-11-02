using BlazorIPaper.Models;
using ClosedXML.Excel;
using System.Reflection;

namespace BlazorIPaper.Services
{
    public class ExcelExportService
    {
        public byte[] ExportStudentsToExcel<T>(List<T> data)
        {
            if (data == null || data.Count == 0)
                throw new ArgumentException("Data cannot be null or empty.", nameof(data));

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(typeof(T).Name + "s");

            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance); 
            //These flags tell reflection which kinds of properties to include.

            // ✅ Header row
            for (int col = 0; col < properties.Length; col++)
            {
                worksheet.Cell(1, col + 1).Value = properties[col].Name;
                worksheet.Cell(1, col + 1).Style.Font.Bold = true;
                worksheet.Cell(1, col + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
            }

            // ✅ Data rows
            for (int row = 0; row < data.Count; row++)
            {
                var item = data[row];
                for (int col = 0; col < properties.Length; col++)
                {
                    var value = properties[col].GetValue(item);

                    // ✅ Safe assignment with type handling
                    if (value == null)
                    {
                        worksheet.Cell(row + 2, col + 1).SetValue(string.Empty);
                    }
                    else
                    {
                        worksheet.Cell(row + 2, col + 1).SetValue(value.ToString());
                    }
                }
            }

            worksheet.Columns().AdjustToContents(); 
            //Gets collection of all columns and Adjust all columns width to the content

            using var stream = new MemoryStream();  
            //Create a temporary in-memory storage area for the Excel file.
            workbook.SaveAs(stream); 
            //Write the workbook’s data into the in-memory stream instead of saving it to disk.
            return stream.ToArray(); 
            // Convert the Excel workbook (now saved in memory) into a byte array and return it.
        }
    }
}
