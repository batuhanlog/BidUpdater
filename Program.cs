using System;
using OfficeOpenXml;
using System.IO;

namespace BidUpdater.BidUpdater
{
    class Program
    {
        static void Main(string[] args)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            var filePath = @"C:\Users\batuh\OneDrive\Masaüstü\pratikler\Bidupdater.xlsx"; //Your file path here


            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sponsored Products Campaigns"]; //Your worksheet name here


                for (int row = 521; row <= worksheet.Dimension.End.Row; row++)
                {
                    var acosValue = worksheet.Cells[row, 46].Text;

                    if (decimal.TryParse(acosValue.TrimEnd('%'), out decimal acos) && acos > 0)
                    {
                        worksheet.Cells[row, 28].Value = 0.75m;
                        Console.WriteLine($"Row {row} - Bid updated to 0.75 because ACOS is {acosValue}");
                    }
                }


                package.Save();
            }

            Console.WriteLine("ACOS değeri %0'dan büyük olan satırlardaki Bid değerleri 0.75 olarak güncellendi.");
        }
    }
}
