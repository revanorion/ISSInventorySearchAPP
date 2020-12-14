using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Style;

namespace ISSIAS_Library.Excel
{
    public static class ExcelService
    {
        public static DataTable OpenBook(string book, string sheetName,
            string startLiteral,
            string endLiteral,
            ToDataTableOptions options)
        {
            using (var package = new ExcelPackage(new FileInfo(book)))
            {
                var workSheet =
                    package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToLower().Contains(sheetName.ToLower()));
                var maxRows = workSheet.Dimension.Rows;


                var start = workSheet.Cells
                    .Where(c => c.Value != null)
                    .First(c => c.Value.ToString().ToLower().Contains(startLiteral.ToLower())).Address;


                var endAddress = workSheet.Cells
                    .Where(c => c.Value != null)
                    .First(c => c.Value.ToString().ToLower().Contains(endLiteral.ToLower())).Address;

                var end = new string(endAddress.TakeWhile(x => !char.IsNumber(x)).ToArray());

                var dataTable = workSheet.Cells[$"{start}:{end}{maxRows}"]
                    .ToDataTable(options);

                return dataTable;
            }
        }


        public static void SaveBook(string path, string sheetName, IEnumerable<asset> assets, MemberInfo[] memberInfos)
        {
            //if (File.Exists(path)) File.Delete(path);

            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var workSheet = package.Workbook.Worksheets.Add(sheetName);
                workSheet.Cells["A1"]
                    .LoadFromCollection(assets, c =>
                    {
                        c.PrintHeaders = true;
                        if (memberInfos != null && memberInfos.Any())
                            c.Members = memberInfos;
                    });

                workSheet.Cells[$"C2:C{workSheet.Dimension.Rows}"].Style.Numberformat.Format = "mm-dd-yy";

                workSheet.Cells.AutoFitColumns();

                workSheet.Cells[$"B2:B{workSheet.Dimension.Rows}"].Value = "";


                ParallelWork(workSheet, assets);

                package.Save();
            }
        }

        private static void ParallelWork(ExcelWorksheet workSheet, IEnumerable<asset> assets)
        {
            Parallel.For(1, workSheet.Dimension.Rows + 1, index => { workSheet.Row(index).Height = 50; });

            var row = 0;
            foreach (var element in assets.AsParallel())
            {
                ++row;
                if (element.AssetBarcode == null) continue;

                var pic = workSheet.Drawings.AddPicture(Guid.NewGuid().ToString(), element.AssetBarcode);
                pic.SetPosition(row, 25, 1, 10);
                pic.LockAspectRatio = true;
                pic.ChangeCellAnchor(eEditAs.TwoCell);
            }
        }
    }
}