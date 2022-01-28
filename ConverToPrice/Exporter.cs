
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ConverToPrice
{
    public static class Exporter
    {
        public static void WriteFile(string fn, List<InputItem> items_list)
        {

            if (File.Exists(fn))
                File.Delete(fn);
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fn)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("ИП Вельмога");
                excelWorksheet.Cells[1, 1].Value = (object)"ID";
                excelWorksheet.Cells[1, 1].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 2].Value = (object)"Название";
                excelWorksheet.Cells[1, 2].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 3].Value = (object)"Год";
                excelWorksheet.Cells[1, 3].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 4].Value = (object)"Цена Маркета";
                excelWorksheet.Cells[1, 4].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 5].Value = (object)"Цена Розница";
                excelWorksheet.Cells[1, 5].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 6].Value = (object)"Кол-во OZON";
                excelWorksheet.Cells[1, 6].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 7].Value = (object)"Кол-во WB";
                excelWorksheet.Cells[1, 7].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 8].Value = (object)"Кол-во";
                excelWorksheet.Cells[1, 8].Style.Font.Bold = true;
                excelWorksheet.Cells[1, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                int Row = 2;
                foreach (InputItem items in items_list)
                {
                    excelWorksheet.Cells[Row, 1].Value = (object)items._1_Id;
                    excelWorksheet.Cells[Row, 2].Value = (object)items._2_Name;
                    excelWorksheet.Cells[Row, 3].Value = (object)items._3_Year;

                    if (!(items._4_Mrk == -1))
                        excelWorksheet.Cells[Row, 4].Value = (object)items._4_Mrk;

                    excelWorksheet.Cells[Row, 6].Value = (object)items._6_CountOZ;
                    excelWorksheet.Cells[Row, 7].Value = (object)items._7_CountWB;
                    excelWorksheet.Cells[Row, 8].Value = (object)items._8_Count;
                    ++Row;
                }
                excelWorksheet.Cells[excelWorksheet.Dimension.Address].AutoFitColumns();
                excelWorksheet.Column(2).Width = 100.0;
                excelPackage.Save();
            }
        }
    }
}
