// Decompiled with JetBrains decompiler
// Type: PriceConvert.ExcelReader
// Assembly: PriceConvert, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: DA1B6D49-AD90-4428-B938-61BD8BB8A206
// Assembly location: C:\MyProjects\HEOX.EBAY\book\PC\PriceConvert.exe

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace ConverToPrice
{
    public static class ExcelReader
    {
        public static List<InputItem> ReadXlsxOzon(string fn)
        {
            FileInfo newFile = new FileInfo(fn);
            List<InputItem> inputItemList = new List<InputItem>();
            using (ExcelPackage excelPackage = new ExcelPackage(newFile))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    if (worksheet.Name.ToLower() == "ozon" || worksheet.Name.ToLower() == "комплекты" || worksheet.Name.ToLower() == "wb")
                    {
                        for (int Row = 2; Row <= worksheet.Dimension.End.Row; ++Row)
                        {
                            try
                            {
                                InputItem inputItem = new InputItem();
                                if (worksheet.Name.ToLower() == "wb")
                                    inputItem.isWB = true;

                                if (worksheet.Cells[Row, 2].Value != null)
                                {
                                    inputItem._1_Id = worksheet.Cells[Row, 2].Value.ToString();
                                    inputItem._2_Name = worksheet.Cells[Row, 3].Value.ToString();
                                    inputItem._3_Year = worksheet.Cells[Row, 4].Value.ToString();
                                    double result;
                                    try
                                    {
                                        if (double.TryParse(worksheet.Cells[Row, 19].Value.ToString(), out result))
                                        {
                                            inputItem._4_Mrk = result;
                                            inputItemList.Add(inputItem);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        inputItem._4_Mrk = -1;
                                        inputItemList.Add(inputItem);
                                    }

                                }
                                else
                                    break;
                            }
                            catch (Exception ex)
                            {
                                int num = (int)MessageBox.Show(ex.Message);
                            }
                        }
                    }
                }
            }
            return inputItemList;
        }
    }
}
