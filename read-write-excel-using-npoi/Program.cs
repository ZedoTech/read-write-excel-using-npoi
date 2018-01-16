using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace read_write_excel_using_npoi
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelRead();
            Console.WriteLine("Done.");
        }

        /// <summary>
        /// 用NPOI讀取EXCEL資料
        /// </summary>
        static void ExcelRead()
        {
            List<DataModel> models = new List<DataModel>();

            XSSFWorkbook hssfwb;
            //HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "input.xlsx", FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("sheet1");
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    DataModel model = new DataModel();
                    for (int i = 0; i < sheet.GetRow(row).Cells.Count; i++)
                    {
                        if (i == 0)
                        {
                            model.ActiveName = sheet.GetRow(row).GetCell(i).StringCellValue;
                        }
                        else if (i == 1)
                        {
                            model.Url = sheet.GetRow(row).GetCell(i).StringCellValue;
                        }
                        //Console.WriteLine(string.Format("Row {0}, Cell {1} = {2}", row, i, sheet.GetRow(row).GetCell(i).StringCellValue));
                    }
                    models.Add(model);
                }
            }

            WriteExcel(models);
        }

        /// <summary>
        /// 用NPOI輸出EXCEL
        /// </summary>
        /// <param name="models"></param>
        static void WriteExcel(List<DataModel> models)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet();
            IRow headerRow = sheet.CreateRow(0);

            int rowIndex = 0;
            foreach (DataModel item in models)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                int cellIndex = 0;
                dataRow.CreateCell(0).SetCellValue(item.ActiveName);
                dataRow.CreateCell(1).SetCellValue(item.Url);
                rowIndex++;
                //for (int i = 1; i <= item.GetType().GetProperties().Count() ; i++)
                //{
                //    dataRow.CreateCell(i).SetCellValue(item.GetType().);
                //}
            }

            var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "output.xlsx", FileMode.Create);
            workbook.Write(file);
            file.Close();

            //// handling header.
            //foreach (DataColumn column in SourceTable.Columns)
            //    headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

            // handling value.
            

            //foreach (DataRow row in SourceTable.Rows)
            //{
            //    HSSFRow dataRow = sheet.CreateRow(rowIndex);

            //    foreach (DataColumn column in SourceTable.Columns)
            //    {
            //        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
            //    }

            //    rowIndex++;
            //}

            //workbook.Write(ms);
            //ms.Flush();
            //ms.Position = 0;

            //sheet = null;
            //headerRow = null;
            //workbook = null;

            //return ms;
        }
    }
}
