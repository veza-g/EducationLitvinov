using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace _ExcelRebuildWF
{
    internal class ExcelActions
    {
        public static void ReadExcel()
        {
            EXL EX_DATA = new EXL();
            object[,] dataArr = null;

            EXL EX_WRITE = new EXL();
            //object[,] dataArr1 = null;

            var listObjects = new List<ExcelObject>();

            try
            {
                OpenFileDialog inputFile = new OpenFileDialog();

                inputFile.Filter = "Файлы Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*";
                inputFile.FilterIndex = 1;
                inputFile.RestoreDirectory = true;

                if (inputFile.ShowDialog() != DialogResult.OK)
                    return;
                string xlFileName = inputFile.FileName;

                Excel.Range Rng;
                EX_DATA.App = new Excel.Application();
                EX_DATA.WBs = EX_DATA.App.Workbooks;
                EX_DATA.WB = EX_DATA.WBs.Open(xlFileName);
                EX_DATA.Shts = EX_DATA.WB.Worksheets;
                EX_DATA.Sht = EX_DATA.Shts.Item[1];
                EX_DATA.cell = EX_DATA.Sht.Cells[1, 1];

                int iLastRow = EX_DATA.cell[EX_DATA.Sht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
                int iLastCol = EX_DATA.cell[1, EX_DATA.Sht.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

                Rng = (Excel.Range)EX_DATA.Sht.Range["A1", EX_DATA.Sht.Cells[iLastRow, iLastCol]];

                dataArr = (object[,])Rng.Value;

                string[] arrCol = new string[iLastCol];

                List<string> materials = new List<string>();

                for (int i = 0; i < iLastRow; i++)
                {
                    listObjects.Add(new ExcelObject());

                    //listObjects[i].Id = i;
                    listObjects[i].Header = EX_DATA.cell[i + 1, "A"].Value;
                    listObjects[i].Наименование = EX_DATA.cell[i + 1, "B"].Value;
                    listObjects[i].Обозначение = EX_DATA.cell[i + 1, "C"].Value;
                    if (EX_DATA.cell[i + 1, "D"].Value.ToString() != null)
                    {
                        double number;
                        bool isNumber = double.TryParse(EX_DATA.cell[i + 1, "D"].Value.ToString().Trim(), out number);
                        if (isNumber)
                        {
                            EX_DATA.cell[i + 1, "D"].NumberFormat = "#,##0";
                            EX_DATA.cell[i + 1, "D"].Value = number;
                            listObjects[i].Количество = EX_DATA.cell[i + 1, "D"].Value;
                        }
                    }
                    listObjects[i].Материал = EX_DATA.cell[i + 1, "E"].Value;
                    if (listObjects[i].Материал != null && !materials.Contains(listObjects[i].Материал))
                    {
                        materials.Add(listObjects[i].Материал);
                    }
                    if (EX_DATA.cell[i + 1, "K"].Value != null)
                    {
                        double number;
                        bool isNumber = double.TryParse(EX_DATA.cell[i + 1, "K"].Value.ToString().Trim(), out number);
                        if (isNumber)
                        {
                            EX_DATA.cell[i + 1, "K"].NumberFormat = "#,##0";
                            EX_DATA.cell[i + 1, "K"].Value = number;
                            listObjects[i].Размер = EX_DATA.cell[i + 1, "K"].Value;
                        }
                    }
                    if (EX_DATA.cell[i + 1, "AA"].Value != null && i != 0)
                    {
                        double number;
                        bool isNumber = double.TryParse(EX_DATA.cell[i + 1, "AA"].Value.ToString().Trim(), out number);
                        if (isNumber)
                        {
                            EX_DATA.cell[i + 1, "AA"].NumberFormat = "#,##0";
                            EX_DATA.cell[i + 1, "AA"].Value = number;
                            listObjects[i].Вес = EX_DATA.cell[i + 1, "AA"].Value;
                        }
                    }
                    if (listObjects[i].Header == "Стандартные изделия" || listObjects[i].Header == "Прочие изделия")
                    {
                        materials.Add(listObjects[i].Наименование);
                    }

                }
                EX_DATA.App.Quit();
                Marshal.ReleaseComObject(EX_DATA.cell);
                Marshal.ReleaseComObject(EX_DATA.Sht);
                Marshal.ReleaseComObject(EX_DATA.Shts);
                Marshal.ReleaseComObject(EX_DATA.WB);
                Marshal.ReleaseComObject(EX_DATA.WBs);
                Marshal.ReleaseComObject(EX_DATA.App);

                EX_WRITE.App = new Excel.Application();
                EX_WRITE.App.SheetsInNewWorkbook = 1;
                EX_WRITE.WB = EX_WRITE.App.Workbooks.Add();
                EX_WRITE.Sht = EX_WRITE.WB.Worksheets[1];

                /*for (int i = 0; i < iLastRow; i++)
                {
                    //for (int j = 0; j < iLastCol; j++)
                    {
                        EX_WRITE.Sht.Range[$"{GetLetter(0)}{i + 1}"].Value = listObjects[i].Header;
                        EX_WRITE.Sht.Range[$"{GetLetter(1)}{i + 1}"].Value = listObjects[i].Наименование;
                        EX_WRITE.Sht.Range[$"{GetLetter(2)}{i + 1}"].Value = listObjects[i].Обозначение;
                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{i + 1}"].Value = listObjects[i].Количество;
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{i + 1}"].Value = listObjects[i].Материал;
                        EX_WRITE.Sht.Range[$"{GetLetter(5)}{i + 1}"].Value = listObjects[i].Размер;
                        EX_WRITE.Sht.Range[$"{GetLetter(6)}{i + 1}"].Value = listObjects[i].Длина;
                        EX_WRITE.Sht.Range[$"{GetLetter(7)}{i + 1}"].Value = listObjects[i].Вес;
                        EX_WRITE.Sht.Range[$"{GetLetter(8)}{i + 1}"].Value = listObjects[i].Id;

                    }
                }*/
                int excelIndex = 0;

                materials.RemoveAt(0);
                materials.RemoveAt(0);

                foreach (string material in materials)
                {
                    EX_WRITE.Sht.Range[$"{GetLetter(0)}{excelIndex + 1}"].Value = material;
                    double materialSumm = 0;
                    foreach (var listObject in listObjects)
                    {
                        if (listObject.Материал == material && (listObject.Материал.Contains("Прокат") || listObject.Материал.Contains("Лист")))
                        {
                            materialSumm += listObject.Вес;
                        }
                        if (listObject.Материал == material && (listObject.Header == "Материалы" || listObject.Наименование.Contains("Шина")))
                        {
                            materialSumm += listObject.Количество;
                        }
                        if (listObject.Материал == material && (listObject.Наименование.Contains("Ригель") || listObject.Наименование.Contains("Стойка")))
                        {
                            materialSumm += listObject.Размер / 1000;
                        }
                        if (listObject.Наименование == material && !listObject.Наименование.Contains("Уплотнитель") && !listObject.Наименование.Contains("Провод"))
                        {
                            materialSumm += listObject.Количество;
                        }
                    }
                    EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].NumberFormat = "0.0#";
                    EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].Value = materialSumm;
                    excelIndex++;
                }




                EX_WRITE.WB.SaveAs(@"C:\Users\litvinov.ls\Documents\Book1.xlsx");
                EX_WRITE.App.Quit();

                Marshal.ReleaseComObject(EX_WRITE.Sht);
                Marshal.ReleaseComObject(EX_WRITE.WB);
                Marshal.ReleaseComObject(EX_WRITE.App);
            }
            finally
            {
                Marshal.ReleaseComObject(EX_DATA.cell);
                Marshal.ReleaseComObject(EX_DATA.Sht);
                Marshal.ReleaseComObject(EX_DATA.Shts);
                Marshal.ReleaseComObject(EX_DATA.WB);
                Marshal.ReleaseComObject(EX_DATA.WBs);
                Marshal.ReleaseComObject(EX_DATA.App);
                //освобождаем память, занятую объектами
                Marshal.ReleaseComObject(EX_WRITE.Sht);
                Marshal.ReleaseComObject(EX_WRITE.WB);
                Marshal.ReleaseComObject(EX_WRITE.App);



            }
        }

        public struct EXL
        {
            public Excel.Application App;
            public Excel.Workbooks WBs;
            public Excel.Workbook WB;
            public Excel.Sheets Shts;
            public Excel.Worksheet Sht;
            public Excel.Range cell;
            public bool load;
        }

        static private string GetLetter(int nn)
        {
            string p1;

            int n2 = nn / 26;
            if (n2 > 0)
            {
                p1 = ((char)((int)('A') + n2 - 1)).ToString() + ((char)((int)('A') + nn - n2 * 26)).ToString();
            }
            else
            {
                p1 = ((char)((int)('A') + nn)).ToString();
            }

            return p1;
        }
    }
}
