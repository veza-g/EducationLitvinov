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
            object[,] dataArr1 = null;

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
                EX_DATA.WB = EX_DATA.App.Workbooks.Open(xlFileName);
                EX_DATA.Sht = EX_DATA.WB.Worksheets[1];

                int iLastRow = EX_DATA.Sht.Cells[EX_DATA.Sht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
                int iLastCol = EX_DATA.Sht.Cells[1, EX_DATA.Sht.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

                Rng = (Excel.Range)EX_DATA.Sht.Range["A1", EX_DATA.Sht.Cells[iLastRow, iLastCol]];

                dataArr = (object[,])Rng.Value;

                string[] arrCol = new string[iLastCol];

                for (int i = 0; i < iLastRow; i++)
                {
                    listObjects.Add(new ExcelObject());

                    listObjects[i].Id = i;
                    //if (EX_DATA.Sht.Cells[i + 1, "A"].Value != null)
                    listObjects[i].Header = EX_DATA.Sht.Cells[i + 1, "A"].Value;
                    //if(EX_DATA.Sht.Cells[i + 1, "B"].Value!=null)
                    listObjects[i].Наименование = EX_DATA.Sht.Cells[i + 1, "B"].Value;
                    listObjects[i].Обозначение = EX_DATA.Sht.Cells[i + 1, "C"].Value;
                    if (EX_DATA.Sht.Cells[i + 1, "D"].Value.ToString() != null)
                        listObjects[i].Количество = EX_DATA.Sht.Cells[i + 1, "D"].Value.ToString();
                    listObjects[i].Материал = EX_DATA.Sht.Cells[i + 1, "E"].Value;
                    listObjects[i].Размер = EX_DATA.Sht.Cells[i + 1, "H"].Value;
                    if (EX_DATA.Sht.Cells[i + 1, "K"].Value != null)
                        listObjects[i].Длина = EX_DATA.Sht.Cells[i + 1, "K"].Value.ToString();
                    if (EX_DATA.Sht.Cells[i + 1, "AA"].Value != null)
                        listObjects[i].Вес = EX_DATA.Sht.Cells[i + 1, "AA"].Value.ToString();
                }
                EX_DATA.App.Quit();

                EX_WRITE.App = new Excel.Application();
                EX_WRITE.App.SheetsInNewWorkbook = 1;
                EX_WRITE.WB = EX_WRITE.App.Workbooks.Add();
                EX_WRITE.Sht = EX_WRITE.WB.Worksheets[1];

                for (int i = 0; i < iLastRow; i++)
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
                }
                EX_WRITE.WB.SaveAs(@"C:\Users\litvinov.ls\Documents\Book1.xlsx");


            }
            finally
            {
                //освобождаем память, занятую объектами

                Marshal.ReleaseComObject(EX_WRITE.Sht);
                Marshal.ReleaseComObject(EX_WRITE.WB);
                Marshal.ReleaseComObject(EX_WRITE.App);
                Marshal.ReleaseComObject(EX_DATA.Sht);
                Marshal.ReleaseComObject(EX_DATA.WB);
                Marshal.ReleaseComObject(EX_DATA.App);

            }
        }

        public struct EXL
        {
            public Excel.Application App;
            public Excel.Workbook WB;
            public Excel.Worksheet Sht;
            public Excel.Range RngAct;
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
