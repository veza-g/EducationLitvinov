using System;
//using Microsoft.Win32;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;
using TFlex.Drawing;
using TFlex.Command;

namespace FRAGMENTSTREE_PLG
{
    public class CommandManager
    {
        public int Level, exportSet;
        public string BasePath;
        public string pathSTPRU, pathDXFRU, pathPDFRU;
        public bool exportDXF, exportSTEP, exportPDF;
        DateTime t1, t2; //поля класса формы
        private readonly System.Diagnostics.Stopwatch uptime = new System.Diagnostics.Stopwatch();

        private Page GetPageDXF(Document doc, string namePage)
        {
            foreach (Page page in doc.GetPages())
            {
                if (page.Name == namePage)
                {
                    return page;
                }
            }
            return null;
        }

        private Page GetPageDXF2(Document doc, PageType pageType)
        {
            foreach (Page page in doc.GetPages())
            {
                if (page.PageType == pageType)
                {
                    return page;
                }
            }
            return null;
        }

        private Layer GetLayer(Document doc, string layerName)
        {
            foreach (Layer layer in doc.GetLayers())
            {
                if (layer.Name == layerName)
                {
                    return layer;
                }
            }
            return null;
        }

        private List<Page> GetPagesPDF(Document doc, PageType pageType)
        {
            List<Page> pagesList = new List<Page>();
            foreach (Page page in doc.GetPages())
            {
                if (page.PageType == pageType)
                {
                    pagesList.Add(page);
                }
            }
            if (pagesList.Count != 0) return pagesList;
            else return null;
        }

        private List<ProductStructure> GetProductStructures(Document doc)
        {
            List<ProductStructure> eleList = new List<ProductStructure>();
            foreach (ProductStructure ele in doc.GetProductStructures())
            {
                if (ele.DisplayName != null)
                {
                    eleList.Add(ele);
                }
            }
            if (eleList.Count != 0) return eleList;
            else return null;
        }

        public List<string> Profile;
        public List<ExcelObject> listObjects = new List<ExcelObject>();

        public void OK(Document doc, ATTR_COM par)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            if (par.pDXF == 1) exportDXF = true;
            if (par.pSTEP == 1) exportSTEP = true;
            if (par.pPDF == 1) exportPDF = true;
            if (!exportDXF && !exportSTEP && !exportPDF) exportSet = -1;
            else exportSet = 1;

            if (exportSet != -1)
            {
                string logname = doc.FileName;
                logname = logname.Replace(".grb", ".log");
                using (StreamWriter sw = new StreamWriter(logname))
                {
                    if (doc == null)
                        return;

                    try
                    {
                        string oboz = "-";
                        string naim = "-";
                        Variable voboz = doc.FindVariable("$Обозначение");
                        Variable vnaim = doc.FindVariable("$Наименование");

                        if (voboz != null)
                        {
                            oboz = voboz.TextValue;
                        }
                        else oboz = "";
                        if (naim != null)
                        {
                            naim = vnaim.TextValue;
                        }
                        else naim = "";



                        FileInfo parFile = new FileInfo(doc.FileName);
                        DirectoryInfo parDir = new DirectoryInfo(parFile.DirectoryName);
                        BasePath = parDir.FullName;
                        string subpathSTPRU = @"STP";
                        string subpathDXFRU = @"DXF";
                        string subpathPDFRU = @"PDF";
                        pathSTPRU = pathDXFRU = pathPDFRU = parFile.DirectoryName;
                        DirectoryInfo dirInfo = new DirectoryInfo(pathSTPRU);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        if (exportSTEP)
                            dirInfo.CreateSubdirectory(subpathSTPRU);
                        pathSTPRU = $"{pathSTPRU}\\{subpathSTPRU}";
                        if (exportDXF)
                            dirInfo.CreateSubdirectory(subpathDXFRU);
                        pathDXFRU = $"{pathDXFRU}\\{subpathDXFRU}";
                        if (exportPDF)
                            dirInfo.CreateSubdirectory(subpathPDFRU);
                        pathPDFRU = $"{pathPDFRU}\\{subpathPDFRU}";
                        string file_name = doc.FileName;
                        Level = 0;
                        TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.AutoRefresh;
                        Profile = new List<string>();
                        RegenerateOptions rg = new RegenerateOptions();
                        rg.Full = true;
                        rg.UpdateAllLinks = true;
                        rg.UpdateProductStructures = true;
                        rg.UpdateBillOfMaterials = true;
                        //rg.UpdateBillOfMaterials = true;
                        rg.Projections = true;

                        //doc.Regenerate(rg);
                        GetFragmentData(doc, sw, file_name, doc.FilePath, oboz, naim, Level);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message, "Ошибка", System.Windows.Forms.MessageBoxButtons.OK);
                        TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";
                    }
                }
            }


            TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";

            doc.Selection.DeselectAll();

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено\nRunTime: " + elapsedTime, "Команда 1");
            else System.Windows.Forms.MessageBox.Show("Не выбраны параметры экспорта", "Команда 1");

            return;
        }

        #region OK2
        public void OK2(Document doc, ATTR_COM par)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();



            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено\nRunTime: " + elapsedTime, "Команда 2");
            else System.Windows.Forms.MessageBox.Show("Не выбраны параметры экспорта", "Команда 2");

            return;
        }
        #endregion OK2

        private void GetFragmentData(Document doc, StreamWriter sw, string name, string path, string oboz, string naim, int lev)
        {
            if (naim.Contains("Моноблок") || naim.Contains("Корпус") || naim.Contains("Панель") || naim.Contains("Обшивка") || naim.Contains("Усилитель"))
            {
                RegenerateOptions rg = new RegenerateOptions();
                rg.Full = true;
                rg.UpdateAllLinks = true;
                rg.UpdateProductStructures = true;
                rg.UpdateBillOfMaterials = true;
                rg.Projections = true;
                rg.UpdateDrawingViews = true;

                string offset = "";
                string confName;
                bool reg = false;
                for (int nn = 0; nn < lev; nn++)
                    offset += "   ";

                confName = oboz;

                #region Conf
                /*if (((doc.FindVariable("$Наименование").TextValue != "") || (doc.FindVariable("$Обозначение").TextValue != ""))
                    && (doc.ModelConfigurations.ConfigurationCount != 0))
                {
                    if (exportSTEP)
                    {
                        //doc.Regenerate(rg);
                        reg = true;
                        ExportToStep exportSTPRU = new ExportToStep(doc);
                        string fileNameSTPRU = ($"{pathSTPRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.stp");
                        if (!File.Exists(fileNameSTPRU))
                        {
                            sw.WriteLine(offset + "Имя: " + name);
                            sw.WriteLine(offset + "Путь: " + path);
                            sw.WriteLine(offset + "Наименование: " + naim);
                            sw.WriteLine(offset + "Обозначение: " + oboz);
                            sw.WriteLine(offset + "STEP Export: OK");
                            exportSTPRU.Export(fileNameSTPRU);
                        }
                    }
                    if (exportDXF)
                    {
                        //if (!reg) doc.Regenerate(rg);
                        ExportToDXF exportDXFRU = new ExportToDXF(doc);

                        Page pgRU = GetPageDXF(doc, "Развертка");
                        Page pgRU2 = GetPageDXF(doc, "Развертка_1");
                        //Page pgRU = GetPageDXF2(doc, PageType.Auxiliary);
                        if (pgRU != null)
                        {
                            List<Page> pgDXFRU = new List<Page>();
                            pgDXFRU.Add(pgRU);
                            exportDXFRU.ExportPages = pgDXFRU;
                            exportDXFRU.BiarcInterpolationForSplines = true;
                            string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{doc.FindVariable("$Наименование_1").TextValue}.dxf");
                            if (!File.Exists(fileNameDXFRU))
                            {
                                sw.WriteLine(offset + "DXF Export: OK");
                                exportDXFRU.Export(fileNameDXFRU);
                            }
                        }
                        if (pgRU2 != null)
                        {
                            List<Page> pgDXFRU = new List<Page>();
                            pgDXFRU.Add(pgRU2);
                            exportDXFRU.ExportPages = pgDXFRU;
                            exportDXFRU.BiarcInterpolationForSplines = true;
                            string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{doc.FindVariable("$Наименование_2").TextValue}.dxf");
                            if (!File.Exists(fileNameDXFRU))
                            {
                                sw.WriteLine(offset + "DXF Export: OK");
                                exportDXFRU.Export(fileNameDXFRU);
                            }
                        }
                    }
                    if (exportPDF)
                    {
                        //if (!reg) doc.Regenerate(rg);
                        /*foreach (ProductStructure product in doc.GetProductStructures())
                        {

                            product.Regenerate(true);
                            product.UpdateStructure();
                            /*ProductStructureExcelExportOptions options = new ProductStructureExcelExportOptions();
                            options.FilePath = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_{product.Name}.xlsx");
                            options.Silent = true;
                            TFlex.Model.Data.ProductStructure.GroupingRules item = new TFlex.Model.Data.ProductStructure.GroupingRules();
                            item.Name = "Спецификация";
                            options.GroupingUID = item.ID;
                            product.ExportToExcel(options);*/
                //}
                /*ExportToPDF exportPDFnormalconf = new ExportToPDF(doc);
                List<Page> pgnormalconf = GetPagesPDF(doc, PageType.Normal);
                if (pgnormalconf != null)
                {
                    exportPDFnormalconf.ExportPages = pgnormalconf;
                    exportPDFnormalconf.OpenExportFile = false;
                    string fileNamePDFRU = ($"{pathPDFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.pdf");
                    if (!File.Exists(fileNamePDFRU))
                    {
                        sw.WriteLine(offset + "PDF Export: OK");
                        exportPDFnormalconf.Export(fileNamePDFRU);
                    }
                }
                ExportToPDF exportPDFBOMconf = new ExportToPDF(doc);
                List<Page> pgPDFBOMconf = GetPagesPDF(doc, PageType.BillOfMaterials);
                if (pgPDFBOMconf != null)
                {
                    exportPDFBOMconf.ExportPages = pgPDFBOMconf;
                    exportPDFBOMconf.OpenExportFile = false;
                    string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_СП.pdf");
                    if (!File.Exists(fileNamePDFRU))
                    {
                        sw.WriteLine(offset + "PDFBOM Export: OK");
                        exportPDFBOMconf.Export(fileNamePDFRU);
                    }
                }
                reg = false;
            }
        }*/
                #endregion

                if (((naim.Contains("Панель") || naim.Contains("Обшивка") || naim.Contains("Усилитель")) &&
                    ((doc.FindVariable("$Наименование").TextValue != "") || (doc.FindVariable("$Обозначение").TextValue != ""))))
                {
                    doc.Regenerate(rg);

                    EXL EX_WRITE = new EXL();

                    string xlFileName = BasePath;

                    EX_WRITE.App = new Excel.Application();
                    EX_WRITE.App.SheetsInNewWorkbook = 1;
                    EX_WRITE.WB = EX_WRITE.App.Workbooks.Add();
                    EX_WRITE.Shts = EX_WRITE.WB.Worksheets;
                    EX_WRITE.Sht = EX_WRITE.Shts.Item[1];

                    if (exportSTEP)
                    {
                        //doc.Regenerate(rg);

                        reg = true;
                        ExportToStep exportSTPRU = new ExportToStep(doc);
                        string fileNameSTPRU = ($"{pathSTPRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}.stp");
                        if (!File.Exists(fileNameSTPRU))
                        {
                            sw.WriteLine(offset + "Имя: " + name);
                            sw.WriteLine(offset + "Путь: " + path);
                            sw.WriteLine(offset + "Наименование: " + naim);
                            sw.WriteLine(offset + "Обозначение: " + oboz);
                            sw.WriteLine(offset + "STEP Export: OK");
                            exportSTPRU.Export(fileNameSTPRU);
                            string flagStandart = "";
                            if (doc.FindVariable("Нестандартная_панель") != null)
                            {
                                var flagStandVariable = doc.FindVariable("Нестандартная_панель");
                                if (doc.FindVariable("Нестандартная_панель").RealValue == 1)
                                    flagStandart = "true";
                                else if (doc.FindVariable("Нестандартная_панель").RealValue == 0)
                                    flagStandart = "false";
                            }

                            listObjects.Add(new ExcelObject()
                            {
                                Наименование = doc.FindVariable("$Наименование").TextValue,
                                Обозначение = doc.FindVariable("$Обозначение").TextValue,
                                Нестандартные = flagStandart,
                                Размер = doc.FindVariable("$Размер").TextValue,
                            });
                        }
                        if (!naim.Contains("Усилитель"))
                        {
                            doc.BeginChanges("11");
                            foreach (Layer layer in doc.GetLayers())
                            {
                                if (layer.Name != "Наружная обшивка 3D")
                                {
                                    layer.Hidden = true;
                                }
                            }
                            doc.ApplyChanges();



                            ExportToStep exportOuter = new ExportToStep(doc);
                            exportOuter.ExportSheetBodies = true;
                            exportOuter.ExportWireBodies = true;
                            exportOuter.ExportSolidBodies = true;
                            string fileNameOuter = ($"{pathSTPRU}\\{doc.FindVariable("$Обозначение_1").TextValue}_{doc.FindVariable("$Наименование_1").TextValue}.stp");
                            if (!File.Exists(fileNameOuter))
                            {
                                exportOuter.Export(fileNameOuter);
                                string flagStandart = "";
                                if (doc.FindVariable("Нестандартная_панель") != null)
                                {
                                    var flagStandVariable = doc.FindVariable("Нестандартная_панель");
                                    if (doc.FindVariable("Нестандартная_панель").RealValue == 1)
                                        flagStandart = "true";
                                    else if (doc.FindVariable("Нестандартная_панель").RealValue == 0)
                                        flagStandart = "false";
                                }

                                listObjects.Add(new ExcelObject()
                                {
                                    Наименование = doc.FindVariable("$Наименование_1").TextValue,
                                    Обозначение = doc.FindVariable("$Обозначение_1").TextValue,
                                    Нестандартные = flagStandart,
                                    Размер = doc.FindVariable("$Размер").TextValue,
                                });
                            }
                            doc.CancelChanges();

                            doc.BeginChanges("12");
                            foreach (Layer layer in doc.GetLayers())
                            {
                                if (layer.Name != "Внутренняя обшивка 3D")
                                {
                                    layer.Hidden = true;
                                }
                            }
                            doc.ApplyChanges();

                            ExportToStep exportInner = new ExportToStep(doc);
                            exportInner.ExportSheetBodies = true;
                            exportInner.ExportWireBodies = true;
                            exportInner.ExportSolidBodies = true;
                            string fileNameInner = ($"{pathSTPRU}\\{doc.FindVariable("$Обозначение_2").TextValue}_{doc.FindVariable("$Наименование_2").TextValue}.stp");
                            if (!File.Exists(fileNameInner))
                            {
                                exportInner.Export(fileNameInner);
                                string flagStandart = "";
                                if (doc.FindVariable("Нестандартная_панель") != null)
                                {
                                    var flagStandVariable = doc.FindVariable("Нестандартная_панель");
                                    if (doc.FindVariable("Нестандартная_панель").RealValue == 1)
                                        flagStandart = "true";
                                    else if (doc.FindVariable("Нестандартная_панель").RealValue == 0)
                                        flagStandart = "false";
                                }

                                listObjects.Add(new ExcelObject()
                                {
                                    Наименование = doc.FindVariable("$Наименование_2").TextValue,
                                    Обозначение = doc.FindVariable("$Обозначение_2").TextValue,
                                    Нестандартные = flagStandart,
                                    Размер = doc.FindVariable("$Размер").TextValue,
                                });
                            }
                            doc.CancelChanges();
                        }
                    }
                    if (exportDXF)
                    {
                        if (!naim.Contains("Усилитель"))
                        {
                            //if (!reg) doc.Regenerate(rg);
                            ExportToDXF exportDXFRU = new ExportToDXF(doc);
                            Page pgRUDXF = GetPageDXF(doc, "Развертка");
                            Page pgRUDXF2 = GetPageDXF(doc, "Развертка_1");
                            //Page pgRUDXF = GetPageDXF2(doc, PageType.Auxiliary);
                            if (pgRUDXF != null)
                            {
                                List<Page> pgDXFRU = new List<Page>();
                                pgDXFRU.Add(pgRUDXF);
                                exportDXFRU.ExportPages = pgDXFRU;
                                string fileNameDXFRU = ($"{pathDXFRU}\\{doc.FindVariable("$Обозначение_1").TextValue}_{doc.FindVariable("$Наименование_1").TextValue}.dxf");
                                if (!File.Exists(fileNameDXFRU))
                                {
                                    sw.WriteLine($"{pathDXFRU}\\{doc.FindVariable("$Обозначение_1").TextValue}_{doc.FindVariable("$Наименование_1").TextValue}");
                                    sw.WriteLine(offset + "DXF Export: OK");
                                    exportDXFRU.Export(fileNameDXFRU);
                                }
                            }
                            if (pgRUDXF2 != null)
                            {
                                List<Page> pgDXFRU = new List<Page>();
                                pgDXFRU.Add(pgRUDXF2);
                                exportDXFRU.ExportPages = pgDXFRU;
                                string fileNameDXFRU = ($"{pathDXFRU}\\{doc.FindVariable("$Обозначение_2").TextValue}_{doc.FindVariable("$Наименование_2").TextValue}.dxf");
                                if (!File.Exists(fileNameDXFRU))
                                {
                                    sw.WriteLine($"{pathDXFRU}\\{doc.FindVariable("$Обозначение_2").TextValue}_{doc.FindVariable("$Наименование_2").TextValue}");
                                    sw.WriteLine(offset + "DXF Export: OK");
                                    exportDXFRU.Export(fileNameDXFRU);
                                }
                            }
                        }
                        else if (naim.Contains("Усилитель"))
                        {
                            ExportToDXF exportDXFRU = new ExportToDXF(doc);
                            Page pgRUDXF = GetPageDXF(doc, "Развертка");
                            if (pgRUDXF != null)
                            {
                                List<Page> pgDXFRU = new List<Page>();
                                pgDXFRU.Add(pgRUDXF);
                                exportDXFRU.ExportPages = pgDXFRU;
                                string fileNameDXFRU = ($"{pathDXFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}.dxf");
                                if (!File.Exists(fileNameDXFRU))
                                {
                                    sw.WriteLine($"{pathDXFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}");
                                    sw.WriteLine(offset + "DXF Export: OK");
                                    exportDXFRU.Export(fileNameDXFRU);
                                }
                            }
                        }
                    }
                    if (exportPDF)
                    {
                        /*doc.BeginChanges("1");
                        if (!reg) doc.Regenerate(rg);
                        foreach (ProductStructure product in doc.GetProductStructures())
                        {
                            product.Regenerate(true);
                            product.UpdateStructure();
                            product.UpdateReports();
                            ProductStructureExcelExportOptions options = new ProductStructureExcelExportOptions();
                            options.FilePath = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_{product.Name}.xlsx");
                            options.Silent = true;
                            TFlex.Model.Data.ProductStructure.GroupingRules item = new TFlex.Model.Data.ProductStructure.GroupingRules();
                            item.Name = "Спецификация";
                            options.GroupingUID = item.ID;
                            product.ExportToExcel(options);
                        }
                        doc.EndChanges();*/

                        ExportToPDF exportPDFnormalRU = new ExportToPDF(doc);
                        List<Page> pgPDFnormalRU = GetPagesPDF(doc, PageType.Normal);
                        if (pgPDFnormalRU != null)
                        {
                            exportPDFnormalRU.ExportPages = pgPDFnormalRU;
                            exportPDFnormalRU.OpenExportFile = false;
                            string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}.pdf");
                            if (!File.Exists(fileNamePDFRU))
                            {
                                sw.WriteLine(offset + "PDFnormal Export: OK");
                                exportPDFnormalRU.Export(fileNamePDFRU);
                            }
                        }

                        ExportToPDF exportPDFBOMRU = new ExportToPDF(doc);
                        List<Page> pgPDFBOMRU = GetPagesPDF(doc, PageType.BillOfMaterials);
                        if (pgPDFBOMRU != null)
                        {
                            exportPDFBOMRU.ExportPages = pgPDFBOMRU;
                            exportPDFBOMRU.OpenExportFile = false;
                            string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_СП.pdf");
                            if (!File.Exists(fileNamePDFRU))
                            {
                                sw.WriteLine(offset + "PDFBOM Export: OK");
                                exportPDFBOMRU.Export(fileNamePDFRU);
                            }
                        }
                        reg = false;
                    }

                    int iter = 0;
                    EX_WRITE.Sht.Range[$"{GetLetter(0)}{1}"].Value2 = "Наименование";
                    EX_WRITE.Sht.Range[$"{GetLetter(1)}{1}"].Value2 = "Обозначение";
                    EX_WRITE.Sht.Range[$"{GetLetter(2)}{1}"].Value2 = "Нестандартные";
                    EX_WRITE.Sht.Range[$"{GetLetter(3)}{1}"].Value2 = "Размер";

                    foreach (var excelObject in listObjects)
                    {
                        iter++;
                        EX_WRITE.Sht.Range[$"{GetLetter(0)}{iter + 1}"].Value2 = excelObject.Наименование;
                        EX_WRITE.Sht.Range[$"{GetLetter(1)}{iter + 1}"].Value2 = excelObject.Обозначение;
                        EX_WRITE.Sht.Range[$"{GetLetter(2)}{iter + 1}"].Value2 = excelObject.Нестандартные;
                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{iter + 1}"].Value2 = excelObject.Размер;
                    }

                    EX_WRITE.App.DisplayAlerts = false;
                    EX_WRITE.WB.SaveAs(xlFileName.Substring(0, xlFileName.Length) + "\\" + "Панели.xlsx");
                    EX_WRITE.App.Quit();
                    EX_WRITE.App.DisplayAlerts = true;

                    Marshal.ReleaseComObject(EX_WRITE.Sht);
                    Marshal.ReleaseComObject(EX_WRITE.WB);
                    Marshal.ReleaseComObject(EX_WRITE.App);
                }

                int n_fr = doc.GetFragments3D().Count;
                foreach (Fragment3D frag in doc.GetFragments3D())
                {
                    if (frag.Suppression.Suppress) continue;
                    else 
                    {
                        Document docFR = null;
                        string obozF = "-";
                        string naimF = "-";
                        bool err = false;
                        string str_err = "";
                        string FRname = frag.FullFilePath;
                        if (lev > 0)
                        {
                            FRname = frag.FilePath;
                            FRname = TFlex.Application.FindPathName(FRname);
                        }

                        if (File.Exists(FRname))
                        {
                            Fragment.OpenPartOptions options = new Fragment.OpenPartOptions();
                            options.DontShowDocument = true;
                            options.QuietMode = true;
                            options.SubstituteGeometry = true;
                            options.SubstituteVariables = true;
                            options.SubstituteStatus = true;
                            //options.SaveDocument = true;
                            /*bool exist = false;
                            foreach (string nameP in Profile)
                            {
                                if (nameP == FRname)
                                {
                                    exist = true;
                                    break;
                                }
                            }
                            if (exist) continue;*/
                            docFR = frag.OpenPart(options);

                            if (docFR != null)
                            {
                                Variable voboz = docFR.FindVariable("$Обозначение");
                                Variable vnaim = docFR.FindVariable("$Наименование");
                                if (vnaim != null)
                                {
                                    naimF = vnaim.TextValue;
                                }
                                else
                                {
                                    naimF = "Переменная $Наименование не найдена";
                                };
                                if (voboz != null)
                                {
                                    obozF = voboz.TextValue;
                                }
                                else
                                {
                                    obozF = "Переменная $Обозначение не найдена";
                                };
                            }
                            else
                            {
                                err = true;
                                str_err = "Ошибка открытия";
                            }
                        }
                        else
                        {
                            err = true;
                            str_err = "Файл не найден";
                        }

                        if (err == false)
                        {
                            //docFR.Regenerate(rg);
                            GetFragmentData(docFR, sw, FRname, frag.FullFilePath, obozF, naimF, lev + 1);
                            //docFR.Save();
                            //Profile.Add(FRname);
                            docFR.Close();
                        }
                    }
                }
            }
            return;
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

        #region Set
        private void Set(Document doc, StreamWriter sw, string name, string path, string oboz, string naim, int lev)
        {
            string offset = "";
            for (int nn = 0; nn < lev; nn++)
                offset += "   ";

            FileInfo file = new FileInfo(name);

            int ConfCount = doc.ModelConfigurations.ConfigurationCount;
            ModelConfiguration[] ConfArray = new ModelConfiguration[ConfCount];
            RegenerateOptions rg = new RegenerateOptions();
            rg.Full = true;
            rg.UpdateAllLinks = true;
            rg.UpdateProductStructures = true;
            rg.UpdateBillOfMaterials = true;
            rg.Projections = true;
            rg.UpdateDrawingViews = true;
            for (int i = 0; i < ConfArray.Length; i++)
            {
                doc.BeginChanges("Пересчёт");
                string confName = doc.ModelConfigurations.GetConfigurationName(i);
                doc.ModelConfigurations.LoadConfigurationVariables(confName);
                //doc.Regenerate(rg);

                /*if (exportSTEP)
                {
                    ExportToStep exportSTPconf = new ExportToStep(doc);
                    string fileNameSTPconf = ($"{pathSTPRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.stp");
                    if (!File.Exists(fileNameSTPconf))
                    {
                        sw.WriteLine(offset + "Имя: " + name);
                        sw.WriteLine(offset + "Путь: " + path);
                        sw.WriteLine(offset + "Наименование: " + naim);
                        sw.WriteLine(offset + "Обозначение: " + confName);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPconf.Export(fileNameSTPconf);
                    }
                }
                if (exportDXF)
                {
                    ExportToDXF exportDXFconf = new ExportToDXF(doc);
                    Page pgconfDXF = GetPageDXF(doc, "Развертка");
                    Page pgconfDXF2 = GetPageDXF(doc, "Unfolding");
                    //Page pgconfDXF = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgconfDXF != null || pgconfDXF2 != null)
                    {
                        List<Page> pgDXFconf = new List<Page>();
                        pgDXFconf.Add(pgconfDXF);
                        pgDXFconf.Add(pgconfDXF2);
                        exportDXFconf.ExportPages = pgDXFconf;
                        string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.dxf");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            sw.WriteLine(offset + "DXF Export: OK");
                            exportDXFconf.Export(fileNameDXFRU);
                        }
                    }
                }*/
                if (exportPDF)
                {
                    ExportToPDF exportPDFnormalconf = new ExportToPDF(doc);
                    List<Page> pgnormalconf = GetPagesPDF(doc, PageType.Normal);
                    if (pgnormalconf != null)
                    {
                        exportPDFnormalconf.ExportPages = pgnormalconf;
                        exportPDFnormalconf.OpenExportFile = false;
                        string fileNamePDFRU = ($"{pathPDFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.pdf");
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFnormal Export: OK");
                            exportPDFnormalconf.Export(fileNamePDFRU);
                        }
                    }
                    ExportToPDF exportPDFBOMconf = new ExportToPDF(doc);
                    List<Page> pgPDFBOMconf = GetPagesPDF(doc, PageType.BillOfMaterials);
                    if (pgPDFBOMconf != null)
                    {
                        exportPDFBOMconf.ExportPages = pgPDFBOMconf;
                        exportPDFBOMconf.OpenExportFile = false;
                        string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_СП.pdf");
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFBOM Export: OK");
                            exportPDFBOMconf.Export(fileNamePDFRU);
                        }
                    }
                }
                doc.EndChanges();
            }

            int n_fr = doc.GetFragments3D().Count;

            foreach (Fragment3D frag in doc.GetFragments3D())
            {
                if (frag.Suppression.Suppress) continue;
                else
                {
                    {
                        Document docFR = null;
                        string obozF = "-";
                        string naimF = "-";
                        bool err = false;
                        string str_err = "";
                        string FRname = frag.FullFilePath;
                        if (lev > 0)
                        {
                            FRname = frag.FilePath;
                            FRname = TFlex.Application.FindPathName(FRname);
                        }

                        if (File.Exists(FRname))
                        {
                            Fragment.OpenPartOptions options = new Fragment.OpenPartOptions();
                            options.DontShowDocument = true;
                            options.QuietMode = true;
                            options.SubstituteGeometry = true;
                            options.SubstituteVariables = true;
                            docFR = frag.OpenPart(options);

                            if (docFR != null)
                            {
                                Variable voboz = docFR.FindVariable("$Обозначение");
                                Variable vnaim = docFR.FindVariable("$Наименование");
                                if (vnaim != null)
                                {
                                    naimF = vnaim.TextValue;
                                }
                                else
                                {
                                    naimF = "Переменная $Наименование не найдена";
                                };
                                if (voboz != null)
                                {
                                    obozF = voboz.TextValue;
                                }
                                else
                                {
                                    obozF = "Переменная $Обозначение не найдена";
                                };
                            }
                            else
                            {
                                err = true;
                                str_err = "Ошибка открытия";
                            }
                        }
                        else
                        {
                            err = true;
                            str_err = "Файл не найден";
                        }

                        if (err == false)
                        {
                            Set(docFR, sw, FRname, frag.FullFilePath, obozF, naimF, lev + 1);
                            docFR.Close();
                        }
                    }
                }
            }
            return;
        }
    }
    #endregion
    public class ATTR_COM
    {
        public Int16 attr;
        public int pDXF // Экспорт в STP
        {
            get { return (attr & 0x0001); }
            set
            {
                attr = (Int16)(attr & 0xFFFE);
                attr = (Int16)(attr | (Int16)value);
            }
        }

        public int pSTEP // Экспорт в DXF
        {
            get { return ((attr & 0x0002) >> 1); }
            set
            {
                attr = (Int16)(attr & 0xFFFD);
                attr = (Int16)(attr | (Int16)(value << 1));
            }
        }
        public int pPDF // Экспорт в PDF
        {
            get { return ((attr & 0x0004) >> 2); }
            set
            {
                attr = (Int16)(attr & 0xFFFB);
                attr = (Int16)(attr | (Int16)(value << 2));
            }
        }

        public void Set(int dxfs, int steps, int pdfs)
        {
            pDXF = dxfs;
            pSTEP = steps;
            pPDF = pdfs;
        }


    }
}