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
        public bool reg = false;
        DateTime t1, t2; //���� ������ �����
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

        public void OK(Document doc, ATTR_COM par)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            //System.Threading.Thread.Sleep(10000);

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
                        Variable voboz;
                        Variable vnaim;
                        if (doc.FindVariable("$�����������") != null)
                        {
                            voboz = doc.FindVariable("$�����������");
                        }
                        else voboz = null;
                        if (doc.FindVariable("$������������") != null)
                        {
                            vnaim = doc.FindVariable("$������������");
                        }
                        else vnaim = null;

                        if (voboz != null)
                        {
                            oboz = voboz.TextValue;
                        }
                        else oboz = "";
                        if (vnaim != null)
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
                        //rg.Full = true;
                        rg.UpdateAllLinks = true;
                        rg.UpdateProductStructures = true;
                        rg.UpdateBillOfMaterials = true;
                        //rg.Projections = true;
                        //rg.UpdateDrawingViews = true;
                        doc.Regenerate(rg);
                        reg = true;
                        GetFragmentData(doc, sw, file_name, doc.FilePath, oboz, naim, Level);

                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message, "������", System.Windows.Forms.MessageBoxButtons.OK);
                        TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";
                    }
                }
            }

            TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";

            doc.Selection.DeselectAll();
            /*switch (exportSet)
            {
                case 2:
                    System.Windows.Forms.MessageBox.Show("������� STEP � DXF ��������", "������ �� ����������(������)");
                    break;
                case 1:
                    System.Windows.Forms.MessageBox.Show("������� STEP ��������", "������ �� ����������(������)");
                    break;
                case 0:
                    System.Windows.Forms.MessageBox.Show("������� DXF ��������", "������ �� ����������(������)");
                    break;
                case -1:
                    System.Windows.Forms.MessageBox.Show("���������� ������� ���������(�� ������� ��������)", "������ �� ����������(������)");
                    break;
            }*/
            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("���������� ������� ���������\nRunTime: " + elapsedTime, "������ �� ����������(������)");
            else System.Windows.Forms.MessageBox.Show("�� ������� ��������� ��������", "������ �� ����������(������)");
            /*if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("���������� ������� ���������");
            else if (exportSet == -1)
                System.Windows.Forms.MessageBox.Show("���������� ������� ��������� (�� ������� ��������)");*/

            return;
        }

        #region OK2
        public void OK2(Document doc, ATTR_COM par)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            //System.Threading.Thread.Sleep(10000);
            if (par.pDXF == 1) exportDXF = true;
            if (par.pSTEP == 1) exportSTEP = true;
            if (par.pPDF == 1) exportPDF = true;
            if (!exportDXF && !exportSTEP && !exportPDF) exportSet = -1;

            if (exportSet != -1)
            {
                string logname = doc.FileName;
                logname = logname.Replace(".grb", ".log");
                using (StreamWriter sw = new StreamWriter(logname))
                {
                    if (doc == null) return;

                    try
                    {
                        string oboz = "-";
                        string naim = "-";
                        Variable voboz = doc.FindVariable("$�����������");
                        Variable vnaim = doc.FindVariable("$������������");

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
                        Set(doc, sw, file_name, doc.FilePath, oboz, naim, Level);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message, "������", System.Windows.Forms.MessageBoxButtons.OK);
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
                System.Windows.Forms.MessageBox.Show("���������� ������� ���������\nRunTime: " + elapsedTime, "������ �� ����������(������)");
            else System.Windows.Forms.MessageBox.Show("�� ������� ��������� ��������", "������ �� ����������(������)");

            return;
        }
        #endregion //OK2

        private void GetFragmentData(Document doc, StreamWriter sw, string name, string path, string oboz, string naim, int lev)
        {
            /*doc.BeginChanges("1");
            foreach (ProductStructure product in doc.GetProductStructures())
            {
                product.Regenerate(true);
                product.UpdateStructure();
                product.UpdateReports();
                //string abc = product.GetTextProperty("������");
                List<string> pp = new List<string>();
                foreach (var rows in product.GetAllRowElements())
                {
                    pp.Add(rows.Name);
                }
            }
            doc.EndChanges();*/

            RegenerateOptions rg = new RegenerateOptions();
            //rg.Full = true;
            rg.UpdateAllLinks = true;
            rg.UpdateProductStructures = true;
            rg.UpdateBillOfMaterials = true;
            //rg.Projections = true;
            //rg.UpdateDrawingViews = true;
            string offset = "";
            string confName;
            
            for (int nn = 0; nn < lev; nn++)
                offset += "   ";
            /*if (doc.ModelConfigurations.GetConfigurationWithCurrentValues() != null)
            {
                confName = doc.ModelConfigurations.GetConfigurationWithCurrentValues();
            }
            else*/
            confName = oboz;

            Variable voboz;
            Variable vnaim;
            if (doc.FindVariable("$�����������") != null)
            {
                voboz = doc.FindVariable("$�����������");
            }
            else voboz = null;
            if (doc.FindVariable("$������������") != null)
            {
                vnaim = doc.FindVariable("$������������");
            }
            else vnaim = null;

            if (voboz != null)
            {
                oboz = voboz.TextValue;
            }
            else oboz = "";
            if (vnaim != null)
            {
                naim = vnaim.TextValue;
            }
            else naim = "";

            if (((naim != "") || (oboz != "")) && (doc.ModelConfigurations.ConfigurationCount != 0))
            {
                if (exportSTEP)
                {
                    doc.Regenerate(rg);
                    reg = true;
                    ExportToStep exportSTPRU = new ExportToStep(doc);
                    string fileNameSTPRU = ($"{pathSTPRU}\\{confName}_{naim}.stp");
                    if (!File.Exists(fileNameSTPRU))
                    {
                        sw.WriteLine(offset + "���: " + name);
                        sw.WriteLine(offset + "����: " + path);
                        sw.WriteLine(offset + "������������: " + naim);
                        sw.WriteLine(offset + "�����������: " + oboz);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPRU.Export(fileNameSTPRU);
                    }
                }
                if (exportDXF)
                {
                    if (!reg) doc.Regenerate(rg);
                    ExportToDXF exportDXFRU = new ExportToDXF(doc);
                    Page pgRU = GetPageDXF(doc, "���������");
                    Page pgRU2 = GetPageDXF(doc, "Unfolding");
                    //Page pgRU = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgRU != null || pgRU2 != null)
                    {
                        List<Page> pgDXFRU = new List<Page>();
                        pgDXFRU.Add(pgRU);
                        pgDXFRU.Add(pgRU2);
                        exportDXFRU.ExportPages = pgDXFRU;
                        exportDXFRU.BiarcInterpolationForSplines = true;
                        string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{naim}.dxf");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            sw.WriteLine(offset + "DXF Export: OK");
                            exportDXFRU.Export(fileNameDXFRU);
                        }
                    }
                }
                if (exportPDF)
                {
                    if (!reg) doc.Regenerate(rg);
                    foreach (ProductStructure product in doc.GetProductStructures())
                    {

                        product.Regenerate(true);
                        product.UpdateStructure();
                        /*ProductStructureExcelExportOptions options = new ProductStructureExcelExportOptions();
                        options.FilePath = ($"{pathPDFRU}\\{doc.FindVariable("$�����������").TextValue}_{doc.FindVariable("$������������").TextValue}_{product.Name}.xlsx");
                        options.Silent = true;
                        TFlex.Model.Data.ProductStructure.GroupingRules item = new TFlex.Model.Data.ProductStructure.GroupingRules();
                        item.Name = "������������";
                        options.GroupingUID = item.ID;
                        product.ExportToExcel(options);*/
                    }
                    ExportToPDF exportPDFnormalconf = new ExportToPDF(doc);
                    List<Page> pgnormalconf = GetPagesPDF(doc, PageType.Normal);
                    if (pgnormalconf != null)
                    {
                        exportPDFnormalconf.ExportPages = pgnormalconf;
                        exportPDFnormalconf.OpenExportFile = false;
                        string fileNamePDFRU = ($"{pathPDFRU}\\{confName}_{naim}.pdf");
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
                        string fileNamePDFRU = ($"{pathPDFRU}\\{oboz}_{naim}_��.pdf");
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFBOM Export: OK");
                            exportPDFBOMconf.Export(fileNamePDFRU);
                        }
                    }
                    reg = false;
                }
            }
            else if (naim != "" || oboz != "")
            {
                if (exportSTEP)
                {
                    doc.Regenerate(rg);
                    reg = true;
                    ExportToStep exportSTPRU = new ExportToStep(doc);
                    string fileNameSTPRU = ($"{pathSTPRU}\\{oboz}_{naim}.stp");
                    if (!File.Exists(fileNameSTPRU))
                    {
                        sw.WriteLine(offset + "���: " + name);
                        sw.WriteLine(offset + "����: " + path);
                        sw.WriteLine(offset + "������������: " + naim);
                        sw.WriteLine(offset + "�����������: " + oboz);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPRU.Export(fileNameSTPRU);
                    }
                }
                if (exportDXF)
                {
                    if (!reg) doc.Regenerate(rg);
                    ExportToDXF exportDXFRU = new ExportToDXF(doc);
                    Page pgRUDXF = GetPageDXF(doc, "���������");
                    Page pgRUDXF2 = GetPageDXF(doc, "Unfolding");
                    //Page pgRUDXF = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgRUDXF != null || pgRUDXF2 != null)
                    {
                        List<Page> pgDXFRU = new List<Page>();
                        pgDXFRU.Add(pgRUDXF);
                        pgDXFRU.Add(pgRUDXF2);
                        exportDXFRU.ExportPages = pgDXFRU;
                        string fileNameDXFRU = ($"{pathDXFRU}\\{oboz}_{naim}.dxf");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            sw.WriteLine(offset + "DXF Export: OK");
                            exportDXFRU.Export(fileNameDXFRU);
                        }
                    }
                }
                if (exportPDF)
                {

                    if (!reg) doc.Regenerate(rg);

                    ExportToPDF exportPDFnormalRU = new ExportToPDF(doc);
                    List<Page> pgPDFnormalRU = GetPagesPDF(doc, PageType.Normal);
                    if (pgPDFnormalRU != null)
                    {
                        exportPDFnormalRU.ExportPages = pgPDFnormalRU;
                        exportPDFnormalRU.OpenExportFile = false;
                        string fileNamePDFRU = ($"{pathPDFRU}\\{oboz}_{naim}.pdf");
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
                        string fileNamePDFRU = ($"{pathPDFRU}\\{oboz}_{naim}_��.pdf");
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFBOM Export: OK");
                            exportPDFBOMRU.Export(fileNamePDFRU);
                        }
                    }
                    reg = false;
                }
            }
            int n_fr = doc.GetFragments3D().Count;
            foreach (Fragment3D frag in doc.GetFragments3D())
            {
                if (frag.Suppression.Suppress) continue;
                if (!frag.VisibleInScene) continue;
                if (frag.Layer.Hidden) continue;
                if (frag.FileName.Contains("����") || frag.FileName.Contains("����") || frag.FileName.Contains("��������") || frag.FileName.Contains("������") ||
                    frag.FileName.Contains("���") || frag.FileName.Contains("�����") || frag.FileName.Contains("�����") || frag.FileName.Contains("�����������") ||
                    frag.FileName.Contains("���������") || frag.FileName.Contains("�������") || frag.FileName.Contains("������") || frag.FileName.Contains("������") ||
                    frag.FileName.Contains("�����") || frag.FileName.Contains("�����") || frag.FileName.Contains("��������"))  continue;
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
                            //Variable voboz = docFR.FindVariable("$�����������");
                            //Variable vnaim = docFR.FindVariable("$������������");
                            if (vnaim != null)
                            {
                                naimF = vnaim.TextValue;
                            }
                            else
                            {
                                naimF = "���������� $������������ �� �������";
                            };
                            if (voboz != null)
                            {
                                obozF = voboz.TextValue;
                            }
                            else
                            {
                                obozF = "���������� $����������� �� �������";
                            };
                        }
                        else
                        {
                            err = true;
                            str_err = "������ ��������";
                        }
                    }
                    else
                    {
                        err = true;
                        str_err = "���� �� ������";
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
            return;
        }

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
                doc.BeginChanges("��������");
                string confName = doc.ModelConfigurations.GetConfigurationName(i);
                doc.ModelConfigurations.LoadConfigurationVariables(confName);
                doc.Regenerate(rg);

                /*if (exportSTEP)
                {
                    ExportToStep exportSTPconf = new ExportToStep(doc);
                    string fileNameSTPconf = ($"{pathSTPRU}\\{confName}_{doc.FindVariable("$������������").TextValue}.stp");
                    if (!File.Exists(fileNameSTPconf))
                    {
                        sw.WriteLine(offset + "���: " + name);
                        sw.WriteLine(offset + "����: " + path);
                        sw.WriteLine(offset + "������������: " + naim);
                        sw.WriteLine(offset + "�����������: " + confName);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPconf.Export(fileNameSTPconf);
                    }
                }
                if (exportDXF)
                {
                    ExportToDXF exportDXFconf = new ExportToDXF(doc);
                    Page pgconfDXF = GetPageDXF(doc, "���������");
                    Page pgconfDXF2 = GetPageDXF(doc, "Unfolding");
                    //Page pgconfDXF = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgconfDXF != null || pgconfDXF2 != null)
                    {
                        List<Page> pgDXFconf = new List<Page>();
                        pgDXFconf.Add(pgconfDXF);
                        pgDXFconf.Add(pgconfDXF2);
                        exportDXFconf.ExportPages = pgDXFconf;
                        string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{doc.FindVariable("$������������").TextValue}.dxf");
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
                        string fileNamePDFRU = ($"{pathPDFRU}\\{confName}_{doc.FindVariable("$������������").TextValue}.pdf");
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
                        string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$�����������").TextValue}_{doc.FindVariable("$������������").TextValue}_��.pdf");
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
                                Variable voboz = docFR.FindVariable("$�����������");
                                Variable vnaim = docFR.FindVariable("$������������");
                                if (vnaim != null)
                                {
                                    naimF = vnaim.TextValue;
                                }
                                else
                                {
                                    naimF = "���������� $������������ �� �������";
                                };
                                if (voboz != null)
                                {
                                    obozF = voboz.TextValue;
                                }
                                else
                                {
                                    obozF = "���������� $����������� �� �������";
                                };
                            }
                            else
                            {
                                err = true;
                                str_err = "������ ��������";
                            }
                        }
                        else
                        {
                            err = true;
                            str_err = "���� �� ������";
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

    public class ATTR_COM
    {
        public Int16 attr;
        public int pDXF // ������� � STP
        {
            get { return (attr & 0x0001); }
            set
            {
                attr = (Int16)(attr & 0xFFFE);
                attr = (Int16)(attr | (Int16)value);
            }
        }

        public int pSTEP // ������� � DXF
        {
            get { return ((attr & 0x0002) >> 1); }
            set
            {
                attr = (Int16)(attr & 0xFFFD);
                attr = (Int16)(attr | (Int16)(value << 1));
            }
        }
        public int pPDF // ������� � PDF
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