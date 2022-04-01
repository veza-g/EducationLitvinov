using System;
using System.Globalization;
using Microsoft.Win32;
using System.IO;
using System.Windows.Forms;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
//using TFlex.Model.Model3D;
using TFlex.Command;
using System.Drawing;

namespace FRAGMENTSTREE_PLG
{
    public class Factory : PluginFactory
    {
        public override Plugin CreateInstance()
        {
            return new FRAGMENTSTREE_PLG_Plugin(this);
        }

        public override Guid ID
        {
            get { return new Guid("{D4DEDD3D-3C83-4E64-BFA8-C0D3E7308F28}"); }
        }

        public override string Name
        {
            get { return "Дерево фрагментов"; }
        }
    };

    enum Commands
    {
        Create = 1, //Команда создания
        Status =2,
        Debug =3,
    };

    class FRAGMENTSTREE_PLG_Plugin : Plugin
    {
        public FRAGMENTSTREE_PLG_Plugin(Factory factory) : base(factory)
        {
        }

        public static string regedit_str = @"Software\TF Plugins\FragmentsTree";

        public static ATTR_COM EXT_PAR;

        System.Drawing.Bitmap LoadBitmapResource(string name)
        {
            System.IO.Stream stream = GetType().Assembly.GetManifestResourceStream("FRAGMENTSTREE_PLG.Resource_Files." + name + ".bmp");
            return new System.Drawing.Bitmap(stream);
        }

        public System.Drawing.Icon LoadIconResource(string name)
        {
            System.IO.Stream stream = GetType().Assembly.GetManifestResourceStream("FRAGMENTSTREE_PLG.Resource_Files." + name + ".ico");
            return new System.Drawing.Icon(stream);
        }

        protected override void OnInitialize()
        {
            base.OnInitialize();
            EXT_PAR = new ATTR_COM();
        }

        protected override void OnCreateTools()
        {
            base.OnCreateTools();

            RegisterCommand((int)Commands.Create,
                "Разбор по фрагментам(сборка)", LoadIconResource("Коннектор_small"), LoadIconResource("Коннектор"));
            RegisterCommand((int)Commands.Status,
                "Генерирование по конфигурациям(деталь)", LoadIconResource("EXT_STATUS_small"), LoadIconResource("EXT_STATUS"));
            RegisterCommand((int)Commands.Debug,
                "Debug", LoadIconResource("Плюс"), LoadIconResource("Плюс"));

            int[] CmdIDs = new int[]
            {
                (int)Commands.Create,
                (int)Commands.Status,
                (int)Commands.Debug
            };

            TFlex.Menu submenu = new TFlex.Menu();
            submenu.CreatePopup();

            submenu.Append((int)Commands.Create, "&Разбор по фрагментам(сборка)", this);
            submenu.Append((int)Commands.Status, "&Генерирование по конфигурациям(деталь)", this);
            TFlex.RibbonGroup ribbonGroup = TFlex.RibbonBar.ApplicationsTab.AddGroup("Разборка");
            ribbonGroup.AddButton((int)Commands.Create, this);
            ribbonGroup.AddButton((int)Commands.Status, this);
            TFlex.Application.ActiveMainWindow.InsertPluginSubMenu(this.Name, submenu, TFlex.MainWindow.InsertMenuPosition.PluginSamples, this);

            CreateToolbar(this.Name, CmdIDs);
        }

        protected override void OnCommand(Document document, int id)
        {
            switch ((Commands)id)
            {
                default:
                    base.OnCommand(document, id);
                    break;

                case Commands.Create:
                    {
                        ComParams par = new ComParams(EXT_PAR);
                        if (par.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            par.SetParams(EXT_PAR);
                            CommandManager Command = new CommandManager();
                            Command.OK(document, EXT_PAR);
                        }
                        else System.Windows.Forms.MessageBox.Show("Выполнение плагина отменено", "Разбор по фрагментам(сборка)");
                        break;
                    }

                case Commands.Status:
                    {
                        ComParams par = new ComParams(EXT_PAR);
                        if (par.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            par.SetParams(EXT_PAR);
                            CommandManager ConfRun = new CommandManager();
                            ConfRun.OK2(document, EXT_PAR);
                        }
                        else System.Windows.Forms.MessageBox.Show("Выполнение плагина отменено", "Генерирование по конфигурациям(деталь)");
                        break;
                    }

                    case Commands.Debug:
                    {
                        /*if (document.Selection.GetSize() == 1)
                        {
                            document.Selection.
                            object obj = document.Selection.GetAt(0);
                            return;
                        }*/
                        break;
                    }
            }
        }

        protected override void OnUpdateCommand(CommandUI cmdUI)
        {
            if (cmdUI == null)
                return;

            if (cmdUI.Document == null)
            {
                cmdUI.Enable(false);
                return;
            }

            cmdUI.Enable();
        }

        protected override void NewDocumentCreatedEventHandler(DocumentEventArgs args)
        {
            args.Document.AttachPlugin(this);
        }

        protected override void DocumentOpenEventHandler(DocumentEventArgs args)
        {
            args.Document.AttachPlugin(this);
        }
    }
}
