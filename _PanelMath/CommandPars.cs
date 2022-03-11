using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FRAGMENTSTREE_PLG
{
    public partial class ComParams : Form
    {
        public ComParams(ATTR_COM par)
        {
            InitializeComponent();

            pEXSTEP.CheckState = par.pSTEP == 1 ? CheckState.Checked : CheckState.Unchecked;
            pEXDXF.CheckState = par.pDXF == 1 ? CheckState.Checked : CheckState.Unchecked;
            pEXPDF.CheckState = par.pPDF == 1 ? CheckState.Checked : CheckState.Unchecked;
        }

        public void SetParams(ATTR_COM par)
        {
            par.pDXF = pEXDXF.CheckState == CheckState.Checked ? 1 : 0;
            par.pSTEP = pEXSTEP.CheckState == CheckState.Checked ? 1 : 0;
            par.pPDF = pEXPDF.CheckState == CheckState.Checked ? 1 : 0;

            RegistryKey test = Registry.CurrentUser.OpenSubKey(FRAGMENTSTREE_PLG_Plugin.regedit_str, RegistryKeyPermissionCheck.ReadWriteSubTree);

            if (test == null)
            {
                test = Registry.CurrentUser.CreateSubKey(FRAGMENTSTREE_PLG_Plugin.regedit_str);
            }

            test.SetValue("ATTR_COM", par.attr.ToString());
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            /*par.pSTEP = 0;
            par.pDXF = 0;
            par.pPDF = 0;*/
            return;
        }
    }
}
