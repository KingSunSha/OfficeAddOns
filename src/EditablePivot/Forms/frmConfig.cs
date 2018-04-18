using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EditablePivot.Forms
{
    public partial class frmConfig : Form
    {
        public bool OK = false;
        public bool EditEnabled = false;

        public frmConfig()
        {
            InitializeComponent();
        }

        public void Init(PivotTable pt) {
            chkEditEnabled.Checked = (pt.Tag == "EditEnabled");
        }

        private void tbtnOK_Click(object sender, EventArgs e)
        {
            OK = true;
            this.Hide();
        }

        private void tbtnCancel_Click(object sender, EventArgs e)
        {
            OK = false;
            this.Hide();
        }

        private void chkEditEnabled_CheckedChanged(object sender, EventArgs e)
        {
            EditEnabled = chkEditEnabled.Checked;
        }
    }
}
