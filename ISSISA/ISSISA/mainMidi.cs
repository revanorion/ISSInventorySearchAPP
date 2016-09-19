using ISSISA;
using ISSISA_Library;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ISSIAS
{
    public partial class mainMidi : Form
    {

        private Form IAS_importedDataView;
        private InventoryForm IAS;
        private Form IAS_FBDataView;
        private Form IAS_foundDataView;
        private Form IAS_missingDataView;
        private FileConnections files;
        public mainMidi()
        {
            InitializeComponent();
            files = new FileConnections();
            IAS = new InventoryForm(files);
            IAS_importedDataView = new dataViewForm(files, 'I');
            IAS_FBDataView = new dataViewForm(files, 'B');
            IAS_foundDataView = new dataViewForm(files, 'F');
            IAS_missingDataView = new dataViewForm(files, 'M');

            IAS.MdiParent = this;
            IAS_FBDataView.MdiParent = this;
            IAS_foundDataView.MdiParent = this;
            IAS_missingDataView.MdiParent = this;
            IAS_importedDataView.MdiParent = this;

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.Cornsilk;
                }
                else if (ctrl is System.Windows.Forms.PictureBox)
                {
                    ctrl.BackColor = Color.Cornsilk;
                }

            }
        }

        private void iASToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (IAS.IsDisposed)
            {
                IAS = new InventoryForm(files);
                IAS.MdiParent = this;
            }

            hideChildren();
            IAS.Show();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void viewImportedDevicesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_importedDataView.IsDisposed)
            {
                IAS_importedDataView = new dataViewForm(files, 'I');
                IAS_importedDataView.MdiParent = this;
            }
            hideChildren();
            IAS_importedDataView.Show();
        }

        private void viewFBDevicesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_FBDataView.IsDisposed)
            {
                IAS_FBDataView = new dataViewForm(files, 'B');
                IAS_FBDataView.MdiParent = this;
            }
            hideChildren();
            IAS_FBDataView.Show();
        }


        private void viewFoundtoolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_foundDataView.IsDisposed)
            {
                IAS_foundDataView = new dataViewForm(files, 'F');
                IAS_foundDataView.MdiParent = this;
            }
            hideChildren();
            IAS_foundDataView.Show();
        }

        private void viewMissingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_missingDataView.IsDisposed)
            {
                IAS_missingDataView = new dataViewForm(files, 'M');
                IAS_missingDataView.MdiParent = this;
            }
            hideChildren();
            IAS_missingDataView.Show();
        }


        private void hideChildren()
        {
            foreach (Form i in this.MdiChildren)
                i.Hide();
        }

    }
}
