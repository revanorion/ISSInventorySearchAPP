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

        private dataViewForm IAS_importedDataView;
        private InventoryForm IAS;
        private dataViewForm IAS_FBDataView;
        private dataViewForm IAS_foundDataView;
        private dataViewForm IAS_missingDataView;

        private dataViewForm IAS_locationValidateDataView;
        private dataViewForm IAS_serialValidateDataView;
        private dataViewForm IAS_roomValidateDataView;


        private readonly FileConnections _files;
        public mainMidi()
        {
            InitializeComponent();
            _files = new FileConnections();
            IAS = new InventoryForm(_files);
            IAS_importedDataView = new dataViewForm(_files, 'I');
            IAS_FBDataView = new dataViewForm(_files, 'B');
            IAS_foundDataView = new dataViewForm(_files, 'F');
            IAS_missingDataView = new dataViewForm(_files, 'M');
            IAS_locationValidateDataView = new dataViewForm(_files, 'L');
            IAS_serialValidateDataView = new dataViewForm(_files, 'S');
            IAS_roomValidateDataView = new dataViewForm(_files, 'R');

            IAS.MdiParent = this;
            IAS_FBDataView.MdiParent = this;
            IAS_foundDataView.MdiParent = this;
            IAS_missingDataView.MdiParent = this;
            IAS_importedDataView.MdiParent = this;
            IAS_locationValidateDataView.MdiParent = this;
            IAS_serialValidateDataView.MdiParent = this;
            IAS_roomValidateDataView.MdiParent = this;

            foreach (Control ctrl in Controls)
            {
                if (ctrl is MdiClient)
                {
                    ctrl.BackColor = Color.Cornsilk;
                }
                else if (ctrl is PictureBox)
                {
                    ctrl.BackColor = Color.Cornsilk;
                }

            }
        }

        private void iASToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (IAS.IsDisposed)
            {
                IAS = new InventoryForm(_files) { MdiParent = this };
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
                IAS_importedDataView = new dataViewForm(_files, 'I');
                IAS_importedDataView.MdiParent = this;
            }
            hideChildren();
            IAS_importedDataView.Show();
        }

        private void viewFBDevicesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_FBDataView.IsDisposed)
            {
                IAS_FBDataView = new dataViewForm(_files, 'B');
                IAS_FBDataView.MdiParent = this;
            }
            hideChildren();
            IAS_FBDataView.Show();
        }


        private void viewFoundtoolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_foundDataView.IsDisposed)
            {
                IAS_foundDataView = new dataViewForm(_files, 'F');
                IAS_foundDataView.MdiParent = this;
            }
            hideChildren();
            IAS_foundDataView.Show();
        }

        private void viewMissingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_missingDataView.IsDisposed)
            {
                IAS_missingDataView = new dataViewForm(_files, 'M');
                IAS_missingDataView.MdiParent = this;
            }
            hideChildren();
            IAS_missingDataView.Show();
        }


        private void viewLocationValidateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_locationValidateDataView.IsDisposed)
            {
                IAS_locationValidateDataView = new dataViewForm(_files, 'L');
                IAS_locationValidateDataView.MdiParent = this;
            }
            hideChildren();
            IAS_locationValidateDataView.Refresh(_files.locationValidate_devices);
            IAS_locationValidateDataView.Show();

        }
        private void viewSerialValidateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_serialValidateDataView.IsDisposed)
            {
                IAS_serialValidateDataView = new dataViewForm(_files, 'S');
                IAS_serialValidateDataView.MdiParent = this;
            }
            hideChildren();
            IAS_serialValidateDataView.Refresh(_files.serialValidate_devices);
            IAS_serialValidateDataView.Show();
        }
        private void viewRoomValidateToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (IAS_roomValidateDataView.IsDisposed)
            {
                IAS_roomValidateDataView = new dataViewForm(_files, 'R');
                IAS_roomValidateDataView.MdiParent = this;
            }
            hideChildren();
            IAS_roomValidateDataView.Refresh(_files.roomValidate_devices);
            IAS_roomValidateDataView.Show();
        }


        private void hideChildren()
        {
            foreach (Form i in MdiChildren)
                i.Hide();
        }

        private void mainMidi_Load(object sender, EventArgs e)
        {
            IAS.Show();
        }

    }
}
