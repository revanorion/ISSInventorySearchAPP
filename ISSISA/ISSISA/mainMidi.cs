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
            IAS_locationValidateDataView = new dataViewForm(files, 'L');
            IAS_serialValidateDataView = new dataViewForm(files, 'S');
            IAS_roomValidateDataView = new dataViewForm(files, 'R');

            IAS.MdiParent = this;
            IAS_FBDataView.MdiParent = this;
            IAS_foundDataView.MdiParent = this;
            IAS_missingDataView.MdiParent = this;
            IAS_importedDataView.MdiParent = this;
            IAS_locationValidateDataView.MdiParent = this;
            IAS_serialValidateDataView.MdiParent = this;
            IAS_roomValidateDataView.MdiParent = this;

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


        private void viewLocationValidateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_locationValidateDataView.IsDisposed)
            {
                IAS_locationValidateDataView = new dataViewForm(files, 'L');
                IAS_locationValidateDataView.MdiParent = this;
            }
            hideChildren();
            IAS_locationValidateDataView.refresh(files.locationValidate_devices);
            IAS_locationValidateDataView.Show();

        }
        private void viewSerialValidateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (IAS_serialValidateDataView.IsDisposed)
            {
                IAS_serialValidateDataView = new dataViewForm(files, 'S');
                IAS_serialValidateDataView.MdiParent = this;
            }
            hideChildren();
            IAS_serialValidateDataView.refresh(files.serialValidate_devices);
            IAS_serialValidateDataView.Show();
        }
        private void viewRoomValidateToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (IAS_roomValidateDataView.IsDisposed)
            {
                IAS_roomValidateDataView = new dataViewForm(files, 'R');
                IAS_roomValidateDataView.MdiParent = this;
            }
            hideChildren();            
            IAS_roomValidateDataView.refresh(files.roomValidate_devices);
            IAS_roomValidateDataView.Show();
        }


        private void hideChildren()
        {
            foreach (Form i in this.MdiChildren)
                i.Hide();
        }

        private void mainMidi_Load(object sender, EventArgs e)
        {
            IAS.Show();
        }

    }
}
