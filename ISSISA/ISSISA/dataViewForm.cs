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
    public partial class dataViewForm : Form
    {
        private BindingSource assetData = new BindingSource();
        private FileConnections files = new FileConnections();

        public dataViewForm()
        {
            InitializeComponent();
            this.assetDataView.DataSource = assetData;
        }

        //the switch will bind data to appropreate asset list. I for imported devices, B for Fiscal Book, 
        //F for found devices, M for missing devices.
        public dataViewForm(FileConnections a, char type)
        {
            InitializeComponent();
            files=a;          
            switch (type)
            {
                case 'I':
                    assetData.DataSource = files.imported_devices;
                    break;
                case 'B':
                    assetData.DataSource = files.fb_assets;
                    break;
                case 'F':
                    assetData.DataSource = files.found_devices;
                    break;
                case 'M':
                    assetData.DataSource = files.missing_devices;
                    break;                
            }           
               
            
            this.assetDataView.DataSource = assetData;
        }

        private void assetDataView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }      
    }
}
