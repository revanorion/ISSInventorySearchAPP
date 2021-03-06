﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ISSIAS_Library;

namespace ISSIAS
{
    public partial class dataViewForm : Form
    {
        private readonly BindingSource _assetData = new BindingSource();

        public dataViewForm()
        {
            InitializeComponent();
            assetDataView.DataSource = _assetData;
        }

        //the switch will bind data to appropreate asset list. I for imported devices, B for Fiscal Book, 
        //F for found devices, M for missing devices.
        public dataViewForm(FileConnections a, char type)
        {
            InitializeComponent();
            var files = a;
            switch (type)
            {
                case 'I':
                    _assetData.DataSource = files.imported_devices;
                    break;
                case 'B':
                    _assetData.DataSource = files.fb_assets;
                    break;
                case 'F':
                    _assetData.DataSource = files.found_devices;
                    break;
                case 'M':
                    _assetData.DataSource = files.missing_devices;
                    break;
                case 'L':
                    _assetData.DataSource = files.locationValidate_devices;
                    break;
                case 'S':
                    _assetData.DataSource = files.serialValidate_devices;
                    break;
                case 'R':
                    _assetData.DataSource = files.roomValidate_devices;
                    break;
            }


            assetDataView.DataSource = _assetData;
        }

        private void assetDataView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        public void Refresh(BindingList<asset> assetData)
        {
            assetDataView.DataSource = null;
            assetDataView.Update();
            assetDataView.Refresh();
            _assetData.DataSource = assetData;
            assetDataView.DataSource = assetData;
        }
    }
}