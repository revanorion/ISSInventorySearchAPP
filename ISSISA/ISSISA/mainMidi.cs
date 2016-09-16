using ISSISA;
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
        public mainMidi()
        {
            InitializeComponent();
        }

        private void iASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form IAS = new InventoryForm();            
            IAS.Show();
        }
    }
}
