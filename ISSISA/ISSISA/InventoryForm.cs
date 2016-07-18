using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ISSISA_Library;

namespace ISSISA
{
    public partial class InventoryForm : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        FileConnections files = new FileConnections();
        BindingSource filesSelectedBinding = new BindingSource();
        

        public InventoryForm()
        {
            InitializeComponent();

            filesSelectedBinding.DataSource = files.files;
            files_selected_list.DataSource = filesSelectedBinding;
            files_selected_list.DisplayMember = "name";		//only takes 1 property
            files_selected_list.ValueMember = "name";

            fiscal_book_label.Text = files.fiscal_book_address;
        }

          private void InventoryForm_Load(object sender, EventArgs e)
        {

        }

        private void add_file_button_Click(object sender, EventArgs e)
        {
            ofd.Filter = "CSV Files (.csv)|*.csv|Text Files (.txt)|*.txt|All Files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach(string x in ofd.FileNames)
                {
                    files.add_file(x);
                }
                foreach(var s in files.files)
                {
                    Console.WriteLine(s.name);
                }
                filesSelectedBinding.ResetBindings(false);
            }

        }

        private void remove_file_button_Click(object sender, EventArgs e)
        {
            if(files_selected_list.SelectedItem != null)
            {
                fileNaming x = ((fileNaming)files_selected_list.SelectedItem);
                files.remove_file(x);
                filesSelectedBinding.ResetBindings(false);
            }
            else
            {
                MessageBox.Show("No File Selected");
            }
        }

        private void open_file_button_Click(object sender, EventArgs e)
        {
            if (files_selected_list.SelectedItem != null)
            {
                fileNaming x = ((fileNaming)files_selected_list.SelectedItem);
                try
                {
                    files.open_file(x);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No File Selected");
            }
        }

        private void fiscal_book_button_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    files.fiscal_book_address = ofd.FileName;
                    fiscal_book_label.Text = files.fiscal_book_address;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void remove_book_button_Click(object sender, EventArgs e)
        {
            if(files.fiscal_book_address!=null)
            {
                files.fiscal_book_address = null;
                fiscal_book_label.Text = files.fiscal_book_address;
            }
            else
            {
                MessageBox.Show("No Fiscal Book");
            }
        }

        private void run_button_Click(object sender, EventArgs e)
        {
            if (files.files.Count >= 1 && files.fiscal_book_address != "No File Selected!")
            {
                //run stuff
                files.import_data();
            }
            else if (files.files.Count < 1)
            {
                MessageBox.Show("Files must be added to run!");
            }
            else if (files.fiscal_book_address == "No File Selected!")
            {
                MessageBox.Show("No Fiscal Book Selected!");
            }
            else
            {
                MessageBox.Show("Unhandled Exception!");
            }
        }
    }
}
