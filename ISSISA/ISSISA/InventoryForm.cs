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

namespace ISSISA
{
    public partial class InventoryForm : Form
    {
        //these are two dialog windows that will appear when user clicks on add or save buttons
        OpenFileDialog ofd = new OpenFileDialog();
        SaveFileDialog sfd = new SaveFileDialog();


        FileConnections files = new FileConnections();

        //these are binding sources that will show the file names of selected files
        BindingSource filesSelectedBinding = new BindingSource();
        BindingSource finishedFilesBinding = new BindingSource();

        public InventoryForm()
        {
            InitializeComponent();

            //binds the data source to the list
            filesSelectedBinding.DataSource = files.files;
            files_selected_list.DataSource = filesSelectedBinding;
            files_selected_list.DisplayMember = "name";		//only takes 1 property
            files_selected_list.ValueMember = "name";

            fiscal_book_label.Text = files.fiscal_book_address;            

            //binds the data source to the list
            finishedFilesBinding.DataSource = files.finished_files;
            finished_files_list.DataSource = finishedFilesBinding;
            finished_files_list.DisplayMember = "name";		//only takes 1 property
            finished_files_list.ValueMember = "name";

            MaximizeBox = false;
            MinimizeBox = false;
            this.files_selected_list.BackColor = Color.AliceBlue;
            this.finished_files_list.BackColor = Color.AliceBlue;

            this.save_files_button.FlatStyle = FlatStyle.Popup;
            this.save_files_button.BackColor = Color.AliceBlue;
            this.run_button.FlatStyle = FlatStyle.Popup;
            this.run_button.BackColor = Color.AliceBlue;
            this.add_file_button.FlatStyle = FlatStyle.Popup;
            this.add_file_button.BackColor = Color.AliceBlue;
            this.remove_book_button.FlatStyle = FlatStyle.Popup;
            this.remove_book_button.BackColor = Color.AliceBlue;
            this.remove_file_button.FlatStyle = FlatStyle.Popup;
            this.remove_file_button.BackColor = Color.AliceBlue;
            this.fiscal_book_button.FlatStyle = FlatStyle.Popup;
            this.fiscal_book_button.BackColor = Color.AliceBlue;


            this.BackColor = Color.Cornsilk;


        }

        private void InventoryForm_Load(object sender, EventArgs e)
        {

        }

        //event handeler for add file button. dialog will appear for user to select files 
        private void add_file_button_Click(object sender, EventArgs e)
        {
            ofd.Filter = "CSV Files (.csv)|*.csv|Text Files (.txt)|*.txt|All Files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string x in ofd.FileNames)
                {
                    files.add_file(x);
                }
                filesSelectedBinding.ResetBindings(false);
            }

        }

        //event handeler for remove file button. will remove the selected file from the list
        private void remove_file_button_Click(object sender, EventArgs e)
        {
            if (files_selected_list.SelectedItem != null)
            {
                fileNaming x = ((fileNaming)files_selected_list.SelectedItem);
                files.remove_file(x);
                filesSelectedBinding.ResetBindings(false);
                files.finished_files.Clear();
                finishedFilesBinding.ResetBindings(false);
            }
            else
            {
                MessageBox.Show("No File Selected");
            }
        }

        //event handeler for fiscal book button. dialog will appear for user to select fiscal book 
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

        //event handeler for remove fiscal book button. will remove the selected fiscal book file
        private void remove_book_button_Click(object sender, EventArgs e)
        {
            if (files.fiscal_book_address != null)
            {
                files.fiscal_book_address = null;
                fiscal_book_label.Text = files.fiscal_book_address;
                files.finished_files.Clear();
                finishedFilesBinding.ResetBindings(false);
            }
            else
            {
                MessageBox.Show("No Fiscal Book");
            }
        }

        //event handeler for remove run button. will import all data from files list and fiscal book then compare the sets
        private void run_button_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (files.files.Count >= 1 && files.fiscal_book_address != "No File Selected!")
            {
                //at some point have it check to see if a specific serial exists rather than deleting lists. reimport handeling
                files.imported_devices.Clear();
                files.fb_assets.Clear();
                files.finished_files.Clear();
                try
                {
                    files.import_data();
                    finishedFilesBinding.ResetBindings(false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Don't have open the files selected! \n" + ex.Message);
                }
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
            Cursor.Current = Cursors.Default;
        }

        //event handeler for save file button. will take the data from compared list and write it to an excel file
        private void save_files_button_Click(object sender, EventArgs e)
        {

            /*
            This needs a dialog for saving missing files

            */
            
            Cursor.Current = Cursors.WaitCursor;
            sfd.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
            sfd.FilterIndex = 1;
            if (files.found_devices != null)
            {
                //This save Process is for found devices
                sfd.FileName = files.finished_files[0].name;
                if (sfd.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        files.write_to_excel(sfd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                //This save process is for missing devices
                sfd.FileName = files.finished_files[1].name;
                if (sfd.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        files.write_missing_to_excel(sfd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("No process ran or no devices found!");
            }
            Cursor.Current = Cursors.Default;
        }

        private void displayHelp(object sender, MouseEventArgs e)
        {
            MessageBox.Show(@"Selected files must be csv or txt and must contain in the file name any of the values below.

Tropos Export Data
Wireless_Controllers                
aps_wireless
UPS
Brocade switch");
        }
    }
}
