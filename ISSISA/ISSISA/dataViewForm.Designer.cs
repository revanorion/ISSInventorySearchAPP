namespace ISSIAS
{
    partial class dataViewForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources =
                new System.ComponentModel.ComponentResourceManager(typeof(dataViewForm));
            this.assetDataView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize) (this.assetDataView)).BeginInit();
            this.SuspendLayout();
            // 
            // assetDataView
            // 
            this.assetDataView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.assetDataView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.assetDataView.ColumnHeadersHeightSizeMode =
                System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.assetDataView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.assetDataView.Location = new System.Drawing.Point(0, 0);
            this.assetDataView.Name = "assetDataView";
            this.assetDataView.Size = new System.Drawing.Size(904, 508);
            this.assetDataView.TabIndex = 0;
            this.assetDataView.CellContentClick +=
                new System.Windows.Forms.DataGridViewCellEventHandler(this.assetDataView_CellContentClick);
            // 
            // dataViewForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.ClientSize = new System.Drawing.Size(904, 508);
            this.Controls.Add(this.assetDataView);
            this.Icon = ((System.Drawing.Icon) (resources.GetObject("$this.Icon")));
            this.Name = "dataViewForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "dataViewForm";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize) (this.assetDataView)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.DataGridView assetDataView;
    }
}