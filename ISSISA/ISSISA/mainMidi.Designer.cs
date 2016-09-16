namespace ISSIAS
{
    partial class mainMidi
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
            this.menuStrip2 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.viewDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
            this.viewFBDevicesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewFoundToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewMissingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip2
            // 
            this.menuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.toolStripMenuItem2});
            this.menuStrip2.Location = new System.Drawing.Point(0, 0);
            this.menuStrip2.Name = "menuStrip2";
            this.menuStrip2.Size = new System.Drawing.Size(768, 24);
            this.menuStrip2.TabIndex = 1;
            this.menuStrip2.Text = "menuStrip2";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem3});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(37, 20);
            this.toolStripMenuItem1.Text = "File";
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(152, 22);
            this.toolStripMenuItem3.Text = "Exit";
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem5,
            this.viewDataToolStripMenuItem});
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(80, 20);
            this.toolStripMenuItem2.Text = "Application";
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(152, 22);
            this.toolStripMenuItem5.Text = "IAS";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.iASToolStripMenuItem_Click);
            // 
            // viewDataToolStripMenuItem
            // 
            this.viewDataToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem6,
            this.viewFBDevicesToolStripMenuItem,
            this.viewFoundToolStripMenuItem,
            this.viewMissingToolStripMenuItem});
            this.viewDataToolStripMenuItem.Name = "viewDataToolStripMenuItem";
            this.viewDataToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.viewDataToolStripMenuItem.Text = "View Data";
            // 
            // toolStripMenuItem6
            // 
            this.toolStripMenuItem6.Name = "toolStripMenuItem6";
            this.toolStripMenuItem6.Size = new System.Drawing.Size(194, 22);
            this.toolStripMenuItem6.Text = "View Imported Devices";
            this.toolStripMenuItem6.Click += new System.EventHandler(this.viewImportedDevicesToolStripMenuItem_Click);
            // 
            // viewFBDevicesToolStripMenuItem
            // 
            this.viewFBDevicesToolStripMenuItem.Name = "viewFBDevicesToolStripMenuItem";
            this.viewFBDevicesToolStripMenuItem.Size = new System.Drawing.Size(194, 22);
            this.viewFBDevicesToolStripMenuItem.Text = "View FB Devices";
            this.viewFBDevicesToolStripMenuItem.Click += new System.EventHandler(this.viewFBDevicesToolStripMenuItem_Click);
            // 
            // viewFoundToolStripMenuItem
            // 
            this.viewFoundToolStripMenuItem.Name = "viewFoundToolStripMenuItem";
            this.viewFoundToolStripMenuItem.Size = new System.Drawing.Size(194, 22);
            this.viewFoundToolStripMenuItem.Text = "View Found";
            this.viewFoundToolStripMenuItem.Click += new System.EventHandler(this.viewFoundtoolStripMenuItem_Click);
            // 
            // viewMissingToolStripMenuItem
            // 
            this.viewMissingToolStripMenuItem.Name = "viewMissingToolStripMenuItem";
            this.viewMissingToolStripMenuItem.Size = new System.Drawing.Size(194, 22);
            this.viewMissingToolStripMenuItem.Text = "View Missing";
            this.viewMissingToolStripMenuItem.Click += new System.EventHandler(this.viewMissingToolStripMenuItem_Click);
            // 
            // mainMidi
            // 
            this.ClientSize = new System.Drawing.Size(768, 584);
            this.Controls.Add(this.menuStrip2);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip2;
            this.Name = "mainMidi";
            this.menuStrip2.ResumeLayout(false);
            this.menuStrip2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip mainToolStrip;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem updateToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem iASToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem applicationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewDataToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem viewImportedDevicesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewFBDevicesToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem viewFoundDevicesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewMissingDevicesToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem viewDataToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem viewFBDevicesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewFoundToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewMissingToolStripMenuItem;
    }
}