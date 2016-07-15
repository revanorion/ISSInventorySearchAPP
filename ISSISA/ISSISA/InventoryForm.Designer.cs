namespace ISSISA
{
    partial class InventoryForm
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
            this.main_header_label = new System.Windows.Forms.Label();
            this.sub_header_label = new System.Windows.Forms.Label();
            this.files_selected_list = new System.Windows.Forms.ListBox();
            this.add_file_button = new System.Windows.Forms.Button();
            this.remove_file_button = new System.Windows.Forms.Button();
            this.finished_files_list = new System.Windows.Forms.ListBox();
            this.save_files_button = new System.Windows.Forms.Button();
            this.open_file_button = new System.Windows.Forms.Button();
            this.fiscal_book_button = new System.Windows.Forms.Button();
            this.remove_book_button = new System.Windows.Forms.Button();
            this.fiscal_book_label = new System.Windows.Forms.Label();
            this.files_header_label = new System.Windows.Forms.Label();
            this.run_button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // main_header_label
            // 
            this.main_header_label.AutoSize = true;
            this.main_header_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.main_header_label.Location = new System.Drawing.Point(200, 30);
            this.main_header_label.Name = "main_header_label";
            this.main_header_label.Size = new System.Drawing.Size(396, 33);
            this.main_header_label.TabIndex = 0;
            this.main_header_label.Text = "ISS Inventory Asset Search";
            // 
            // sub_header_label
            // 
            this.sub_header_label.AutoSize = true;
            this.sub_header_label.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sub_header_label.Location = new System.Drawing.Point(376, 63);
            this.sub_header_label.Name = "sub_header_label";
            this.sub_header_label.Size = new System.Drawing.Size(45, 25);
            this.sub_header_label.TabIndex = 1;
            this.sub_header_label.Text = "IAS";
            // 
            // files_selected_list
            // 
            this.files_selected_list.FormattingEnabled = true;
            this.files_selected_list.Location = new System.Drawing.Point(12, 160);
            this.files_selected_list.Name = "files_selected_list";
            this.files_selected_list.Size = new System.Drawing.Size(249, 329);
            this.files_selected_list.TabIndex = 2;
            // 
            // add_file_button
            // 
            this.add_file_button.Location = new System.Drawing.Point(12, 495);
            this.add_file_button.Name = "add_file_button";
            this.add_file_button.Size = new System.Drawing.Size(75, 23);
            this.add_file_button.TabIndex = 3;
            this.add_file_button.Text = "Add File(s)";
            this.add_file_button.UseVisualStyleBackColor = true;
            this.add_file_button.Click += new System.EventHandler(this.add_file_button_Click);
            // 
            // remove_file_button
            // 
            this.remove_file_button.Location = new System.Drawing.Point(186, 495);
            this.remove_file_button.Name = "remove_file_button";
            this.remove_file_button.Size = new System.Drawing.Size(75, 23);
            this.remove_file_button.TabIndex = 4;
            this.remove_file_button.Text = "Remove File";
            this.remove_file_button.UseVisualStyleBackColor = true;
            this.remove_file_button.Click += new System.EventHandler(this.remove_file_button_Click);
            // 
            // finished_files_list
            // 
            this.finished_files_list.FormattingEnabled = true;
            this.finished_files_list.Location = new System.Drawing.Point(535, 160);
            this.finished_files_list.Name = "finished_files_list";
            this.finished_files_list.Size = new System.Drawing.Size(249, 329);
            this.finished_files_list.TabIndex = 5;
            // 
            // save_files_button
            // 
            this.save_files_button.Location = new System.Drawing.Point(622, 495);
            this.save_files_button.Name = "save_files_button";
            this.save_files_button.Size = new System.Drawing.Size(75, 23);
            this.save_files_button.TabIndex = 6;
            this.save_files_button.Text = "Save File(s)";
            this.save_files_button.UseVisualStyleBackColor = true;
            // 
            // open_file_button
            // 
            this.open_file_button.Location = new System.Drawing.Point(93, 495);
            this.open_file_button.Name = "open_file_button";
            this.open_file_button.Size = new System.Drawing.Size(87, 23);
            this.open_file_button.TabIndex = 7;
            this.open_file_button.Text = "Open File";
            this.open_file_button.UseVisualStyleBackColor = true;
            this.open_file_button.Click += new System.EventHandler(this.open_file_button_Click);
            // 
            // fiscal_book_button
            // 
            this.fiscal_book_button.Location = new System.Drawing.Point(282, 495);
            this.fiscal_book_button.Name = "fiscal_book_button";
            this.fiscal_book_button.Size = new System.Drawing.Size(109, 23);
            this.fiscal_book_button.TabIndex = 8;
            this.fiscal_book_button.Text = "Select Fiscal Book";
            this.fiscal_book_button.UseVisualStyleBackColor = true;
            this.fiscal_book_button.Click += new System.EventHandler(this.fiscal_book_button_Click);
            // 
            // remove_book_button
            // 
            this.remove_book_button.Location = new System.Drawing.Point(397, 495);
            this.remove_book_button.Name = "remove_book_button";
            this.remove_book_button.Size = new System.Drawing.Size(117, 23);
            this.remove_book_button.TabIndex = 9;
            this.remove_book_button.Text = "Remove Fiscal Book";
            this.remove_book_button.UseVisualStyleBackColor = true;
            this.remove_book_button.Click += new System.EventHandler(this.remove_book_button_Click);
            // 
            // fiscal_book_label
            // 
            this.fiscal_book_label.AutoSize = true;
            this.fiscal_book_label.Location = new System.Drawing.Point(354, 475);
            this.fiscal_book_label.Name = "fiscal_book_label";
            this.fiscal_book_label.Size = new System.Drawing.Size(88, 13);
            this.fiscal_book_label.TabIndex = 10;
            this.fiscal_book_label.Text = "No File Selected!";
            // 
            // files_header_label
            // 
            this.files_header_label.AutoSize = true;
            this.files_header_label.Location = new System.Drawing.Point(29, 131);
            this.files_header_label.Name = "files_header_label";
            this.files_header_label.Size = new System.Drawing.Size(212, 26);
            this.files_header_label.TabIndex = 11;
            this.files_header_label.Text = "Select files to compare against Fiscal Book.\r\nDo not select fiscal book.";
            this.files_header_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // run_button
            // 
            this.run_button.Location = new System.Drawing.Point(361, 259);
            this.run_button.Name = "run_button";
            this.run_button.Size = new System.Drawing.Size(75, 23);
            this.run_button.TabIndex = 12;
            this.run_button.Text = "Run";
            this.run_button.UseVisualStyleBackColor = true;
            this.run_button.Click += new System.EventHandler(this.run_button_Click);
            // 
            // InventoryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(796, 530);
            this.Controls.Add(this.run_button);
            this.Controls.Add(this.files_header_label);
            this.Controls.Add(this.fiscal_book_label);
            this.Controls.Add(this.remove_book_button);
            this.Controls.Add(this.fiscal_book_button);
            this.Controls.Add(this.open_file_button);
            this.Controls.Add(this.save_files_button);
            this.Controls.Add(this.finished_files_list);
            this.Controls.Add(this.remove_file_button);
            this.Controls.Add(this.add_file_button);
            this.Controls.Add(this.files_selected_list);
            this.Controls.Add(this.sub_header_label);
            this.Controls.Add(this.main_header_label);
            this.Name = "InventoryForm";
            this.Text = "IAS";
            this.Load += new System.EventHandler(this.InventoryForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label main_header_label;
        private System.Windows.Forms.Label sub_header_label;
        private System.Windows.Forms.ListBox files_selected_list;
        private System.Windows.Forms.Button add_file_button;
        private System.Windows.Forms.Button remove_file_button;
        private System.Windows.Forms.ListBox finished_files_list;
        private System.Windows.Forms.Button save_files_button;
        private System.Windows.Forms.Button open_file_button;
        private System.Windows.Forms.Button fiscal_book_button;
        private System.Windows.Forms.Button remove_book_button;
        private System.Windows.Forms.Label fiscal_book_label;
        private System.Windows.Forms.Label files_header_label;
        private System.Windows.Forms.Button run_button;
    }
}

