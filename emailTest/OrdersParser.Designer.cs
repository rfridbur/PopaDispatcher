namespace Anko
{
    partial class OrdersParser
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OrdersParser));
            this.reports_btn = new System.Windows.Forms.Button();
            this.logTextBox = new System.Windows.Forms.RichTextBox();
            this.picbox = new System.Windows.Forms.PictureBox();
            this.loadConfirm_btn = new System.Windows.Forms.Button();
            this.bookConfirm_btn = new System.Windows.Forms.Button();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.arrivalsTab = new System.Windows.Forms.TabPage();
            this.haifaLinkLbl = new System.Windows.Forms.LinkLabel();
            this.ashdodLinkLbl = new System.Windows.Forms.LinkLabel();
            this.arrivals_lbl = new System.Windows.Forms.Label();
            this.arrivalsDataGrid = new System.Windows.Forms.DataGridView();
            this.sailsTab = new System.Windows.Forms.TabPage();
            this.sails_lbl = new System.Windows.Forms.Label();
            this.docReceipts_btn = new System.Windows.Forms.Button();
            this.sailsDataGrid = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.picbox)).BeginInit();
            this.tabControl.SuspendLayout();
            this.arrivalsTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.arrivalsDataGrid)).BeginInit();
            this.sailsTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sailsDataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // reports_btn
            // 
            this.reports_btn.BackColor = System.Drawing.SystemColors.Window;
            this.reports_btn.Font = new System.Drawing.Font("Kristen ITC", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.reports_btn.ForeColor = System.Drawing.Color.Magenta;
            this.reports_btn.Location = new System.Drawing.Point(16, 15);
            this.reports_btn.Margin = new System.Windows.Forms.Padding(4);
            this.reports_btn.Name = "reports_btn";
            this.reports_btn.Size = new System.Drawing.Size(272, 52);
            this.reports_btn.TabIndex = 1;
            this.reports_btn.Text = "Reports";
            this.reports_btn.UseVisualStyleBackColor = false;
            this.reports_btn.Click += new System.EventHandler(this.parse_btn_Click);
            // 
            // logTextBox
            // 
            this.logTextBox.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logTextBox.Location = new System.Drawing.Point(308, 333);
            this.logTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.logTextBox.Name = "logTextBox";
            this.logTextBox.Size = new System.Drawing.Size(1245, 164);
            this.logTextBox.TabIndex = 3;
            this.logTextBox.Text = "";
            // 
            // picbox
            // 
            this.picbox.Image = ((System.Drawing.Image)(resources.GetObject("picbox.Image")));
            this.picbox.Location = new System.Drawing.Point(-8, 175);
            this.picbox.Margin = new System.Windows.Forms.Padding(4);
            this.picbox.Name = "picbox";
            this.picbox.Size = new System.Drawing.Size(328, 373);
            this.picbox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picbox.TabIndex = 4;
            this.picbox.TabStop = false;
            // 
            // loadConfirm_btn
            // 
            this.loadConfirm_btn.BackColor = System.Drawing.SystemColors.Window;
            this.loadConfirm_btn.Font = new System.Drawing.Font("Kristen ITC", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loadConfirm_btn.ForeColor = System.Drawing.Color.Magenta;
            this.loadConfirm_btn.Location = new System.Drawing.Point(16, 75);
            this.loadConfirm_btn.Margin = new System.Windows.Forms.Padding(4);
            this.loadConfirm_btn.Name = "loadConfirm_btn";
            this.loadConfirm_btn.Size = new System.Drawing.Size(272, 52);
            this.loadConfirm_btn.TabIndex = 5;
            this.loadConfirm_btn.Text = "Loading Confirmation";
            this.loadConfirm_btn.UseVisualStyleBackColor = false;
            this.loadConfirm_btn.Click += new System.EventHandler(this.loadConfirm_btn_Click);
            // 
            // bookConfirm_btn
            // 
            this.bookConfirm_btn.BackColor = System.Drawing.SystemColors.Window;
            this.bookConfirm_btn.Font = new System.Drawing.Font("Kristen ITC", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bookConfirm_btn.ForeColor = System.Drawing.Color.Magenta;
            this.bookConfirm_btn.Location = new System.Drawing.Point(16, 135);
            this.bookConfirm_btn.Margin = new System.Windows.Forms.Padding(4);
            this.bookConfirm_btn.Name = "bookConfirm_btn";
            this.bookConfirm_btn.Size = new System.Drawing.Size(272, 52);
            this.bookConfirm_btn.TabIndex = 6;
            this.bookConfirm_btn.Text = "Booking Confirmation";
            this.bookConfirm_btn.UseVisualStyleBackColor = false;
            this.bookConfirm_btn.Click += new System.EventHandler(this.bookConfirm_btn_Click);
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.arrivalsTab);
            this.tabControl.Controls.Add(this.sailsTab);
            this.tabControl.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl.Location = new System.Drawing.Point(308, 15);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1245, 311);
            this.tabControl.TabIndex = 7;
            // 
            // arrivalsTab
            // 
            this.arrivalsTab.Controls.Add(this.haifaLinkLbl);
            this.arrivalsTab.Controls.Add(this.ashdodLinkLbl);
            this.arrivalsTab.Controls.Add(this.arrivals_lbl);
            this.arrivalsTab.Controls.Add(this.arrivalsDataGrid);
            this.arrivalsTab.Location = new System.Drawing.Point(4, 30);
            this.arrivalsTab.Name = "arrivalsTab";
            this.arrivalsTab.Padding = new System.Windows.Forms.Padding(3);
            this.arrivalsTab.Size = new System.Drawing.Size(1237, 277);
            this.arrivalsTab.TabIndex = 0;
            this.arrivalsTab.Text = "Today\'s arrivals";
            this.arrivalsTab.UseVisualStyleBackColor = true;
            // 
            // haifaLinkLbl
            // 
            this.haifaLinkLbl.AutoSize = true;
            this.haifaLinkLbl.Location = new System.Drawing.Point(1132, 51);
            this.haifaLinkLbl.Name = "haifaLinkLbl";
            this.haifaLinkLbl.Size = new System.Drawing.Size(80, 21);
            this.haifaLinkLbl.TabIndex = 3;
            this.haifaLinkLbl.TabStop = true;
            this.haifaLinkLbl.Text = "Haifa Port";
            // 
            // ashdodLinkLbl
            // 
            this.ashdodLinkLbl.AutoSize = true;
            this.ashdodLinkLbl.Location = new System.Drawing.Point(1132, 30);
            this.ashdodLinkLbl.Name = "ashdodLinkLbl";
            this.ashdodLinkLbl.Size = new System.Drawing.Size(97, 21);
            this.ashdodLinkLbl.TabIndex = 2;
            this.ashdodLinkLbl.TabStop = true;
            this.ashdodLinkLbl.Text = "Ashdod Port";
            // 
            // arrivals_lbl
            // 
            this.arrivals_lbl.AutoSize = true;
            this.arrivals_lbl.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.arrivals_lbl.Location = new System.Drawing.Point(6, 3);
            this.arrivals_lbl.Name = "arrivals_lbl";
            this.arrivals_lbl.Size = new System.Drawing.Size(86, 21);
            this.arrivals_lbl.TabIndex = 1;
            this.arrivals_lbl.Text = "arrivals_lbl";
            // 
            // arrivalsDataGrid
            // 
            this.arrivalsDataGrid.AllowUserToAddRows = false;
            this.arrivalsDataGrid.AllowUserToDeleteRows = false;
            this.arrivalsDataGrid.BackgroundColor = System.Drawing.SystemColors.Window;
            this.arrivalsDataGrid.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.arrivalsDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.arrivalsDataGrid.Location = new System.Drawing.Point(3, 30);
            this.arrivalsDataGrid.Name = "arrivalsDataGrid";
            this.arrivalsDataGrid.ReadOnly = true;
            this.arrivalsDataGrid.RowTemplate.Height = 24;
            this.arrivalsDataGrid.Size = new System.Drawing.Size(1067, 241);
            this.arrivalsDataGrid.TabIndex = 0;
            // 
            // sailsTab
            // 
            this.sailsTab.Controls.Add(this.sails_lbl);
            this.sailsTab.Controls.Add(this.docReceipts_btn);
            this.sailsTab.Controls.Add(this.sailsDataGrid);
            this.sailsTab.Location = new System.Drawing.Point(4, 30);
            this.sailsTab.Name = "sailsTab";
            this.sailsTab.Padding = new System.Windows.Forms.Padding(3);
            this.sailsTab.Size = new System.Drawing.Size(1237, 277);
            this.sailsTab.TabIndex = 1;
            this.sailsTab.Text = "Sails";
            this.sailsTab.UseVisualStyleBackColor = true;
            // 
            // sails_lbl
            // 
            this.sails_lbl.AutoSize = true;
            this.sails_lbl.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sails_lbl.Location = new System.Drawing.Point(6, 3);
            this.sails_lbl.Name = "sails_lbl";
            this.sails_lbl.Size = new System.Drawing.Size(65, 21);
            this.sails_lbl.TabIndex = 10;
            this.sails_lbl.Text = "sails_lbl";
            // 
            // docReceipts_btn
            // 
            this.docReceipts_btn.BackColor = System.Drawing.SystemColors.Window;
            this.docReceipts_btn.Font = new System.Drawing.Font("Kristen ITC", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.docReceipts_btn.ForeColor = System.Drawing.Color.Magenta;
            this.docReceipts_btn.Location = new System.Drawing.Point(1077, 30);
            this.docReceipts_btn.Margin = new System.Windows.Forms.Padding(4);
            this.docReceipts_btn.Name = "docReceipts_btn";
            this.docReceipts_btn.Size = new System.Drawing.Size(153, 83);
            this.docReceipts_btn.TabIndex = 9;
            this.docReceipts_btn.Text = "Documents Receipts";
            this.docReceipts_btn.UseVisualStyleBackColor = false;
            this.docReceipts_btn.Click += new System.EventHandler(this.docReceipts_btn_Click);
            // 
            // sailsDataGrid
            // 
            this.sailsDataGrid.AllowUserToAddRows = false;
            this.sailsDataGrid.AllowUserToDeleteRows = false;
            this.sailsDataGrid.BackgroundColor = System.Drawing.SystemColors.Window;
            this.sailsDataGrid.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.sailsDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sailsDataGrid.Location = new System.Drawing.Point(3, 30);
            this.sailsDataGrid.Name = "sailsDataGrid";
            this.sailsDataGrid.ReadOnly = true;
            this.sailsDataGrid.RowTemplate.Height = 24;
            this.sailsDataGrid.Size = new System.Drawing.Size(1067, 241);
            this.sailsDataGrid.TabIndex = 0;
            // 
            // OrdersParser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(1571, 515);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.bookConfirm_btn);
            this.Controls.Add(this.loadConfirm_btn);
            this.Controls.Add(this.logTextBox);
            this.Controls.Add(this.picbox);
            this.Controls.Add(this.reports_btn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "OrdersParser";
            this.RightToLeftLayout = true;
            this.Text = "Anko";
            ((System.ComponentModel.ISupportInitialize)(this.picbox)).EndInit();
            this.tabControl.ResumeLayout(false);
            this.arrivalsTab.ResumeLayout(false);
            this.arrivalsTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.arrivalsDataGrid)).EndInit();
            this.sailsTab.ResumeLayout(false);
            this.sailsTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sailsDataGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button reports_btn;
        private System.Windows.Forms.RichTextBox logTextBox;
        private System.Windows.Forms.PictureBox picbox;
        private System.Windows.Forms.Button loadConfirm_btn;
        private System.Windows.Forms.Button bookConfirm_btn;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage arrivalsTab;
        private System.Windows.Forms.TabPage sailsTab;
        private System.Windows.Forms.DataGridView arrivalsDataGrid;
        private System.Windows.Forms.DataGridView sailsDataGrid;
        private System.Windows.Forms.Button docReceipts_btn;
        private System.Windows.Forms.Label arrivals_lbl;
        private System.Windows.Forms.Label sails_lbl;
        private System.Windows.Forms.LinkLabel haifaLinkLbl;
        private System.Windows.Forms.LinkLabel ashdodLinkLbl;
    }
}