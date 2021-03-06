namespace Sql2Xl
{
    partial class frmSqlExport
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
            this.btnExlExport = new System.Windows.Forms.Button();
            this.txtSql = new System.Windows.Forms.TextBox();
            this.grdSummary = new System.Windows.Forms.DataGridView();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.cmbTables = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnShow = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnExportAll = new System.Windows.Forms.Button();
            this.cmbServers = new System.Windows.Forms.ComboBox();
            this.cmbDB = new System.Windows.Forms.ComboBox();
            this.chkExcludeSystem = new System.Windows.Forms.CheckBox();
            this.grpExportType = new System.Windows.Forms.GroupBox();
            this.optCSV = new System.Windows.Forms.RadioButton();
            this.optXl = new System.Windows.Forms.RadioButton();
            this.txtTimeOut = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.lblRecCount = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblVersion = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.grdSummary)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.grpExportType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeOut)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExlExport
            // 
            this.btnExlExport.Location = new System.Drawing.Point(12, 490);
            this.btnExlExport.Name = "btnExlExport";
            this.btnExlExport.Size = new System.Drawing.Size(85, 33);
            this.btnExlExport.TabIndex = 9;
            this.btnExlExport.Text = "Export Table";
            this.btnExlExport.UseVisualStyleBackColor = true;
            this.btnExlExport.Click += new System.EventHandler(this.btnExlExport_Click);
            // 
            // txtSql
            // 
            this.txtSql.Location = new System.Drawing.Point(12, 79);
            this.txtSql.Multiline = true;
            this.txtSql.Name = "txtSql";
            this.txtSql.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSql.Size = new System.Drawing.Size(776, 138);
            this.txtSql.TabIndex = 5;
            // 
            // grdSummary
            // 
            this.grdSummary.AllowUserToAddRows = false;
            this.grdSummary.AllowUserToDeleteRows = false;
            this.grdSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdSummary.Location = new System.Drawing.Point(12, 304);
            this.grdSummary.Name = "grdSummary";
            this.grdSummary.Size = new System.Drawing.Size(776, 167);
            this.grdSummary.TabIndex = 8;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2,
            this.lblVersion});
            this.statusStrip1.Location = new System.Drawing.Point(0, 542);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(858, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // cmbTables
            // 
            this.cmbTables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTables.FormattingEnabled = true;
            this.cmbTables.Location = new System.Drawing.Point(12, 267);
            this.cmbTables.Name = "cmbTables";
            this.cmbTables.Size = new System.Drawing.Size(208, 21);
            this.cmbTables.TabIndex = 7;
            this.cmbTables.SelectedIndexChanged += new System.EventHandler(this.cmbTables_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "SQL";
            // 
            // btnShow
            // 
            this.btnShow.Location = new System.Drawing.Point(12, 224);
            this.btnShow.Name = "btnShow";
            this.btnShow.Size = new System.Drawing.Size(75, 23);
            this.btnShow.TabIndex = 6;
            this.btnShow.Text = "Show Data";
            this.btnShow.UseVisualStyleBackColor = true;
            this.btnShow.Click += new System.EventHandler(this.btnShow_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Server";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(312, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(22, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "DB";
            // 
            // btnExportAll
            // 
            this.btnExportAll.Location = new System.Drawing.Point(157, 490);
            this.btnExportAll.Name = "btnExportAll";
            this.btnExportAll.Size = new System.Drawing.Size(85, 33);
            this.btnExportAll.TabIndex = 10;
            this.btnExportAll.Text = "Export All";
            this.btnExportAll.UseVisualStyleBackColor = true;
            this.btnExportAll.Click += new System.EventHandler(this.btnExportAll_Click);
            // 
            // cmbServers
            // 
            this.cmbServers.FormattingEnabled = true;
            this.cmbServers.Location = new System.Drawing.Point(56, 17);
            this.cmbServers.Name = "cmbServers";
            this.cmbServers.Size = new System.Drawing.Size(250, 21);
            this.cmbServers.TabIndex = 1;
            this.cmbServers.SelectedIndexChanged += new System.EventHandler(this.cmbServers_SelectedIndexChanged);
            this.cmbServers.TextChanged += new System.EventHandler(this.cmbServers_TextChanged);
            this.cmbServers.Validated += new System.EventHandler(this.cmbServers_Validated);
            // 
            // cmbDB
            // 
            this.cmbDB.FormattingEnabled = true;
            this.cmbDB.Location = new System.Drawing.Point(340, 17);
            this.cmbDB.Name = "cmbDB";
            this.cmbDB.Size = new System.Drawing.Size(250, 21);
            this.cmbDB.TabIndex = 2;
            this.cmbDB.SelectedIndexChanged += new System.EventHandler(this.cmbDB_SelectedIndexChanged);
            this.cmbDB.TextChanged += new System.EventHandler(this.cmbDB_TextChanged);
            // 
            // chkExcludeSystem
            // 
            this.chkExcludeSystem.AutoSize = true;
            this.chkExcludeSystem.Checked = true;
            this.chkExcludeSystem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExcludeSystem.Location = new System.Drawing.Point(340, 44);
            this.chkExcludeSystem.Name = "chkExcludeSystem";
            this.chkExcludeSystem.Size = new System.Drawing.Size(121, 17);
            this.chkExcludeSystem.TabIndex = 3;
            this.chkExcludeSystem.Text = "Exclude system db\'s";
            this.chkExcludeSystem.UseVisualStyleBackColor = true;
            this.chkExcludeSystem.CheckedChanged += new System.EventHandler(this.chkExcludeSystem_CheckedChanged);
            // 
            // grpExportType
            // 
            this.grpExportType.Controls.Add(this.optCSV);
            this.grpExportType.Controls.Add(this.optXl);
            this.grpExportType.Location = new System.Drawing.Point(469, 474);
            this.grpExportType.Name = "grpExportType";
            this.grpExportType.Size = new System.Drawing.Size(107, 65);
            this.grpExportType.TabIndex = 11;
            this.grpExportType.TabStop = false;
            this.grpExportType.Text = "ExportType";
            // 
            // optCSV
            // 
            this.optCSV.AutoSize = true;
            this.optCSV.Location = new System.Drawing.Point(12, 39);
            this.optCSV.Name = "optCSV";
            this.optCSV.Size = new System.Drawing.Size(46, 17);
            this.optCSV.TabIndex = 13;
            this.optCSV.TabStop = true;
            this.optCSV.Text = "CSV";
            this.optCSV.UseVisualStyleBackColor = true;
            this.optCSV.CheckedChanged += new System.EventHandler(this.optCSV_CheckedChanged);
            // 
            // optXl
            // 
            this.optXl.AutoSize = true;
            this.optXl.Location = new System.Drawing.Point(12, 16);
            this.optXl.Name = "optXl";
            this.optXl.Size = new System.Drawing.Size(51, 17);
            this.optXl.TabIndex = 12;
            this.optXl.TabStop = true;
            this.optXl.Text = "Excel";
            this.optXl.UseVisualStyleBackColor = true;
            this.optXl.CheckedChanged += new System.EventHandler(this.optXl_CheckedChanged);
            // 
            // txtTimeOut
            // 
            this.txtTimeOut.Location = new System.Drawing.Point(663, 19);
            this.txtTimeOut.Maximum = new decimal(new int[] {
            3600,
            0,
            0,
            0});
            this.txtTimeOut.Name = "txtTimeOut";
            this.txtTimeOut.Size = new System.Drawing.Size(55, 20);
            this.txtTimeOut.TabIndex = 4;
            this.txtTimeOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtTimeOut.Value = new decimal(new int[] {
            60,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(612, 21);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 17;
            this.label4.Text = "Timeout";
            // 
            // lblRecCount
            // 
            this.lblRecCount.AutoSize = true;
            this.lblRecCount.Location = new System.Drawing.Point(386, 270);
            this.lblRecCount.Name = "lblRecCount";
            this.lblRecCount.Size = new System.Drawing.Size(13, 13);
            this.lblRecCount.TabIndex = 18;
            this.lblRecCount.Text = "0";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(271, 270);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Record Count";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(45, 17);
            this.toolStripStatusLabel2.Text = "Version";
            // 
            // lblVersion
            // 
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(61, 17);
            this.lblVersion.Text = "versionNo";
            // 
            // frmSqlExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(858, 564);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblRecCount);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtTimeOut);
            this.Controls.Add(this.grpExportType);
            this.Controls.Add(this.btnExportAll);
            this.Controls.Add(this.btnExlExport);
            this.Controls.Add(this.chkExcludeSystem);
            this.Controls.Add(this.cmbDB);
            this.Controls.Add(this.cmbServers);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnShow);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbTables);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.grdSummary);
            this.Controls.Add(this.txtSql);
            this.Name = "frmSqlExport";
            this.Text = "Export SQL to Excel";
            this.Activated += new System.EventHandler(this.frmSqlExport_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmSqlExport_FormClosing);
            this.Load += new System.EventHandler(this.frmSqlExport_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdSummary)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.grpExportType.ResumeLayout(false);
            this.grpExportType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeOut)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExlExport;
        private System.Windows.Forms.TextBox txtSql;
        private System.Windows.Forms.DataGridView grdSummary;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ComboBox cmbTables;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnShow;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button btnExportAll;
        private System.Windows.Forms.ComboBox cmbServers;
        private System.Windows.Forms.ComboBox cmbDB;
        private System.Windows.Forms.CheckBox chkExcludeSystem;
        private System.Windows.Forms.GroupBox grpExportType;
        private System.Windows.Forms.RadioButton optCSV;
        private System.Windows.Forms.RadioButton optXl;
        private System.Windows.Forms.NumericUpDown txtTimeOut;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblRecCount;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel lblVersion;
    }
}

