namespace ExcelMySQLTool
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            this.menuStripMain = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainerMain = new System.Windows.Forms.SplitContainer();
            this.gbMySQLExport = new System.Windows.Forms.GroupBox();
            this.btnExportFromMySQL = new System.Windows.Forms.Button();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            this.lblEndDate = new System.Windows.Forms.Label();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.lblStartDate = new System.Windows.Forms.Label();
            this.txtDateColumn = new System.Windows.Forms.TextBox();
            this.lblDateColumn = new System.Windows.Forms.Label();
            this.chkFilterByDate = new System.Windows.Forms.CheckBox();
            this.lblSelectedExportFile = new System.Windows.Forms.Label();
            this.btnSelectExportLocation = new System.Windows.Forms.Button();
            this.gbExcelImport = new System.Windows.Forms.GroupBox();
            this.btnImportToMySQL = new System.Windows.Forms.Button();
            this.lblSelectedExcelFile = new System.Windows.Forms.Label();
            this.btnSelectExcelFile = new System.Windows.Forms.Button();
            this.gbDbConfig = new System.Windows.Forms.GroupBox();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.lblTableName = new System.Windows.Forms.Label();
            this.txtDbPassword = new System.Windows.Forms.TextBox();
            this.lblDbPassword = new System.Windows.Forms.Label();
            this.txtDbUser = new System.Windows.Forms.TextBox();
            this.lblDbUser = new System.Windows.Forms.Label();
            this.txtDbName = new System.Windows.Forms.TextBox();
            this.lblDbName = new System.Windows.Forms.Label();
            this.txtDbPort = new System.Windows.Forms.TextBox();
            this.lblDbPort = new System.Windows.Forms.Label();
            this.txtDbHost = new System.Windows.Forms.TextBox();
            this.lblDbHost = new System.Windows.Forms.Label();
            this.splitContainerRight = new System.Windows.Forms.SplitContainer();
            this.gbDataPreview = new System.Windows.Forms.GroupBox();
            this.dgvPreview = new System.Windows.Forms.DataGridView();
            this.cmbPreviewType = new System.Windows.Forms.ComboBox();
            this.gbLogOutput = new System.Windows.Forms.GroupBox();
            this.txtLogOutput = new System.Windows.Forms.TextBox();
            this.statusStripBottom = new System.Windows.Forms.StatusStrip();
            this.lblStatusTime = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblStatusMemory = new System.Windows.Forms.ToolStripStatusLabel();
            this.timerSystemHealth = new System.Windows.Forms.Timer(this.components);
            this.menuStripMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).BeginInit();
            this.splitContainerMain.Panel1.SuspendLayout();
            this.splitContainerMain.Panel2.SuspendLayout();
            this.splitContainerMain.SuspendLayout();
            this.gbMySQLExport.SuspendLayout();
            this.gbExcelImport.SuspendLayout();
            this.gbDbConfig.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerRight)).BeginInit();
            this.splitContainerRight.Panel1.SuspendLayout();
            this.splitContainerRight.Panel2.SuspendLayout();
            this.splitContainerRight.SuspendLayout();
            this.gbDataPreview.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPreview)).BeginInit();
            this.gbLogOutput.SuspendLayout();
            this.statusStripBottom.SuspendLayout();
            this.SuspendLayout();
            //
            // menuStripMain
            //
            this.menuStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStripMain.Location = new System.Drawing.Point(0, 0);
            this.menuStripMain.Name = "menuStripMain";
            this.menuStripMain.Size = new System.Drawing.Size(884, 24);
            this.menuStripMain.TabIndex = 0;
            this.menuStripMain.Text = "menuStrip1";
            //
            // fileToolStripMenuItem
            //
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            //
            // exitToolStripMenuItem
            //
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(93, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            //
            // helpToolStripMenuItem
            //
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            //
            // aboutToolStripMenuItem
            //
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.aboutToolStripMenuItem.Text = "&About";
            //
            // splitContainerMain
            //
            this.splitContainerMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerMain.Location = new System.Drawing.Point(0, 24);
            this.splitContainerMain.Name = "splitContainerMain";
            //
            // splitContainerMain.Panel1
            //
            this.splitContainerMain.Panel1.Controls.Add(this.gbMySQLExport);
            this.splitContainerMain.Panel1.Controls.Add(this.gbExcelImport);
            this.splitContainerMain.Panel1.Controls.Add(this.gbDbConfig);
            //
            // splitContainerMain.Panel2
            //
            this.splitContainerMain.Panel2.Controls.Add(this.splitContainerRight);
            this.splitContainerMain.Size = new System.Drawing.Size(884, 515);
            this.splitContainerMain.SplitterDistance = 589; // Approx 2/3 of 900
            this.splitContainerMain.TabIndex = 1;
            //
            // gbMySQLExport
            //
            this.gbMySQLExport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbMySQLExport.Controls.Add(this.btnExportFromMySQL);
            this.gbMySQLExport.Controls.Add(this.dtpEndDate);
            this.gbMySQLExport.Controls.Add(this.lblEndDate);
            this.gbMySQLExport.Controls.Add(this.dtpStartDate);
            this.gbMySQLExport.Controls.Add(this.lblStartDate);
            this.gbMySQLExport.Controls.Add(this.txtDateColumn);
            this.gbMySQLExport.Controls.Add(this.lblDateColumn);
            this.gbMySQLExport.Controls.Add(this.chkFilterByDate);
            this.gbMySQLExport.Controls.Add(this.lblSelectedExportFile);
            this.gbMySQLExport.Controls.Add(this.btnSelectExportLocation);
            this.gbMySQLExport.Location = new System.Drawing.Point(12, 335);
            this.gbMySQLExport.Name = "gbMySQLExport";
            this.gbMySQLExport.Size = new System.Drawing.Size(565, 170);
            this.gbMySQLExport.TabIndex = 2;
            this.gbMySQLExport.TabStop = false;
            this.gbMySQLExport.Text = "MySQL Export";
            //
            // btnExportFromMySQL
            //
            this.btnExportFromMySQL.Location = new System.Drawing.Point(9, 135);
            this.btnExportFromMySQL.Name = "btnExportFromMySQL";
            this.btnExportFromMySQL.Size = new System.Drawing.Size(150, 23);
            this.btnExportFromMySQL.TabIndex = 9;
            this.btnExportFromMySQL.Text = "Export from MySQL";
            this.btnExportFromMySQL.UseVisualStyleBackColor = true;
            //
            // dtpEndDate
            //
            this.dtpEndDate.Enabled = false;
            this.dtpEndDate.Location = new System.Drawing.Point(300, 105);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(200, 20);
            this.dtpEndDate.TabIndex = 8;
            //
            // lblEndDate
            //
            this.lblEndDate.AutoSize = true;
            this.lblEndDate.Location = new System.Drawing.Point(239, 108);
            this.lblEndDate.Name = "lblEndDate";
            this.lblEndDate.Size = new System.Drawing.Size(55, 13);
            this.lblEndDate.TabIndex = 7;
            this.lblEndDate.Text = "End Date:";
            //
            // dtpStartDate
            //
            this.dtpStartDate.Enabled = false;
            this.dtpStartDate.Location = new System.Drawing.Point(300, 79);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(200, 20);
            this.dtpStartDate.TabIndex = 6;
            //
            // lblStartDate
            //
            this.lblStartDate.AutoSize = true;
            this.lblStartDate.Location = new System.Drawing.Point(239, 82);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(58, 13);
            this.lblStartDate.TabIndex = 5;
            this.lblStartDate.Text = "Start Date:";
            //
            // txtDateColumn
            //
            this.txtDateColumn.Enabled = false;
            this.txtDateColumn.Location = new System.Drawing.Point(88, 105);
            this.txtDateColumn.Name = "txtDateColumn";
            this.txtDateColumn.Size = new System.Drawing.Size(130, 20);
            this.txtDateColumn.TabIndex = 4;
            //
            // lblDateColumn
            //
            this.lblDateColumn.AutoSize = true;
            this.lblDateColumn.Location = new System.Drawing.Point(6, 108);
            this.lblDateColumn.Name = "lblDateColumn";
            this.lblDateColumn.Size = new System.Drawing.Size(70, 13);
            this.lblDateColumn.TabIndex = 3;
            this.lblDateColumn.Text = "Date Column:";
            //
            // chkFilterByDate
            //
            this.chkFilterByDate.AutoSize = true;
            this.chkFilterByDate.Location = new System.Drawing.Point(9, 81);
            this.chkFilterByDate.Name = "chkFilterByDate";
            this.chkFilterByDate.Size = new System.Drawing.Size(112, 17);
            this.chkFilterByDate.TabIndex = 2;
            this.chkFilterByDate.Text = "Filter by Date Range";
            this.chkFilterByDate.UseVisualStyleBackColor = true;
            //
            // lblSelectedExportFile
            //
            this.lblSelectedExportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSelectedExportFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSelectedExportFile.Location = new System.Drawing.Point(9, 48);
            this.lblSelectedExportFile.Name = "lblSelectedExportFile";
            this.lblSelectedExportFile.Size = new System.Drawing.Size(550, 23);
            this.lblSelectedExportFile.TabIndex = 1;
            this.lblSelectedExportFile.Text = "No location selected";
            this.lblSelectedExportFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // btnSelectExportLocation
            //
            this.btnSelectExportLocation.Location = new System.Drawing.Point(9, 19);
            this.btnSelectExportLocation.Name = "btnSelectExportLocation";
            this.btnSelectExportLocation.Size = new System.Drawing.Size(200, 23);
            this.btnSelectExportLocation.TabIndex = 0;
            this.btnSelectExportLocation.Text = "Select Export Location & Format...";
            this.btnSelectExportLocation.UseVisualStyleBackColor = true;
            //
            // gbExcelImport
            //
            this.gbExcelImport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbExcelImport.Controls.Add(this.btnImportToMySQL);
            this.gbExcelImport.Controls.Add(this.lblSelectedExcelFile);
            this.gbExcelImport.Controls.Add(this.btnSelectExcelFile);
            this.gbExcelImport.Location = new System.Drawing.Point(12, 220);
            this.gbExcelImport.Name = "gbExcelImport";
            this.gbExcelImport.Size = new System.Drawing.Size(565, 100);
            this.gbExcelImport.TabIndex = 1;
            this.gbExcelImport.TabStop = false;
            this.gbExcelImport.Text = "Excel Import";
            //
            // btnImportToMySQL
            //
            this.btnImportToMySQL.Location = new System.Drawing.Point(9, 65);
            this.btnImportToMySQL.Name = "btnImportToMySQL";
            this.btnImportToMySQL.Size = new System.Drawing.Size(150, 23);
            this.btnImportToMySQL.TabIndex = 2;
            this.btnImportToMySQL.Text = "Import to MySQL";
            this.btnImportToMySQL.UseVisualStyleBackColor = true;
            //
            // lblSelectedExcelFile
            //
            this.lblSelectedExcelFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSelectedExcelFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSelectedExcelFile.Location = new System.Drawing.Point(9, 40);
            this.lblSelectedExcelFile.Name = "lblSelectedExcelFile";
            this.lblSelectedExcelFile.Size = new System.Drawing.Size(550, 23);
            this.lblSelectedExcelFile.TabIndex = 1;
            this.lblSelectedExcelFile.Text = "No file selected";
            this.lblSelectedExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // btnSelectExcelFile
            //
            this.btnSelectExcelFile.Location = new System.Drawing.Point(9, 19);
            this.btnSelectExcelFile.Name = "btnSelectExcelFile";
            this.btnSelectExcelFile.Size = new System.Drawing.Size(150, 23);
            this.btnSelectExcelFile.TabIndex = 0;
            this.btnSelectExcelFile.Text = "Select Excel File...";
            this.btnSelectExcelFile.UseVisualStyleBackColor = true;
            //
            // gbDbConfig
            //
            this.gbDbConfig.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbDbConfig.Controls.Add(this.txtTableName);
            this.gbDbConfig.Controls.Add(this.lblTableName);
            this.gbDbConfig.Controls.Add(this.txtDbPassword);
            this.gbDbConfig.Controls.Add(this.lblDbPassword);
            this.gbDbConfig.Controls.Add(this.txtDbUser);
            this.gbDbConfig.Controls.Add(this.lblDbUser);
            this.gbDbConfig.Controls.Add(this.txtDbName);
            this.gbDbConfig.Controls.Add(this.lblDbName);
            this.gbDbConfig.Controls.Add(this.txtDbPort);
            this.gbDbConfig.Controls.Add(this.lblDbPort);
            this.gbDbConfig.Controls.Add(this.txtDbHost);
            this.gbDbConfig.Controls.Add(this.lblDbHost);
            this.gbDbConfig.Location = new System.Drawing.Point(12, 15);
            this.gbDbConfig.Name = "gbDbConfig";
            this.gbDbConfig.Size = new System.Drawing.Size(565, 185);
            this.gbDbConfig.TabIndex = 0;
            this.gbDbConfig.TabStop = false;
            this.gbDbConfig.Text = "Database Configuration";
            //
            // txtTableName
            //
            this.txtTableName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTableName.Location = new System.Drawing.Point(88, 150);
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(471, 20);
            this.txtTableName.TabIndex = 11;
            //
            // lblTableName
            //
            this.lblTableName.AutoSize = true;
            this.lblTableName.Location = new System.Drawing.Point(6, 153);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(68, 13);
            this.lblTableName.TabIndex = 10;
            this.lblTableName.Text = "Table Name:";
            //
            // txtDbPassword
            //
            this.txtDbPassword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDbPassword.Location = new System.Drawing.Point(88, 124);
            this.txtDbPassword.Name = "txtDbPassword";
            this.txtDbPassword.Size = new System.Drawing.Size(471, 20);
            this.txtDbPassword.TabIndex = 9;
            this.txtDbPassword.UseSystemPasswordChar = true;
            //
            // lblDbPassword
            //
            this.lblDbPassword.AutoSize = true;
            this.lblDbPassword.Location = new System.Drawing.Point(6, 127);
            this.lblDbPassword.Name = "lblDbPassword";
            this.lblDbPassword.Size = new System.Drawing.Size(56, 13);
            this.lblDbPassword.TabIndex = 8;
            this.lblDbPassword.Text = "Password:";
            //
            // txtDbUser
            //
            this.txtDbUser.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDbUser.Location = new System.Drawing.Point(88, 98);
            this.txtDbUser.Name = "txtDbUser";
            this.txtDbUser.Size = new System.Drawing.Size(471, 20);
            this.txtDbUser.TabIndex = 7;
            //
            // lblDbUser
            //
            this.lblDbUser.AutoSize = true;
            this.lblDbUser.Location = new System.Drawing.Point(6, 101);
            this.lblDbUser.Name = "lblDbUser";
            this.lblDbUser.Size = new System.Drawing.Size(32, 13);
            this.lblDbUser.TabIndex = 6;
            this.lblDbUser.Text = "User:";
            //
            // txtDbName
            //
            this.txtDbName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDbName.Location = new System.Drawing.Point(88, 72);
            this.txtDbName.Name = "txtDbName";
            this.txtDbName.Size = new System.Drawing.Size(471, 20);
            this.txtDbName.TabIndex = 5;
            //
            // lblDbName
            //
            this.lblDbName.AutoSize = true;
            this.lblDbName.Location = new System.Drawing.Point(6, 75);
            this.lblDbName.Name = "lblDbName";
            this.lblDbName.Size = new System.Drawing.Size(87, 13);
            this.lblDbName.TabIndex = 4;
            this.lblDbName.Text = "Database Name:";
            //
            // txtDbPort
            //
            this.txtDbPort.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDbPort.Location = new System.Drawing.Point(88, 46);
            this.txtDbPort.Name = "txtDbPort";
            this.txtDbPort.Size = new System.Drawing.Size(471, 20);
            this.txtDbPort.TabIndex = 3;
            this.txtDbPort.Text = "3306";
            //
            // lblDbPort
            //
            this.lblDbPort.AutoSize = true;
            this.lblDbPort.Location = new System.Drawing.Point(6, 49);
            this.lblDbPort.Name = "lblDbPort";
            this.lblDbPort.Size = new System.Drawing.Size(29, 13);
            this.lblDbPort.TabIndex = 2;
            this.lblDbPort.Text = "Port:";
            //
            // txtDbHost
            //
            this.txtDbHost.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDbHost.Location = new System.Drawing.Point(88, 20);
            this.txtDbHost.Name = "txtDbHost";
            this.txtDbHost.Size = new System.Drawing.Size(471, 20);
            this.txtDbHost.TabIndex = 1;
            this.txtDbHost.Text = "localhost";
            //
            // lblDbHost
            //
            this.lblDbHost.AutoSize = true;
            this.lblDbHost.Location = new System.Drawing.Point(6, 23);
            this.lblDbHost.Name = "lblDbHost";
            this.lblDbHost.Size = new System.Drawing.Size(32, 13);
            this.lblDbHost.TabIndex = 0;
            this.lblDbHost.Text = "Host:";
            //
            // splitContainerRight
            //
            this.splitContainerRight.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerRight.Location = new System.Drawing.Point(0, 0);
            this.splitContainerRight.Name = "splitContainerRight";
            this.splitContainerRight.Orientation = System.Windows.Forms.Orientation.Horizontal;
            //
            // splitContainerRight.Panel1
            //
            this.splitContainerRight.Panel1.Controls.Add(this.gbDataPreview);
            //
            // splitContainerRight.Panel2
            //
            this.splitContainerRight.Panel2.Controls.Add(this.gbLogOutput);
            this.splitContainerRight.Size = new System.Drawing.Size(291, 515);
            this.splitContainerRight.SplitterDistance = 257; // Approx 1/2 of 515
            this.splitContainerRight.TabIndex = 0;
            //
            // gbDataPreview
            //
            this.gbDataPreview.Controls.Add(this.dgvPreview);
            this.gbDataPreview.Controls.Add(this.cmbPreviewType);
            this.gbDataPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gbDataPreview.Location = new System.Drawing.Point(0, 0);
            this.gbDataPreview.Name = "gbDataPreview";
            this.gbDataPreview.Size = new System.Drawing.Size(291, 257);
            this.gbDataPreview.TabIndex = 0;
            this.gbDataPreview.TabStop = false;
            this.gbDataPreview.Text = "Data Preview";
            //
            // dgvPreview
            //
            this.dgvPreview.AllowUserToAddRows = false;
            this.dgvPreview.AllowUserToDeleteRows = false;
            this.dgvPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvPreview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPreview.Location = new System.Drawing.Point(6, 46);
            this.dgvPreview.Name = "dgvPreview";
            this.dgvPreview.ReadOnly = true;
            this.dgvPreview.Size = new System.Drawing.Size(279, 205);
            this.dgvPreview.TabIndex = 1;
            //
            // cmbPreviewType
            //
            this.cmbPreviewType.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbPreviewType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPreviewType.FormattingEnabled = true;
            this.cmbPreviewType.Items.AddRange(new object[] {
            "Excel Preview",
            "MySQL Table Preview"});
            this.cmbPreviewType.Location = new System.Drawing.Point(6, 19);
            this.cmbPreviewType.Name = "cmbPreviewType";
            this.cmbPreviewType.Size = new System.Drawing.Size(279, 21);
            this.cmbPreviewType.TabIndex = 0;
            //
            // gbLogOutput
            //
            this.gbLogOutput.Controls.Add(this.txtLogOutput);
            this.gbLogOutput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gbLogOutput.Location = new System.Drawing.Point(0, 0);
            this.gbLogOutput.Name = "gbLogOutput";
            this.gbLogOutput.Size = new System.Drawing.Size(291, 254);
            this.gbLogOutput.TabIndex = 0;
            this.gbLogOutput.TabStop = false;
            this.gbLogOutput.Text = "Log Output";
            //
            // txtLogOutput
            //
            this.txtLogOutput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLogOutput.Location = new System.Drawing.Point(3, 16);
            this.txtLogOutput.Multiline = true;
            this.txtLogOutput.Name = "txtLogOutput";
            this.txtLogOutput.ReadOnly = true;
            this.txtLogOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLogOutput.Size = new System.Drawing.Size(285, 235);
            this.txtLogOutput.TabIndex = 0;
            //
            // statusStripBottom
            //
            this.statusStripBottom.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatusTime,
            this.lblStatusMemory});
            this.statusStripBottom.Location = new System.Drawing.Point(0, 539);
            this.statusStripBottom.Name = "statusStripBottom";
            this.statusStripBottom.Size = new System.Drawing.Size(884, 22);
            this.statusStripBottom.TabIndex = 2;
            this.statusStripBottom.Text = "statusStrip1";
            //
            // lblStatusTime
            //
            this.lblStatusTime.Name = "lblStatusTime";
            this.lblStatusTime.Size = new System.Drawing.Size(40, 17);
            this.lblStatusTime.Text = "Time: ";
            //
            // lblStatusMemory
            //
            this.lblStatusMemory.Name = "lblStatusMemory";
            this.lblStatusMemory.Size = new System.Drawing.Size(58, 17);
            this.lblStatusMemory.Text = "Memory: ";
            this.lblStatusMemory.Spring = true; // This will push it to the right if no other spring items exist
            //
            // timerSystemHealth
            //
            this.timerSystemHealth.Interval = 1000;
            //
            // MainForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(100)))), ((int)(((byte)(200)))));
            this.ClientSize = new System.Drawing.Size(884, 561); // Approx 900x600, adjusted for status bar
            this.Controls.Add(this.splitContainerMain);
            this.Controls.Add(this.menuStripMain);
            this.Controls.Add(this.statusStripBottom);
            this.MainMenuStrip = this.menuStripMain;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel-MySQL Data Utility";
            this.menuStripMain.ResumeLayout(false);
            this.menuStripMain.PerformLayout();
            this.splitContainerMain.Panel1.ResumeLayout(false);
            this.splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).EndInit();
            this.splitContainerMain.ResumeLayout(false);
            this.gbMySQLExport.ResumeLayout(false);
            this.gbMySQLExport.PerformLayout();
            this.gbExcelImport.ResumeLayout(false);
            this.gbExcelImport.PerformLayout();
            this.gbDbConfig.ResumeLayout(false);
            this.gbDbConfig.PerformLayout();
            this.splitContainerRight.Panel1.ResumeLayout(false);
            this.splitContainerRight.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerRight)).EndInit();
            this.splitContainerRight.ResumeLayout(false);
            this.gbDataPreview.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPreview)).EndInit();
            this.gbLogOutput.ResumeLayout(false);
            this.gbLogOutput.PerformLayout();
            this.statusStripBottom.ResumeLayout(false);
            this.statusStripBottom.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStripMain;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.GroupBox gbDbConfig;
        private System.Windows.Forms.TextBox txtDbHost;
        private System.Windows.Forms.Label lblDbHost;
        private System.Windows.Forms.TextBox txtDbPort;
        private System.Windows.Forms.Label lblDbPort;
        private System.Windows.Forms.TextBox txtDbName;
        private System.Windows.Forms.Label lblDbName;
        private System.Windows.Forms.TextBox txtDbUser;
        private System.Windows.Forms.Label lblDbUser;
        private System.Windows.Forms.TextBox txtDbPassword;
        private System.Windows.Forms.Label lblDbPassword;
        private System.Windows.Forms.TextBox txtTableName;
        private System.Windows.Forms.Label lblTableName;
        private System.Windows.Forms.GroupBox gbExcelImport;
        private System.Windows.Forms.Button btnSelectExcelFile;
        private System.Windows.Forms.Label lblSelectedExcelFile;
        private System.Windows.Forms.Button btnImportToMySQL;
        private System.Windows.Forms.GroupBox gbMySQLExport;
        private System.Windows.Forms.Button btnSelectExportLocation;
        private System.Windows.Forms.Label lblSelectedExportFile;
        private System.Windows.Forms.CheckBox chkFilterByDate;
        private System.Windows.Forms.TextBox txtDateColumn;
        private System.Windows.Forms.Label lblDateColumn;
        private System.Windows.Forms.DateTimePicker dtpStartDate;
        private System.Windows.Forms.Label lblStartDate;
        private System.Windows.Forms.DateTimePicker dtpEndDate;
        private System.Windows.Forms.Label lblEndDate;
        private System.Windows.Forms.Button btnExportFromMySQL;
        private System.Windows.Forms.SplitContainer splitContainerRight;
        private System.Windows.Forms.GroupBox gbDataPreview;
        private System.Windows.Forms.ComboBox cmbPreviewType;
        private System.Windows.Forms.DataGridView dgvPreview;
        private System.Windows.Forms.GroupBox gbLogOutput;
        private System.Windows.Forms.TextBox txtLogOutput;
        private System.Windows.Forms.StatusStrip statusStripBottom;
        private System.Windows.Forms.ToolStripStatusLabel lblStatusTime;
        private System.Windows.Forms.ToolStripStatusLabel lblStatusMemory;
        private System.Windows.Forms.Timer timerSystemHealth;
    }
}
