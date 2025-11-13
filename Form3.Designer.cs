namespace DAT_ToolReports
{
    partial class Form3
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
            this.btnOpen = new System.Windows.Forms.Button();
            this.dgvSessionsExcels = new System.Windows.Forms.DataGridView();
            this.stt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MaHV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TenHV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MaPH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TGTruyen = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TGBatDau = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TGKetThuc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TgDaoTao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QDDaoTao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BienSoXe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TrungHV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TrungXe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ViPham = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnCheck = new System.Windows.Forms.Button();
            this.txtLogs = new System.Windows.Forms.TextBox();
            this.ckbTrungHV = new System.Windows.Forms.CheckBox();
            this.ckbTrungXe = new System.Windows.Forms.CheckBox();
            this.txtMaxDuration = new System.Windows.Forms.TextBox();
            this.txtMinDuration = new System.Windows.Forms.TextBox();
            this.ckbMax = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ckbMin = new System.Windows.Forms.CheckBox();
            this.btnFilter = new System.Windows.Forms.Button();
            this.btnSaveExcel = new System.Windows.Forms.Button();
            this.txtMinDupMinute = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSessionsExcels)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(12, 12);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(94, 29);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // dgvSessionsExcels
            // 
            this.dgvSessionsExcels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSessionsExcels.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.stt,
            this.ID,
            this.MaHV,
            this.TenHV,
            this.MaPH,
            this.TGTruyen,
            this.TGBatDau,
            this.TGKetThuc,
            this.TgDaoTao,
            this.QDDaoTao,
            this.BienSoXe,
            this.TrungHV,
            this.TrungXe,
            this.ViPham});
            this.dgvSessionsExcels.Location = new System.Drawing.Point(12, 47);
            this.dgvSessionsExcels.Name = "dgvSessionsExcels";
            this.dgvSessionsExcels.RowHeadersVisible = false;
            this.dgvSessionsExcels.RowHeadersWidth = 51;
            this.dgvSessionsExcels.RowTemplate.Height = 29;
            this.dgvSessionsExcels.Size = new System.Drawing.Size(1572, 519);
            this.dgvSessionsExcels.TabIndex = 1;
            // 
            // stt
            // 
            this.stt.HeaderText = "STT";
            this.stt.MinimumWidth = 6;
            this.stt.Name = "stt";
            this.stt.Width = 70;
            // 
            // ID
            // 
            this.ID.HeaderText = "ID";
            this.ID.MinimumWidth = 6;
            this.ID.Name = "ID";
            this.ID.Width = 70;
            // 
            // MaHV
            // 
            this.MaHV.HeaderText = "Mã Học Viên";
            this.MaHV.MinimumWidth = 6;
            this.MaHV.Name = "MaHV";
            this.MaHV.Width = 150;
            // 
            // TenHV
            // 
            this.TenHV.HeaderText = "Tên Học Viên";
            this.TenHV.MinimumWidth = 6;
            this.TenHV.Name = "TenHV";
            this.TenHV.Width = 200;
            // 
            // MaPH
            // 
            this.MaPH.HeaderText = "Mã Phiên Học";
            this.MaPH.MinimumWidth = 6;
            this.MaPH.Name = "MaPH";
            this.MaPH.Width = 300;
            // 
            // TGTruyen
            // 
            this.TGTruyen.HeaderText = "Thời gian truyền";
            this.TGTruyen.MinimumWidth = 6;
            this.TGTruyen.Name = "TGTruyen";
            this.TGTruyen.Width = 180;
            // 
            // TGBatDau
            // 
            this.TGBatDau.HeaderText = "Bắt đầu";
            this.TGBatDau.MinimumWidth = 6;
            this.TGBatDau.Name = "TGBatDau";
            this.TGBatDau.Width = 180;
            // 
            // TGKetThuc
            // 
            this.TGKetThuc.HeaderText = "Kết thúc";
            this.TGKetThuc.MinimumWidth = 6;
            this.TGKetThuc.Name = "TGKetThuc";
            this.TGKetThuc.Width = 180;
            // 
            // TgDaoTao
            // 
            this.TgDaoTao.HeaderText = "TG đào tạo";
            this.TgDaoTao.MinimumWidth = 6;
            this.TgDaoTao.Name = "TgDaoTao";
            this.TgDaoTao.Width = 90;
            // 
            // QDDaoTao
            // 
            this.QDDaoTao.HeaderText = "QĐ đào tạo";
            this.QDDaoTao.MinimumWidth = 6;
            this.QDDaoTao.Name = "QDDaoTao";
            this.QDDaoTao.Width = 90;
            // 
            // BienSoXe
            // 
            this.BienSoXe.HeaderText = "Biển số xe";
            this.BienSoXe.MinimumWidth = 6;
            this.BienSoXe.Name = "BienSoXe";
            this.BienSoXe.Width = 80;
            // 
            // TrungHV
            // 
            this.TrungHV.HeaderText = "Trùng Học Viên";
            this.TrungHV.MinimumWidth = 6;
            this.TrungHV.Name = "TrungHV";
            this.TrungHV.Width = 125;
            // 
            // TrungXe
            // 
            this.TrungXe.HeaderText = "Trùng xe";
            this.TrungXe.MinimumWidth = 6;
            this.TrungXe.Name = "TrungXe";
            this.TrungXe.Width = 125;
            // 
            // ViPham
            // 
            this.ViPham.HeaderText = "Vi phạm";
            this.ViPham.MinimumWidth = 6;
            this.ViPham.Name = "ViPham";
            this.ViPham.Width = 125;
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(112, 12);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(94, 29);
            this.btnCheck.TabIndex = 2;
            this.btnCheck.Text = "Kiểm tra";
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // txtLogs
            // 
            this.txtLogs.Location = new System.Drawing.Point(12, 572);
            this.txtLogs.Multiline = true;
            this.txtLogs.Name = "txtLogs";
            this.txtLogs.Size = new System.Drawing.Size(1572, 58);
            this.txtLogs.TabIndex = 3;
            // 
            // ckbTrungHV
            // 
            this.ckbTrungHV.AutoSize = true;
            this.ckbTrungHV.Location = new System.Drawing.Point(212, 17);
            this.ckbTrungHV.Name = "ckbTrungHV";
            this.ckbTrungHV.Size = new System.Drawing.Size(133, 24);
            this.ckbTrungHV.TabIndex = 4;
            this.ckbTrungHV.Text = "Check trùng HV";
            this.ckbTrungHV.UseVisualStyleBackColor = true;
            // 
            // ckbTrungXe
            // 
            this.ckbTrungXe.AutoSize = true;
            this.ckbTrungXe.Location = new System.Drawing.Point(351, 17);
            this.ckbTrungXe.Name = "ckbTrungXe";
            this.ckbTrungXe.Size = new System.Drawing.Size(128, 24);
            this.ckbTrungXe.TabIndex = 5;
            this.ckbTrungXe.Text = "Check trùng xe";
            this.ckbTrungXe.UseVisualStyleBackColor = true;
            // 
            // txtMaxDuration
            // 
            this.txtMaxDuration.Location = new System.Drawing.Point(664, 14);
            this.txtMaxDuration.Name = "txtMaxDuration";
            this.txtMaxDuration.Size = new System.Drawing.Size(63, 27);
            this.txtMaxDuration.TabIndex = 6;
            // 
            // txtMinDuration
            // 
            this.txtMinDuration.Location = new System.Drawing.Point(971, 14);
            this.txtMinDuration.Name = "txtMinDuration";
            this.txtMinDuration.Size = new System.Drawing.Size(63, 27);
            this.txtMinDuration.TabIndex = 7;
            // 
            // ckbMax
            // 
            this.ckbMax.AutoSize = true;
            this.ckbMax.Location = new System.Drawing.Point(485, 15);
            this.ckbMax.Name = "ckbMax";
            this.ckbMax.Size = new System.Drawing.Size(173, 24);
            this.ckbMax.TabIndex = 8;
            this.ckbMax.Text = "Bỏ qua phiên lớn hơn";
            this.ckbMax.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(733, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 20);
            this.label1.TabIndex = 9;
            this.label1.Text = "phút";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1040, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 20);
            this.label2.TabIndex = 10;
            this.label2.Text = "phút";
            // 
            // ckbMin
            // 
            this.ckbMin.AutoSize = true;
            this.ckbMin.Location = new System.Drawing.Point(788, 17);
            this.ckbMin.Name = "ckbMin";
            this.ckbMin.Size = new System.Drawing.Size(177, 24);
            this.ckbMin.TabIndex = 11;
            this.ckbMin.Text = "Bỏ qua phiên nhỏ hơn";
            this.ckbMin.UseVisualStyleBackColor = true;
            // 
            // btnFilter
            // 
            this.btnFilter.Location = new System.Drawing.Point(1402, 14);
            this.btnFilter.Name = "btnFilter";
            this.btnFilter.Size = new System.Drawing.Size(60, 29);
            this.btnFilter.TabIndex = 12;
            this.btnFilter.Text = "Lọc";
            this.btnFilter.UseVisualStyleBackColor = true;
            this.btnFilter.Click += new System.EventHandler(this.btnFilter_Click);
            // 
            // btnSaveExcel
            // 
            this.btnSaveExcel.Location = new System.Drawing.Point(1468, 14);
            this.btnSaveExcel.Name = "btnSaveExcel";
            this.btnSaveExcel.Size = new System.Drawing.Size(108, 29);
            this.btnSaveExcel.TabIndex = 13;
            this.btnSaveExcel.Text = "Save to Excel";
            this.btnSaveExcel.UseVisualStyleBackColor = true;
            this.btnSaveExcel.Click += new System.EventHandler(this.btnSaveExcel_Click);
            // 
            // txtMinDupMinute
            // 
            this.txtMinDupMinute.Location = new System.Drawing.Point(1259, 14);
            this.txtMinDupMinute.Name = "txtMinDupMinute";
            this.txtMinDupMinute.Size = new System.Drawing.Size(80, 27);
            this.txtMinDupMinute.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1095, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(158, 20);
            this.label3.TabIndex = 15;
            this.label3.Text = "Số phút trùng tối thiểu";
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1588, 637);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtMinDupMinute);
            this.Controls.Add(this.btnSaveExcel);
            this.Controls.Add(this.btnFilter);
            this.Controls.Add(this.ckbMin);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ckbMax);
            this.Controls.Add(this.txtMinDuration);
            this.Controls.Add(this.txtMaxDuration);
            this.Controls.Add(this.ckbTrungXe);
            this.Controls.Add(this.ckbTrungHV);
            this.Controls.Add(this.txtLogs);
            this.Controls.Add(this.btnCheck);
            this.Controls.Add(this.dgvSessionsExcels);
            this.Controls.Add(this.btnOpen);
            this.Name = "Form3";
            this.Text = "Form3";
            ((System.ComponentModel.ISupportInitialize)(this.dgvSessionsExcels)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button btnOpen;
        private DataGridView dgvSessionsExcels;
        private Button btnCheck;
        private TextBox txtLogs;
        private DataGridViewTextBoxColumn stt;
        private DataGridViewTextBoxColumn ID;
        private DataGridViewTextBoxColumn MaHV;
        private DataGridViewTextBoxColumn TenHV;
        private DataGridViewTextBoxColumn MaPH;
        private DataGridViewTextBoxColumn TGTruyen;
        private DataGridViewTextBoxColumn TGBatDau;
        private DataGridViewTextBoxColumn TGKetThuc;
        private DataGridViewTextBoxColumn TgDaoTao;
        private DataGridViewTextBoxColumn QDDaoTao;
        private DataGridViewTextBoxColumn BienSoXe;
        private DataGridViewTextBoxColumn TrungHV;
        private DataGridViewTextBoxColumn TrungXe;
        private DataGridViewTextBoxColumn ViPham;
        private CheckBox ckbTrungHV;
        private CheckBox ckbTrungXe;
        private TextBox txtMaxDuration;
        private TextBox txtMinDuration;
        private CheckBox ckbMax;
        private Label label1;
        private Label label2;
        private CheckBox ckbMin;
        private Button btnFilter;
        private Button btnSaveExcel;
        private TextBox txtMinDupMinute;
        private Label label3;
    }
}