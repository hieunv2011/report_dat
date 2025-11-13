namespace DAT_ToolReports
{
    partial class Form2
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
            this.txtCompany = new System.Windows.Forms.TextBox();
            this.txtCentre = new System.Windows.Forms.TextBox();
            this.txtLinkSever = new System.Windows.Forms.TextBox();
            this.txtProvince = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbxHour1 = new System.Windows.Forms.ComboBox();
            this.cbxMinute1 = new System.Windows.Forms.ComboBox();
            this.cbxHour2 = new System.Windows.Forms.ComboBox();
            this.cbxMinute2 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.ckbCountATforNight = new System.Windows.Forms.CheckBox();
            this.ckbNightByStart = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // txtCompany
            // 
            this.txtCompany.Location = new System.Drawing.Point(168, 95);
            this.txtCompany.Name = "txtCompany";
            this.txtCompany.Size = new System.Drawing.Size(443, 27);
            this.txtCompany.TabIndex = 0;
            // 
            // txtCentre
            // 
            this.txtCentre.Location = new System.Drawing.Point(168, 128);
            this.txtCentre.Name = "txtCentre";
            this.txtCentre.Size = new System.Drawing.Size(443, 27);
            this.txtCentre.TabIndex = 1;
            // 
            // txtLinkSever
            // 
            this.txtLinkSever.Location = new System.Drawing.Point(168, 161);
            this.txtLinkSever.Name = "txtLinkSever";
            this.txtLinkSever.Size = new System.Drawing.Size(443, 27);
            this.txtLinkSever.TabIndex = 2;
            // 
            // txtProvince
            // 
            this.txtProvince.Location = new System.Drawing.Point(168, 62);
            this.txtProvince.Name = "txtProvince";
            this.txtProvince.Size = new System.Drawing.Size(125, 27);
            this.txtProvince.TabIndex = 3;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(327, 398);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(149, 29);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(493, 398);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(142, 29);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "Tỉnh";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(34, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 20);
            this.label2.TabIndex = 7;
            this.label2.Text = "Cơ quan chủ quản";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(34, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 20);
            this.label3.TabIndex = 8;
            this.label3.Text = "Tên trung tâm ĐT";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(34, 164);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 20);
            this.label4.TabIndex = 9;
            this.label4.Text = "Link sever";
            // 
            // cbxHour1
            // 
            this.cbxHour1.FormattingEnabled = true;
            this.cbxHour1.Items.AddRange(new object[] {
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23"});
            this.cbxHour1.Location = new System.Drawing.Point(168, 194);
            this.cbxHour1.Name = "cbxHour1";
            this.cbxHour1.Size = new System.Drawing.Size(43, 28);
            this.cbxHour1.TabIndex = 10;
            // 
            // cbxMinute1
            // 
            this.cbxMinute1.FormattingEnabled = true;
            this.cbxMinute1.Items.AddRange(new object[] {
            "0",
            "15",
            "30",
            "45"});
            this.cbxMinute1.Location = new System.Drawing.Point(254, 194);
            this.cbxMinute1.Name = "cbxMinute1";
            this.cbxMinute1.Size = new System.Drawing.Size(43, 28);
            this.cbxMinute1.TabIndex = 11;
            // 
            // cbxHour2
            // 
            this.cbxHour2.FormattingEnabled = true;
            this.cbxHour2.Items.AddRange(new object[] {
            "3",
            "4",
            "5",
            "6",
            "7"});
            this.cbxHour2.Location = new System.Drawing.Point(377, 194);
            this.cbxHour2.Name = "cbxHour2";
            this.cbxHour2.Size = new System.Drawing.Size(43, 28);
            this.cbxHour2.TabIndex = 12;
            // 
            // cbxMinute2
            // 
            this.cbxMinute2.FormattingEnabled = true;
            this.cbxMinute2.Items.AddRange(new object[] {
            "0",
            "15",
            "30",
            "45"});
            this.cbxMinute2.Location = new System.Drawing.Point(463, 194);
            this.cbxMinute2.Name = "cbxMinute2";
            this.cbxMinute2.Size = new System.Drawing.Size(43, 28);
            this.cbxMinute2.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(34, 197);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 20);
            this.label5.TabIndex = 14;
            this.label5.Text = "Giờ đêm là từ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(217, 197);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(31, 20);
            this.label6.TabIndex = 15;
            this.label6.Text = "giờ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(303, 197);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 20);
            this.label7.TabIndex = 16;
            this.label7.Text = "phút đến";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(426, 197);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(31, 20);
            this.label8.TabIndex = 17;
            this.label8.Text = "giờ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(512, 197);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(99, 20);
            this.label9.TabIndex = 18;
            this.label9.Text = "phút hôm sau";
            // 
            // ckbCountATforNight
            // 
            this.ckbCountATforNight.AutoSize = true;
            this.ckbCountATforNight.Location = new System.Drawing.Point(415, 61);
            this.ckbCountATforNight.Name = "ckbCountATforNight";
            this.ckbCountATforNight.Size = new System.Drawing.Size(194, 24);
            this.ckbCountATforNight.TabIndex = 19;
            this.ckbCountATforNight.Text = "Tính giờ AT vào giờ đêm";
            this.ckbCountATforNight.UseVisualStyleBackColor = true;
            // 
            // ckbNightByStart
            // 
            this.ckbNightByStart.AutoSize = true;
            this.ckbNightByStart.Location = new System.Drawing.Point(168, 228);
            this.ckbNightByStart.Name = "ckbNightByStart";
            this.ckbNightByStart.Size = new System.Drawing.Size(345, 24);
            this.ckbNightByStart.TabIndex = 20;
            this.ckbNightByStart.Text = "Tính giờ đêm theo thời điểm bắt đầu phiên học";
            this.ckbNightByStart.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(726, 450);
            this.Controls.Add(this.ckbNightByStart);
            this.Controls.Add(this.ckbCountATforNight);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cbxMinute2);
            this.Controls.Add(this.cbxHour2);
            this.Controls.Add(this.cbxMinute1);
            this.Controls.Add(this.cbxHour1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.txtProvince);
            this.Controls.Add(this.txtLinkSever);
            this.Controls.Add(this.txtCentre);
            this.Controls.Add(this.txtCompany);
            this.Name = "Form2";
            this.Text = "Cầu hình thông tin hệ thống";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TextBox txtCompany;
        private TextBox txtCentre;
        private TextBox txtLinkSever;
        private TextBox txtProvince;
        private Button btnOk;
        private Button btnCancel;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private ComboBox cbxHour1;
        private ComboBox cbxMinute1;
        private ComboBox cbxHour2;
        private ComboBox cbxMinute2;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private CheckBox ckbCountATforNight;
        private CheckBox ckbNightByStart;
    }
}