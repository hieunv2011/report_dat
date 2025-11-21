namespace DAT_ToolReports
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            button1 = new Button();
            textEmail = new TextBox();
            textPassword = new TextBox();
            btnLogin = new Button();
            dgvCoures = new DataGridView();
            ID = new DataGridViewTextBoxColumn();
            MaKH = new DataGridViewTextBoxColumn();
            TenKH = new DataGridViewTextBoxColumn();
            Hang = new DataGridViewTextBoxColumn();
            SoHV = new DataGridViewTextBoxColumn();
            NgayKG = new DataGridViewTextBoxColumn();
            NgayBG = new DataGridViewTextBoxColumn();
            NgáyH = new DataGridViewTextBoxColumn();
            chkbSortID = new CheckBox();
            dgwTrainees = new DataGridView();
            STT = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn1 = new DataGridViewTextBoxColumn();
            HoTen = new DataGridViewTextBoxColumn();
            NgaySinh = new DataGridViewTextBoxColumn();
            SoGioDB = new DataGridViewTextBoxColumn();
            SoKmDB = new DataGridViewTextBoxColumn();
            SoGio = new DataGridViewTextBoxColumn();
            SoKM = new DataGridViewTextBoxColumn();
            SoPhien = new DataGridViewTextBoxColumn();
            MaDK = new DataGridViewTextBoxColumn();
            Anh = new DataGridViewTextBoxColumn();
            MenuGetSessions = new ContextMenuStrip(components);
            xemCácPhiênHọcToolStripMenuItem = new ToolStripMenuItem();
            xuấtRaFileToolStripMenuItem = new ToolStripMenuItem();
            xuấtRaFileToolStripMenuItemPdf = new ToolStripMenuItem();
            inBáoCáoToolStripMenuItem = new ToolStripMenuItem();
            dgvSessions = new DataGridView();
            SessionID = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn4 = new DataGridViewTextBoxColumn();
            KhoiHanh = new DataGridViewTextBoxColumn();
            ThoiGian = new DataGridViewTextBoxColumn();
            QuangDuong = new DataGridViewTextBoxColumn();
            BienSo = new DataGridViewTextBoxColumn();
            SoAnh = new DataGridViewTextBoxColumn();
            DongBo = new DataGridViewTextBoxColumn();
            ViPham = new DataGridViewTextBoxColumn();
            label1 = new Label();
            chkbDongBo = new CheckBox();
            btnConfig = new Button();
            dgvVehicles = new DataGridView();
            dataGridViewTextBoxColumn2 = new DataGridViewTextBoxColumn();
            Plate = new DataGridViewTextBoxColumn();
            Model = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn3 = new DataGridViewTextBoxColumn();
            BranchName = new DataGridViewTextBoxColumn();
            dtpFrom = new DateTimePicker();
            dtpTo = new DateTimePicker();
            MenuGetSessionsVehicle = new ContextMenuStrip(components);
            xemCácPhiênHọcToolStripMenuItem1 = new ToolStripMenuItem();
            xuấtRaFileToolStripMenuItem1 = new ToolStripMenuItem();
            inBáoCáoToolStripMenuItem1 = new ToolStripMenuItem();
            label2 = new Label();
            label3 = new Label();
            MenuGetTrainees = new ContextMenuStrip(components);
            xemDanhSáchHọcViênToolStripMenuItem = new ToolStripMenuItem();
            xuấtDanhSáchRaFileToolStripMenuItem = new ToolStripMenuItem();
            xuấtDanhSáchRaFilePdfToolStripMenuItem = new ToolStripMenuItem();
            inDanhSáchToolStripMenuItem = new ToolStripMenuItem();
            xemDanhSáchPhiênHọcToolStripMenuItem = new ToolStripMenuItem();
            inDanhSáchPhiênHọcToolStripMenuItem = new ToolStripMenuItem();
            btnFindTrainee = new Button();
            btnFindVehicle = new Button();
            btnFindCouse = new Button();
            txtFind = new TextBox();
            btnOpenExcel = new Button();
            groupBox1 = new GroupBox();
            chkNonCheck = new CheckBox();
            chkCheckOk = new CheckBox();
            chkCheckNonOk = new CheckBox();
            txtLogs = new TextBox();
            btnOpenExcelCT = new Button();
            ((System.ComponentModel.ISupportInitialize)dgvCoures).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgwTrainees).BeginInit();
            MenuGetSessions.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvSessions).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvVehicles).BeginInit();
            MenuGetSessionsVehicle.SuspendLayout();
            MenuGetTrainees.SuspendLayout();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(1380, 791);
            button1.Name = "button1";
            button1.Size = new Size(61, 29);
            button1.TabIndex = 0;
            button1.Text = "button1";
            button1.UseVisualStyleBackColor = true;
            button1.Visible = false;
            button1.Click += button1_Click;
            // 
            // textEmail
            // 
            textEmail.Location = new Point(96, 3);
            textEmail.Name = "textEmail";
            textEmail.Size = new Size(200, 27);
            textEmail.TabIndex = 1;
            // 
            // textPassword
            // 
            textPassword.Location = new Point(302, 3);
            textPassword.Name = "textPassword";
            textPassword.PasswordChar = '*';
            textPassword.Size = new Size(200, 27);
            textPassword.TabIndex = 2;
            // 
            // btnLogin
            // 
            btnLogin.Location = new Point(508, 1);
            btnLogin.Name = "btnLogin";
            btnLogin.Size = new Size(67, 29);
            btnLogin.TabIndex = 3;
            btnLogin.Text = "Login";
            btnLogin.UseVisualStyleBackColor = true;
            btnLogin.Click += btnLogin_Click;
            // 
            // dgvCoures
            // 
            dgvCoures.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvCoures.Columns.AddRange(new DataGridViewColumn[] { ID, MaKH, TenKH, Hang, SoHV, NgayKG, NgayBG, NgáyH });
            dgvCoures.Location = new Point(12, 35);
            dgvCoures.Name = "dgvCoures";
            dgvCoures.RowHeadersVisible = false;
            dgvCoures.RowHeadersWidth = 51;
            dgvCoures.RowTemplate.Height = 29;
            dgvCoures.Size = new Size(876, 188);
            dgvCoures.TabIndex = 4;
            dgvCoures.CellDoubleClick += dgvCoures_CellDoubleClick;
            dgvCoures.MouseDown += dgvCoures_MouseDown;
            // 
            // ID
            // 
            ID.HeaderText = "ID";
            ID.MinimumWidth = 6;
            ID.Name = "ID";
            ID.Width = 70;
            // 
            // MaKH
            // 
            MaKH.HeaderText = "Mã Khóa học";
            MaKH.MinimumWidth = 6;
            MaKH.Name = "MaKH";
            MaKH.Width = 125;
            // 
            // TenKH
            // 
            TenKH.HeaderText = "Tên khóa";
            TenKH.MinimumWidth = 6;
            TenKH.Name = "TenKH";
            TenKH.Width = 125;
            // 
            // Hang
            // 
            Hang.HeaderText = "Hạng";
            Hang.MinimumWidth = 6;
            Hang.Name = "Hang";
            Hang.Width = 50;
            // 
            // SoHV
            // 
            SoHV.HeaderText = "Số HV";
            SoHV.MinimumWidth = 6;
            SoHV.Name = "SoHV";
            SoHV.Width = 125;
            // 
            // NgayKG
            // 
            NgayKG.HeaderText = "Ngày KG";
            NgayKG.MinimumWidth = 6;
            NgayKG.Name = "NgayKG";
            NgayKG.Width = 125;
            // 
            // NgayBG
            // 
            NgayBG.HeaderText = "Ngày BG";
            NgayBG.MinimumWidth = 6;
            NgayBG.Name = "NgayBG";
            NgayBG.Width = 125;
            // 
            // NgáyH
            // 
            NgáyH.HeaderText = "Ngày SH";
            NgáyH.MinimumWidth = 6;
            NgáyH.Name = "NgáyH";
            NgáyH.Width = 125;
            // 
            // chkbSortID
            // 
            chkbSortID.AutoSize = true;
            chkbSortID.Location = new Point(606, 4);
            chkbSortID.Name = "chkbSortID";
            chkbSortID.Size = new Size(163, 24);
            chkbSortID.TabIndex = 5;
            chkbSortID.Text = "Sắp xếp theo số giờ";
            chkbSortID.UseVisualStyleBackColor = true;
            // 
            // dgwTrainees
            // 
            dgwTrainees.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgwTrainees.Columns.AddRange(new DataGridViewColumn[] { STT, dataGridViewTextBoxColumn1, HoTen, NgaySinh, SoGioDB, SoKmDB, SoGio, SoKM, SoPhien, MaDK, Anh });
            dgwTrainees.Location = new Point(12, 229);
            dgwTrainees.Name = "dgwTrainees";
            dgwTrainees.RowHeadersVisible = false;
            dgwTrainees.RowHeadersWidth = 51;
            dgwTrainees.RowTemplate.Height = 29;
            dgwTrainees.Size = new Size(1356, 218);
            dgwTrainees.TabIndex = 6;
            dgwTrainees.MouseDown += dgwTrainees_MouseDown;
            // 
            // STT
            // 
            STT.HeaderText = "STT";
            STT.MinimumWidth = 6;
            STT.Name = "STT";
            STT.Width = 40;
            // 
            // dataGridViewTextBoxColumn1
            // 
            dataGridViewTextBoxColumn1.HeaderText = "ID";
            dataGridViewTextBoxColumn1.MinimumWidth = 6;
            dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            dataGridViewTextBoxColumn1.Width = 80;
            // 
            // HoTen
            // 
            HoTen.HeaderText = "Họ Tên";
            HoTen.MinimumWidth = 6;
            HoTen.Name = "HoTen";
            HoTen.Width = 200;
            // 
            // NgaySinh
            // 
            NgaySinh.HeaderText = "Ngày sinh";
            NgaySinh.MinimumWidth = 6;
            NgaySinh.Name = "NgaySinh";
            NgaySinh.Width = 125;
            // 
            // SoGioDB
            // 
            SoGioDB.HeaderText = "Số Giờ ĐB";
            SoGioDB.MinimumWidth = 6;
            SoGioDB.Name = "SoGioDB";
            SoGioDB.Width = 125;
            // 
            // SoKmDB
            // 
            SoKmDB.HeaderText = "Số Km ĐB";
            SoKmDB.MinimumWidth = 6;
            SoKmDB.Name = "SoKmDB";
            SoKmDB.Width = 125;
            // 
            // SoGio
            // 
            SoGio.HeaderText = "Số Giờ";
            SoGio.MinimumWidth = 6;
            SoGio.Name = "SoGio";
            SoGio.Width = 125;
            // 
            // SoKM
            // 
            SoKM.HeaderText = "Số KM";
            SoKM.MinimumWidth = 6;
            SoKM.Name = "SoKM";
            SoKM.Width = 125;
            // 
            // SoPhien
            // 
            SoPhien.HeaderText = "Số Phiên";
            SoPhien.MinimumWidth = 6;
            SoPhien.Name = "SoPhien";
            SoPhien.Width = 125;
            // 
            // MaDK
            // 
            MaDK.HeaderText = "Mã đăng ký";
            MaDK.MinimumWidth = 6;
            MaDK.Name = "MaDK";
            MaDK.Width = 125;
            // 
            // Anh
            // 
            Anh.HeaderText = "Anh";
            Anh.MinimumWidth = 6;
            Anh.Name = "Anh";
            Anh.Width = 125;
            // 
            // MenuGetSessions
            // 
            MenuGetSessions.ImageScalingSize = new Size(20, 20);
            MenuGetSessions.Items.AddRange(new ToolStripItem[] { xemCácPhiênHọcToolStripMenuItem, xuấtRaFileToolStripMenuItem, xuấtRaFileToolStripMenuItemPdf, inBáoCáoToolStripMenuItem });
            MenuGetSessions.Name = "MenuGetSessions";
            MenuGetSessions.Size = new Size(204, 76);
            // 
            // xemCácPhiênHọcToolStripMenuItem
            // 
            xemCácPhiênHọcToolStripMenuItem.Name = "xemCácPhiênHọcToolStripMenuItem";
            xemCácPhiênHọcToolStripMenuItem.Size = new Size(203, 24);
            xemCácPhiênHọcToolStripMenuItem.Text = "Xem các phiên học";
            xemCácPhiênHọcToolStripMenuItem.Click += xemCácPhiênHọcToolStripMenuItem_Click;
            // 
            // xuấtRaFileToolStripMenuItem
            // 
            xuấtRaFileToolStripMenuItem.Name = "xuấtRaFileToolStripMenuItem";
            xuấtRaFileToolStripMenuItem.Size = new Size(203, 24);
            xuấtRaFileToolStripMenuItem.Text = "Xuất ra file báo cáo HV";
            xuấtRaFileToolStripMenuItem.Click += xuấtRaFileToolStripMenuItem_Click;
            // xuấtRaFileToolStripMenuItemPdf
            // 
            xuấtRaFileToolStripMenuItemPdf.Name = "xuấtRaFileToolStripMenuItemPdf";
            xuấtRaFileToolStripMenuItemPdf.Size = new Size(203, 24);
            xuấtRaFileToolStripMenuItemPdf.Text = "Xuất ra file báo cáo HV PDF";
            xuấtRaFileToolStripMenuItemPdf.Click += xuấtRaFileToolStripMenuItemPdf_Click;
            // 
            // inBáoCáoToolStripMenuItem
            // 
            inBáoCáoToolStripMenuItem.Name = "inBáoCáoToolStripMenuItem";
            inBáoCáoToolStripMenuItem.Size = new Size(203, 24);
            inBáoCáoToolStripMenuItem.Text = "In báo cáo";
            inBáoCáoToolStripMenuItem.Click += inBáoCáoToolStripMenuItem_Click;
            // 
            // dgvSessions
            // 
            dgvSessions.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvSessions.Columns.AddRange(new DataGridViewColumn[] { SessionID, dataGridViewTextBoxColumn4, KhoiHanh, ThoiGian, QuangDuong, BienSo, SoAnh, DongBo, ViPham });
            dgvSessions.Location = new Point(11, 453);
            dgvSessions.Name = "dgvSessions";
            dgvSessions.RowHeadersVisible = false;
            dgvSessions.RowHeadersWidth = 51;
            dgvSessions.RowTemplate.Height = 29;
            dgvSessions.Size = new Size(1356, 358);
            dgvSessions.TabIndex = 8;
            // 
            // SessionID
            // 
            SessionID.HeaderText = "ID phiên học";
            SessionID.MinimumWidth = 6;
            SessionID.Name = "SessionID";
            SessionID.Width = 320;
            // 
            // dataGridViewTextBoxColumn4
            // 
            dataGridViewTextBoxColumn4.HeaderText = "Họ Tên HV";
            dataGridViewTextBoxColumn4.MinimumWidth = 6;
            dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            dataGridViewTextBoxColumn4.Width = 225;
            // 
            // KhoiHanh
            // 
            KhoiHanh.HeaderText = "Khởi hành";
            KhoiHanh.MinimumWidth = 6;
            KhoiHanh.Name = "KhoiHanh";
            KhoiHanh.Width = 180;
            // 
            // ThoiGian
            // 
            ThoiGian.HeaderText = "Thời gian";
            ThoiGian.MinimumWidth = 6;
            ThoiGian.Name = "ThoiGian";
            ThoiGian.Width = 125;
            // 
            // QuangDuong
            // 
            QuangDuong.HeaderText = "Quãng đường";
            QuangDuong.MinimumWidth = 6;
            QuangDuong.Name = "QuangDuong";
            QuangDuong.Width = 145;
            // 
            // BienSo
            // 
            BienSo.HeaderText = "Biển số";
            BienSo.MinimumWidth = 6;
            BienSo.Name = "BienSo";
            BienSo.Width = 125;
            // 
            // SoAnh
            // 
            SoAnh.HeaderText = "Số ảnh";
            SoAnh.MinimumWidth = 6;
            SoAnh.Name = "SoAnh";
            SoAnh.Width = 125;
            // 
            // DongBo
            // 
            DongBo.HeaderText = "Đồng bộ";
            DongBo.MinimumWidth = 6;
            DongBo.Name = "DongBo";
            DongBo.Width = 125;
            // 
            // ViPham
            // 
            ViPham.HeaderText = "Vi phạm";
            ViPham.MinimumWidth = 6;
            ViPham.Name = "ViPham";
            ViPham.Width = 525;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(19, 6);
            label1.Name = "label1";
            label1.Size = new Size(71, 20);
            label1.TabIndex = 9;
            label1.Text = "Tài khoản";
            // 
            // chkbDongBo
            // 
            chkbDongBo.AutoSize = true;
            chkbDongBo.Location = new Point(777, 6);
            chkbDongBo.Name = "chkbDongBo";
            chkbDongBo.Size = new Size(177, 24);
            chkbDongBo.TabIndex = 11;
            chkbDongBo.Text = "Chỉ lấy phiên đồng bộ";
            chkbDongBo.UseVisualStyleBackColor = true;
            // 
            // btnConfig
            // 
            btnConfig.Location = new Point(1372, 2);
            btnConfig.Name = "btnConfig";
            btnConfig.Size = new Size(137, 29);
            btnConfig.TabIndex = 12;
            btnConfig.Text = "Cấu hình";
            btnConfig.UseCompatibleTextRendering = true;
            btnConfig.UseVisualStyleBackColor = true;
            btnConfig.Click += btnConfig_Click;
            // 
            // dgvVehicles
            // 
            dgvVehicles.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvVehicles.Columns.AddRange(new DataGridViewColumn[] { dataGridViewTextBoxColumn2, Plate, Model, dataGridViewTextBoxColumn3, BranchName });
            dgvVehicles.Location = new Point(894, 35);
            dgvVehicles.Name = "dgvVehicles";
            dgvVehicles.RowHeadersVisible = false;
            dgvVehicles.RowHeadersWidth = 51;
            dgvVehicles.RowTemplate.Height = 29;
            dgvVehicles.Size = new Size(474, 188);
            dgvVehicles.TabIndex = 13;
            dgvVehicles.MouseDown += dgvVehicles_MouseDown;
            // 
            // dataGridViewTextBoxColumn2
            // 
            dataGridViewTextBoxColumn2.HeaderText = "Stt";
            dataGridViewTextBoxColumn2.MinimumWidth = 6;
            dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            dataGridViewTextBoxColumn2.Width = 60;
            // 
            // Plate
            // 
            Plate.HeaderText = "Biến số";
            Plate.MinimumWidth = 6;
            Plate.Name = "Plate";
            Plate.Width = 125;
            // 
            // Model
            // 
            Model.HeaderText = "Model";
            Model.MinimumWidth = 6;
            Model.Name = "Model";
            Model.Width = 125;
            // 
            // dataGridViewTextBoxColumn3
            // 
            dataGridViewTextBoxColumn3.HeaderText = "Hạng";
            dataGridViewTextBoxColumn3.MinimumWidth = 6;
            dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            dataGridViewTextBoxColumn3.Width = 80;
            // 
            // BranchName
            // 
            BranchName.HeaderText = "CSDT";
            BranchName.MinimumWidth = 6;
            BranchName.Name = "BranchName";
            BranchName.Width = 125;
            // 
            // dtpFrom
            // 
            dtpFrom.Format = DateTimePickerFormat.Custom;
            dtpFrom.Location = new Point(1061, 6);
            dtpFrom.Name = "dtpFrom";
            dtpFrom.Size = new Size(103, 27);
            dtpFrom.TabIndex = 14;
            // 
            // dtpTo
            // 
            dtpTo.Format = DateTimePickerFormat.Custom;
            dtpTo.Location = new Point(1264, 5);
            dtpTo.Name = "dtpTo";
            dtpTo.Size = new Size(103, 27);
            dtpTo.TabIndex = 15;
            // 
            // MenuGetSessionsVehicle
            // 
            MenuGetSessionsVehicle.ImageScalingSize = new Size(20, 20);
            MenuGetSessionsVehicle.Items.AddRange(new ToolStripItem[] { xemCácPhiênHọcToolStripMenuItem1, xuấtRaFileToolStripMenuItem1, inBáoCáoToolStripMenuItem1 });
            MenuGetSessionsVehicle.Name = "MenuGetSessionsVehicle";
            MenuGetSessionsVehicle.Size = new Size(204, 76);
            // 
            // xemCácPhiênHọcToolStripMenuItem1
            // 
            xemCácPhiênHọcToolStripMenuItem1.Name = "xemCácPhiênHọcToolStripMenuItem1";
            xemCácPhiênHọcToolStripMenuItem1.Size = new Size(203, 24);
            xemCácPhiênHọcToolStripMenuItem1.Text = "Xem các phiên học";
            xemCácPhiênHọcToolStripMenuItem1.Click += xemCácPhiênHọcToolStripMenuItem1_Click;
            // 
            // xuấtRaFileToolStripMenuItem1
            // 
            xuấtRaFileToolStripMenuItem1.Name = "xuấtRaFileToolStripMenuItem1";
            xuấtRaFileToolStripMenuItem1.Size = new Size(203, 24);
            xuấtRaFileToolStripMenuItem1.Text = "Xuất ra file";
            xuấtRaFileToolStripMenuItem1.Click += xuấtRaFileToolStripMenuItem1_Click;
            // 
            // inBáoCáoToolStripMenuItem1
            // 
            inBáoCáoToolStripMenuItem1.Name = "inBáoCáoToolStripMenuItem1";
            inBáoCáoToolStripMenuItem1.Size = new Size(203, 24);
            inBáoCáoToolStripMenuItem1.Text = "In báo cáo";
            inBáoCáoToolStripMenuItem1.Click += inBáoCáoToolStripMenuItem1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(990, 10);
            label2.Name = "label2";
            label2.Size = new Size(65, 20);
            label2.TabIndex = 16;
            label2.Text = "Từ ngày:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(1170, 10);
            label3.Name = "label3";
            label3.Size = new Size(75, 20);
            label3.TabIndex = 17;
            label3.Text = "Đến ngày:";
            // 
            // MenuGetTrainees
            // 
            MenuGetTrainees.ImageScalingSize = new Size(20, 20);
            MenuGetTrainees.Items.AddRange(new ToolStripItem[] { xemDanhSáchHọcViênToolStripMenuItem, xuấtDanhSáchRaFileToolStripMenuItem, xuấtDanhSáchRaFilePdfToolStripMenuItem, inDanhSáchToolStripMenuItem, xemDanhSáchPhiênHọcToolStripMenuItem, inDanhSáchPhiênHọcToolStripMenuItem });
            MenuGetTrainees.Name = "MenuGetTrainees";
            MenuGetTrainees.Size = new Size(248, 124);
            // 
            // xemDanhSáchHọcViênToolStripMenuItem
            // 
            xemDanhSáchHọcViênToolStripMenuItem.Name = "xemDanhSáchHọcViênToolStripMenuItem";
            xemDanhSáchHọcViênToolStripMenuItem.Size = new Size(247, 24);
            xemDanhSáchHọcViênToolStripMenuItem.Text = "Xem danh sách học viên";
            xemDanhSáchHọcViênToolStripMenuItem.Click += xemDanhSáchHọcViênToolStripMenuItem_Click;
            // 
            // xuấtDanhSáchRaFileToolStripMenuItem
            // 
            xuấtDanhSáchRaFileToolStripMenuItem.Name = "xuấtDanhSáchRaFileToolStripMenuItem";
            xuấtDanhSáchRaFileToolStripMenuItem.Size = new Size(247, 24);
            xuấtDanhSáchRaFileToolStripMenuItem.Text = "Xuất danh sách HV ra file Excel";
            xuấtDanhSáchRaFileToolStripMenuItem.Click += xuấtDanhSáchRaFileToolStripMenuItem_Click;
            // 
            // xuấtDanhSáchRaFilePdfToolStripMenuItem
            // 
            xuấtDanhSáchRaFilePdfToolStripMenuItem.Name = "xuấtDanhSáchRaFilePdfToolStripMenuItem";
            xuấtDanhSáchRaFilePdfToolStripMenuItem.Size = new Size(247, 24);
            xuấtDanhSáchRaFilePdfToolStripMenuItem.Text = "Xuất danh sách HV ra file PDF";
            xuấtDanhSáchRaFilePdfToolStripMenuItem.Click += xuấtDanhSáchRaFilePdfToolStripMenuItem_Click;
            // 
            // inDanhSáchToolStripMenuItem
            // 
            inDanhSáchToolStripMenuItem.Name = "inDanhSáchToolStripMenuItem";
            inDanhSáchToolStripMenuItem.Size = new Size(247, 24);
            inDanhSáchToolStripMenuItem.Text = "In danh sách HV";
            inDanhSáchToolStripMenuItem.Click += inDanhSáchToolStripMenuItem_Click;
            // 
            // xemDanhSáchPhiênHọcToolStripMenuItem
            // 
            xemDanhSáchPhiênHọcToolStripMenuItem.Name = "xemDanhSáchPhiênHọcToolStripMenuItem";
            xemDanhSáchPhiênHọcToolStripMenuItem.Size = new Size(247, 24);
            xemDanhSáchPhiênHọcToolStripMenuItem.Text = "Xem danh sách phiên học";
            xemDanhSáchPhiênHọcToolStripMenuItem.Click += xemDanhSáchPhiênHọcToolStripMenuItem_Click;
            // 
            // inDanhSáchPhiênHọcToolStripMenuItem
            // 
            inDanhSáchPhiênHọcToolStripMenuItem.Name = "inDanhSáchPhiênHọcToolStripMenuItem";
            inDanhSáchPhiênHọcToolStripMenuItem.Size = new Size(247, 24);
            inDanhSáchPhiênHọcToolStripMenuItem.Text = "Xuất danh sách phiên học";
            inDanhSáchPhiênHọcToolStripMenuItem.Click += inDanhSáchPhiênHọcToolStripMenuItem_Click;
            // 
            // btnFindTrainee
            // 
            btnFindTrainee.Location = new Point(1374, 138);
            btnFindTrainee.Name = "btnFindTrainee";
            btnFindTrainee.Size = new Size(137, 29);
            btnFindTrainee.TabIndex = 18;
            btnFindTrainee.Text = "Tìm học viên";
            btnFindTrainee.UseVisualStyleBackColor = true;
            btnFindTrainee.Click += btnFindTrainee_Click;
            // 
            // btnFindVehicle
            // 
            btnFindVehicle.Location = new Point(1372, 103);
            btnFindVehicle.Name = "btnFindVehicle";
            btnFindVehicle.Size = new Size(137, 29);
            btnFindVehicle.TabIndex = 19;
            btnFindVehicle.Text = "Tìm xe tập lái";
            btnFindVehicle.UseVisualStyleBackColor = true;
            btnFindVehicle.Click += btnFindVehicle_Click;
            // 
            // btnFindCouse
            // 
            btnFindCouse.Location = new Point(1372, 68);
            btnFindCouse.Name = "btnFindCouse";
            btnFindCouse.Size = new Size(137, 29);
            btnFindCouse.TabIndex = 20;
            btnFindCouse.Text = "Tìm khóa học";
            btnFindCouse.UseVisualStyleBackColor = true;
            btnFindCouse.Click += btnFindCouse_Click;
            // 
            // txtFind
            // 
            txtFind.Location = new Point(1372, 35);
            txtFind.Name = "txtFind";
            txtFind.Size = new Size(137, 27);
            txtFind.TabIndex = 21;
            // 
            // btnOpenExcel
            // 
            btnOpenExcel.Location = new Point(1374, 173);
            btnOpenExcel.Name = "btnOpenExcel";
            btnOpenExcel.Size = new Size(135, 29);
            btnOpenExcel.TabIndex = 22;
            btnOpenExcel.Text = "Open Excel TH";
            btnOpenExcel.UseVisualStyleBackColor = true;
            btnOpenExcel.Click += btnOpenExcel_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(chkNonCheck);
            groupBox1.Controls.Add(chkCheckOk);
            groupBox1.Controls.Add(chkCheckNonOk);
            groupBox1.Location = new Point(1374, 237);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(135, 125);
            groupBox1.TabIndex = 23;
            groupBox1.TabStop = false;
            groupBox1.Text = "Lấy các phiên";
            // 
            // chkNonCheck
            // 
            chkNonCheck.AutoSize = true;
            chkNonCheck.Location = new Point(6, 84);
            chkNonCheck.Name = "chkNonCheck";
            chkNonCheck.Size = new Size(123, 24);
            chkNonCheck.TabIndex = 2;
            chkNonCheck.Text = "Chưa kiểm tra";
            chkNonCheck.UseVisualStyleBackColor = true;
            // 
            // chkCheckOk
            // 
            chkCheckOk.AutoSize = true;
            chkCheckOk.Location = new Point(6, 54);
            chkCheckOk.Name = "chkCheckOk";
            chkCheckOk.Size = new Size(131, 24);
            chkCheckOk.TabIndex = 1;
            chkCheckOk.Text = "Không vi phạm";
            chkCheckOk.UseVisualStyleBackColor = true;
            // 
            // chkCheckNonOk
            // 
            chkCheckNonOk.AutoSize = true;
            chkCheckNonOk.Location = new Point(5, 24);
            chkCheckNonOk.Name = "chkCheckNonOk";
            chkCheckNonOk.Size = new Size(86, 24);
            chkCheckNonOk.TabIndex = 0;
            chkCheckNonOk.Text = "Vi phạm";
            chkCheckNonOk.UseVisualStyleBackColor = true;
            // 
            // txtLogs
            // 
            txtLogs.Location = new Point(1374, 368);
            txtLogs.Multiline = true;
            txtLogs.Name = "txtLogs";
            txtLogs.Size = new Size(135, 417);
            txtLogs.TabIndex = 24;
            // 
            // btnOpenExcelCT
            // 
            btnOpenExcelCT.Location = new Point(1374, 208);
            btnOpenExcelCT.Name = "btnOpenExcelCT";
            btnOpenExcelCT.Size = new Size(135, 29);
            btnOpenExcelCT.TabIndex = 25;
            btnOpenExcelCT.Text = "Open Excel CT";
            btnOpenExcelCT.UseVisualStyleBackColor = true;
            btnOpenExcelCT.Click += btnOpenExcelCT_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1510, 821);
            Controls.Add(btnOpenExcelCT);
            Controls.Add(txtLogs);  
            Controls.Add(groupBox1);
            Controls.Add(btnOpenExcel);
            Controls.Add(txtFind);
            Controls.Add(btnFindCouse);
            Controls.Add(btnFindVehicle);
            Controls.Add(btnFindTrainee);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(dtpTo);
            Controls.Add(dtpFrom);
            Controls.Add(dgvVehicles);
            Controls.Add(btnConfig);
            Controls.Add(chkbDongBo);
            Controls.Add(label1);
            Controls.Add(dgvSessions);
            Controls.Add(dgwTrainees);
            Controls.Add(chkbSortID);
            Controls.Add(dgvCoures);
            Controls.Add(btnLogin);
            Controls.Add(textPassword);
            Controls.Add(textEmail);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Công cụ xuất báo cáo hệ thống giám sát học thực hành";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dgvCoures).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgwTrainees).EndInit();
            MenuGetSessions.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvSessions).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvVehicles).EndInit();
            MenuGetSessionsVehicle.ResumeLayout(false);
            MenuGetTrainees.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private TextBox textEmail;
        private TextBox textPassword;
        private Button btnLogin;
        private DataGridView dgvCoures;
        private CheckBox chkbSortID;
        private DataGridView dgwTrainees;
        private ContextMenuStrip MenuGetSessions;
        private ToolStripMenuItem xemCácPhiênHọcToolStripMenuItem;
        private ToolStripMenuItem xuấtRaFileToolStripMenuItem;
        private ToolStripMenuItem xuấtRaFileToolStripMenuItemPdf;
        private ToolStripMenuItem inBáoCáoToolStripMenuItem;
        private DataGridView dgvSessions;
        private Label label1;
        private CheckBox chkbDongBo;
        private Button btnConfig;
        private DataGridView dgvVehicles;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private DataGridViewTextBoxColumn Plate;
        private DataGridViewTextBoxColumn Model;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private DataGridViewTextBoxColumn BranchName;
        private DateTimePicker dtpFrom;
        private DateTimePicker dtpTo;
        private ContextMenuStrip MenuGetSessionsVehicle;
        private ToolStripMenuItem xemCácPhiênHọcToolStripMenuItem1;
        private ToolStripMenuItem xuấtRaFileToolStripMenuItem1;
        private ToolStripMenuItem inBáoCáoToolStripMenuItem1;
        private DataGridViewTextBoxColumn STT;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private DataGridViewTextBoxColumn HoTen;
        private DataGridViewTextBoxColumn NgaySinh;
        private DataGridViewTextBoxColumn SoGioDB;
        private DataGridViewTextBoxColumn SoKmDB;
        private DataGridViewTextBoxColumn SoGio;
        private DataGridViewTextBoxColumn SoKM;
        private DataGridViewTextBoxColumn SoPhien;
        private DataGridViewTextBoxColumn MaDK;
        private DataGridViewTextBoxColumn Anh;
        private Label label2;
        private Label label3;
        private ContextMenuStrip MenuGetTrainees;
        private ToolStripMenuItem xemDanhSáchHọcViênToolStripMenuItem;
        private ToolStripMenuItem xuấtDanhSáchRaFileToolStripMenuItem;
        private ToolStripMenuItem xuấtDanhSáchRaFilePdfToolStripMenuItem;
        private ToolStripMenuItem inDanhSáchToolStripMenuItem;
        private Button btnFindTrainee;
        private Button btnFindVehicle;
        private Button btnFindCouse;
        private DataGridViewTextBoxColumn ID;
        private DataGridViewTextBoxColumn MaKH;
        private DataGridViewTextBoxColumn TenKH;
        private DataGridViewTextBoxColumn Hang;
        private DataGridViewTextBoxColumn SoHV;
        private DataGridViewTextBoxColumn NgayKG;
        private DataGridViewTextBoxColumn NgayBG;
        private DataGridViewTextBoxColumn NgáyH;
        private TextBox txtFind;
        private Button btnOpenExcel;
        private DataGridViewTextBoxColumn SessionID;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private DataGridViewTextBoxColumn KhoiHanh;
        private DataGridViewTextBoxColumn ThoiGian;
        private DataGridViewTextBoxColumn QuangDuong;
        private DataGridViewTextBoxColumn BienSo;
        private DataGridViewTextBoxColumn SoAnh;
        private DataGridViewTextBoxColumn DongBo;
        private DataGridViewTextBoxColumn ViPham;
        private GroupBox groupBox1;
        private CheckBox chkNonCheck;
        private CheckBox chkCheckOk;
        private CheckBox chkCheckNonOk;
        private ToolStripMenuItem xemDanhSáchPhiênHọcToolStripMenuItem;
        private ToolStripMenuItem inDanhSáchPhiênHọcToolStripMenuItem;
        private TextBox txtLogs;
        private Button btnOpenExcelCT;

    }
}