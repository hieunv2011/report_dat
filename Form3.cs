using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Microsoft.Office.Interop.Excel;

namespace DAT_ToolReports
{
    public partial class Form3 : Form
    {
        public class SessionExcel
        {
            public int STT;
            public string MaPhienHoc;
            public string TrangThai;
            public DateTime ThoiGianTruyen;
            public DateTime ThoiGianPhienHoc;
            public string MaHocVien;
            public string TenHocVien;
            public string MaKhoaHoc;
            public string LoaiKhoaHoc;
            public string DonViDaoTao;
            public string SoGTVT;
            public string DonViTruyenDL;
            public double ThoiGianDaoTao;
            public double QuangDuongDaoTao;
            public string BienSoXe;
            public string ViPhamQuyDinh;
            public DateTime ThoiGianPhienHoc_KT;
            public string TrungHV;
            public string TrungXe;
        }
        public List<SessionExcel> SessionExcels;
        public Form3()
        {
            InitializeComponent();
            SessionExcels = new List<SessionExcel>();
            ckbTrungHV.Checked = true;
            ckbTrungXe.Checked = true;
            txtMaxDuration.Text = "720";
            txtMinDuration.Text = "3";
            txtMinDupMinute.Text = "15";
            ckbMax.Checked = true;
            ckbMin.Checked = true;
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                SessionExcels.Clear();
                CultureInfo culture;
                DateTimeStyles styles;

                Microsoft.Office.Interop.Excel.Application excelApplication;
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.Visible = false;
                //string fileName = "C:\\sampleExcelFile.xlsx";

                //open the workbook
                Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)excelApplication.Workbooks.Open(openFileDialog1.FileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);

                //access the cells
                string format = "dd-MM-yyyy HH:mm:ss";
                for (int row = 5; row <= worksheet.UsedRange.Rows.Count - 2; ++row)
                {
                    if((row%1000) == 0)
                        txtLogs.AppendText(row.ToString() + "/" + (worksheet.UsedRange.Rows.Count - 5).ToString() + "--");
                    SessionExcel items = new SessionExcel();
                    if (valueArray[row, 1] is null)
                        break;
                    items.STT = Int32.Parse(valueArray[row, 1].ToString());
                    if (items.STT != (row - 4)) 
                        break;
                    items.MaPhienHoc = valueArray[row, 2].ToString();
                    items.TrangThai = valueArray[row, 3].ToString();
                    items.ThoiGianTruyen = DateTime.ParseExact(valueArray[row, 4].ToString(), format, CultureInfo.InvariantCulture) ;
                    items.ThoiGianPhienHoc = DateTime.ParseExact(valueArray[row, 5].ToString(), format, CultureInfo.InvariantCulture);
                    items.MaHocVien = valueArray[row, 6].ToString();
                    items.TenHocVien = valueArray[row, 7].ToString();
                    items.MaKhoaHoc = valueArray[row, 8].ToString();
                    items.LoaiKhoaHoc = valueArray[row, 9].ToString();
                    items.DonViDaoTao = valueArray[row, 10].ToString();
                    items.SoGTVT = valueArray[row, 11].ToString();
                    items.DonViTruyenDL = valueArray[row, 12].ToString();
                    items.ThoiGianDaoTao = double.Parse(valueArray[row, 13].ToString(), CultureInfo.InvariantCulture);
                    items.QuangDuongDaoTao = double.Parse(valueArray[row, 14].ToString(), CultureInfo.InvariantCulture);
                    items.BienSoXe = valueArray[row, 15].ToString();
                    items.ViPhamQuyDinh = valueArray[row, 16].ToString();
                    items.ThoiGianPhienHoc_KT = items.ThoiGianPhienHoc.AddSeconds(items.ThoiGianDaoTao * 3600);
                    items.TrungHV = "";
                    items.TrungXe = "";

                    SessionExcels.Add(items);
                    //for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    //{
                    //    //access each cell
                    //    MessageBox.Show(valueArray[row, col].ToString());
                    //}
                }

                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);

                excelApplication.Quit();
                Marshal.FinalReleaseComObject(excelApplication);
                SessionExcels = SessionExcels.OrderBy(item => item.MaHocVien).ThenBy(item => item.ThoiGianPhienHoc).ToList();
                int Count = 0;
                foreach (SessionExcel session in SessionExcels)
                {
                    Count++;
                    dgvSessionsExcels.Rows.Add(Count.ToString(), session.STT, session.MaHocVien, session.TenHocVien, session.MaPhienHoc,
                        session.ThoiGianTruyen.ToString(), session.ThoiGianPhienHoc.ToString(), session.ThoiGianPhienHoc_KT.ToString()
                        ,session.ThoiGianDaoTao.ToString("0.00"), session.QuangDuongDaoTao.ToString("0.00"), session.BienSoXe, session.TrungHV, 
                        session.TrungXe, session.ViPhamQuyDinh);
                }
                Cursor.Current = Cursors.Default;
            }
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            DateTime ChkStart, ChkStop;
            string ChkMaHV, sTrungHocVien;
            string ChkBienSoXe, sTrungXeTL;
            int ChkSTT, countSoPhienTrungHV, countSoPhienTrungXe;
            int MaxDuration = 0;
            int MinDuration = 0;
            int MinDupMinute = 0;

            Cursor.Current = Cursors.WaitCursor;

            MinDupMinute = Int32.Parse(txtMinDupMinute.Text);
            if (ckbMax.Checked)
            {
                MaxDuration = Int32.Parse(txtMaxDuration.Text);
            }
            if (ckbMin.Checked)
            {
                MinDuration = Int32.Parse(txtMinDuration.Text);
            }

            if ((!ckbTrungHV.Checked)&&(!ckbTrungXe.Checked))
            {
                MessageBox.Show("Phải chọn kiểm tra trùng học viên học trùng xe");
                return;
            }
            else if((ckbTrungHV.Checked) && (!ckbTrungXe.Checked))
            {
                MessageBox.Show("Check trùng học viên giống công thức vị Thanh");
                SessionExcels = SessionExcels.OrderBy(item => item.MaHocVien).ThenBy(item => item.ThoiGianPhienHoc).ToList();
                for (int i = 0; i < SessionExcels.Count - 1; i++)
                {
                    SessionExcels[i].TrungXe = "";
                    SessionExcels[i].TrungHV = "";
                    if ((SessionExcels[i].MaHocVien == SessionExcels[i+1].MaHocVien)&& (SessionExcels[i].BienSoXe != SessionExcels[i + 1].BienSoXe))
                    {
                        if ((SessionExcels[i + 1].ThoiGianPhienHoc >= SessionExcels[i].ThoiGianPhienHoc) &&
                            (SessionExcels[i + 1].ThoiGianPhienHoc <= SessionExcels[i].ThoiGianPhienHoc_KT) &&
                            (SessionExcels[i].ThoiGianPhienHoc_KT > SessionExcels[i + 1].ThoiGianPhienHoc.AddMinutes(MinDupMinute)))
                        {
                            SessionExcels[i].TrungHV = SessionExcels[i + 1].STT.ToString();
                        }
                    }
                }
            }
            else if((!ckbTrungHV.Checked) && (ckbTrungXe.Checked))
            {
                MessageBox.Show("Check trùng xe giống công thức vị Thanh");
                SessionExcels = SessionExcels.OrderBy(item => item.BienSoXe).ThenBy(item => item.ThoiGianPhienHoc).ToList();
                for (int i = 0; i < SessionExcels.Count - 1; i++)
                {
                    SessionExcels[i].TrungXe = "";
                    SessionExcels[i].TrungHV = "";
                    if (SessionExcels[i].BienSoXe == SessionExcels[i + 1].BienSoXe)
                    {
                        if ((SessionExcels[i + 1].ThoiGianPhienHoc >= SessionExcels[i].ThoiGianPhienHoc) &&
                            (SessionExcels[i + 1].ThoiGianPhienHoc <= SessionExcels[i].ThoiGianPhienHoc_KT) &&
                            (SessionExcels[i].ThoiGianPhienHoc_KT >= SessionExcels[i + 1].ThoiGianPhienHoc.AddMinutes(MinDupMinute)))
                        {
                            SessionExcels[i].TrungXe = SessionExcels[i + 1].STT.ToString();
                        }
                    }
                }
            }
            else
            {
                sTrungHocVien = "";
                sTrungXeTL = "";
                SessionExcels = SessionExcels.OrderBy(item => item.MaHocVien).ThenBy(item => item.ThoiGianPhienHoc).ToList();
                for (int i = 0; i < SessionExcels.Count; i++)
                {
                    if ((i % 1000) == 0)
                        txtLogs.AppendText("<" + i.ToString() + "/" + SessionExcels.Count.ToString() + ">--");
                    if (ckbMax.Checked)
                    {
                        if ((SessionExcels[i].ThoiGianDaoTao * 60) >= MaxDuration) continue;
                    }
                    if (ckbMin.Checked)
                    {
                        if ((SessionExcels[i].ThoiGianDaoTao * 60) <= MinDuration) continue;
                    }
                    ChkStart = SessionExcels[i].ThoiGianPhienHoc;
                    ChkStop = SessionExcels[i].ThoiGianPhienHoc_KT;
                    ChkMaHV = SessionExcels[i].MaHocVien;
                    ChkBienSoXe = SessionExcels[i].BienSoXe;
                    ChkSTT = SessionExcels[i].STT;
                    sTrungHocVien = "";
                    sTrungXeTL = "";
                    SessionExcels[i].TrungXe = "";
                    SessionExcels[i].TrungHV = "";
                    foreach (SessionExcel sessionChk in SessionExcels)
                    {
                        if ((sessionChk.MaHocVien != ChkMaHV)&& (sessionChk.BienSoXe != ChkBienSoXe)) continue;
                        if (ckbMax.Checked)
                        {
                            if ((sessionChk.ThoiGianDaoTao * 60) >= MaxDuration) continue;
                        }
                        if (ckbMin.Checked)
                        {
                            if ((sessionChk.ThoiGianDaoTao * 60) <= MinDuration) continue;
                        }
                        if (sessionChk.ThoiGianPhienHoc >= ChkStop) continue;
                        if (sessionChk.ThoiGianPhienHoc_KT <= ChkStart) continue;
                        if (sessionChk.STT == ChkSTT) continue;
                        if (ckbTrungHV.Checked)
                        {
                            if (sessionChk.MaHocVien == ChkMaHV)
                            {
                                if (((sessionChk.ThoiGianPhienHoc > ChkStart) && (sessionChk.ThoiGianPhienHoc.AddMinutes(MinDupMinute) < ChkStop)) ||
                                    ((sessionChk.ThoiGianPhienHoc_KT > ChkStart.AddMinutes(MinDupMinute)) && (sessionChk.ThoiGianPhienHoc_KT < ChkStop)) ||
                                    ((ChkStart > sessionChk.ThoiGianPhienHoc) && (ChkStart.AddMinutes(MinDupMinute) < sessionChk.ThoiGianPhienHoc_KT)) ||
                                    ((ChkStop > sessionChk.ThoiGianPhienHoc.AddMinutes(MinDupMinute)) && (ChkStop < sessionChk.ThoiGianPhienHoc_KT)))
                                {
                                    sTrungHocVien = sTrungHocVien + "," + sessionChk.STT.ToString();
                                }
                            }
                        }
                        if (ckbTrungXe.Checked)
                        {
                            if (sessionChk.BienSoXe == ChkBienSoXe)
                            {
                                if (((sessionChk.ThoiGianPhienHoc > ChkStart) && (sessionChk.ThoiGianPhienHoc.AddMinutes(MinDupMinute) < ChkStop)) ||
                                    ((sessionChk.ThoiGianPhienHoc_KT > ChkStart.AddMinutes(MinDupMinute)) && (sessionChk.ThoiGianPhienHoc_KT < ChkStop)) ||
                                    ((ChkStart > sessionChk.ThoiGianPhienHoc) && (ChkStart.AddMinutes(MinDupMinute) < sessionChk.ThoiGianPhienHoc_KT)) ||
                                    ((ChkStop > sessionChk.ThoiGianPhienHoc.AddMinutes(MinDupMinute)) && (ChkStop < sessionChk.ThoiGianPhienHoc_KT)))
                                {
                                    sTrungXeTL = sTrungXeTL + "," + sessionChk.STT.ToString();
                                }
                            }
                        }

                    }
                    SessionExcels[i].TrungHV = sTrungHocVien;
                    SessionExcels[i].TrungXe = sTrungXeTL;
                }

            }
            
            countSoPhienTrungHV = 0;
            countSoPhienTrungXe = 0;
            dgvSessionsExcels.Rows.Clear();
            int Count = 0;
            foreach (SessionExcel session in SessionExcels)
            {
                if (session.TrungHV.Length > 1) countSoPhienTrungHV = countSoPhienTrungHV + 1;
                if (session.TrungXe.Length > 1) countSoPhienTrungXe = countSoPhienTrungXe + 1;
                Count++;
                dgvSessionsExcels.Rows.Add(Count.ToString(), session.STT, session.MaHocVien, session.TenHocVien, session.MaPhienHoc, 
                    session.ThoiGianTruyen.ToString(), session.ThoiGianPhienHoc.ToString(), session.ThoiGianPhienHoc_KT.ToString(),
                    session.ThoiGianDaoTao.ToString("0.00"), session.QuangDuongDaoTao.ToString("0.00"), session.BienSoXe, session.TrungHV, 
                    session.TrungXe, session.ViPhamQuyDinh);
            }
            Cursor.Current = Cursors.Default;
            txtLogs.AppendText("So phien trung HV: " + countSoPhienTrungHV.ToString() + "---So phien trung Xe: " + countSoPhienTrungXe.ToString());
            MessageBox.Show("So phien trung HV: " + countSoPhienTrungHV.ToString() + "---So phien trung Xe: " + countSoPhienTrungXe.ToString());
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            dgvSessionsExcels.Rows.Clear();
            int Count = 0;
            foreach (SessionExcel session in SessionExcels)
            {
                if ((session.TrungHV.Length > 1)|| (session.TrungXe.Length > 1))
                { 
                    Count++;
                    dgvSessionsExcels.Rows.Add(Count.ToString(), session.STT, session.MaHocVien, session.TenHocVien, session.MaPhienHoc,
                        session.ThoiGianTruyen.ToString(), session.ThoiGianPhienHoc.ToString(), session.ThoiGianPhienHoc_KT.ToString()
                        , session.ThoiGianDaoTao.ToString("0.00"), session.QuangDuongDaoTao.ToString("0.00"), session.BienSoXe, session.TrungHV,
                        session.TrungXe, session.ViPhamQuyDinh);
                }
            }
        }

        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                object misValue = System.Reflection.Missing.Value;

                worksheet.PageSetup.Orientation = PageOrientationType.Portrait;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0;
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.TopMargin = 1;
                worksheet.PageSetup.BottomMargin = 1;
                worksheet.PageSetup.LeftMargin = 1;
                worksheet.PageSetup.RightMargin = 0.3;

                Aspose.Cells.Cells Cells = worksheet.Cells;

                Aspose.Cells.Style style1;

                style1 = Cells["A1"].GetStyle();
                style1.VerticalAlignment = TextAlignmentType.Center;
                style1.HorizontalAlignment = TextAlignmentType.Center;
                style1.Font.Color = System.Drawing.Color.Black;
                style1.Font.IsBold = true;
                style1.Font.IsItalic = false;
                style1.Font.Size = 13;
                style1.Font.Name = "Times New Roman";


                Aspose.Cells.Style styleDate;
                styleDate = Cells["A1"].GetStyle();
                styleDate.VerticalAlignment = TextAlignmentType.Center;
                styleDate.HorizontalAlignment = TextAlignmentType.Center;
                styleDate.Font.Color = System.Drawing.Color.Black;
                styleDate.Font.IsBold = true;
                styleDate.Font.IsItalic = true;
                styleDate.Font.Size = 13;
                styleDate.Font.Name = "Times New Roman";
               

                

                Aspose.Cells.Style style5;
                style5 = Cells["A1"].GetStyle();
                style5.Font.Size = 13;
                style5.Font.Name = "Times New Roman";
                style5.Font.IsBold = false;
                //style5.ShrinkToFit = true;
                style5.HorizontalAlignment = TextAlignmentType.Left;

                

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(1, 0, 1, 13);
                Cells["A2"].Value = "BÁO CÁO CÁC PHIÊN HỌC BỊ TRÙNG HỌC VIÊN VÀ XE";
                Cells["A2"].SetStyle(styleTitle);

                int row = 3;

                Aspose.Cells.Style styleHeader;
                styleHeader = Cells["A1"].GetStyle();
                styleHeader.Font.Size = 13;
                styleHeader.Font.Name = "Times New Roman";
                styleHeader.ShrinkToFit = false;
                styleHeader.IsTextWrapped = true;
                styleHeader.HorizontalAlignment = TextAlignmentType.Center;
                styleHeader.Font.IsBold = true;
                styleHeader.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleHeader.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleHeader.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleHeader.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                

                Cells[row, 0].Value = "STT";
                Cells[row, 1].Value = "ID";
                Cells[row, 2].Value = "Mã học viên";
                Cells[row, 3].Value = "Tên học viên";
                Cells[row, 4].Value = "Mã phiên học";
                Cells[row, 5].Value = "Thời gian truyền";
                Cells[row, 6].Value = "Băt đầu";
                Cells[row, 7].Value = "Kết thúc";
                Cells[row, 8].Value = "TG đào tạo";
                Cells[row, 9].Value = "QĐ đào tạo";
                Cells[row, 10].Value = "Biến số xe";
                Cells[row, 11].Value = "Trùng HV";
                Cells[row, 12].Value = "Trùng xe";
                Cells[row, 13].Value = "Vi phạm";

                Cells[row, 0].SetStyle(styleHeader);
                Cells[row, 1].SetStyle(styleHeader);
                Cells[row, 2].SetStyle(styleHeader);
                Cells[row, 3].SetStyle(styleHeader);
                Cells[row, 4].SetStyle(styleHeader);
                Cells[row, 5].SetStyle(styleHeader);
                Cells[row, 6].SetStyle(styleHeader);
                Cells[row, 7].SetStyle(styleHeader);
                Cells[row, 8].SetStyle(styleHeader);
                Cells[row, 9].SetStyle(styleHeader);
                Cells[row, 10].SetStyle(styleHeader);
                Cells[row, 11].SetStyle(styleHeader);
                Cells[row, 12].SetStyle(styleHeader);
                Cells[row, 13].SetStyle(styleHeader);


                Aspose.Cells.Style style4;
                style4 = Cells[row + 1, 0].GetStyle();
                style4.Font.Size = 13;
                style4.Font.Name = "Times New Roman";
                style4.ShrinkToFit = false;
                style4.IsTextWrapped = true;
                style4.HorizontalAlignment = TextAlignmentType.Center;
                style4.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Aspose.Cells.Style style8;
                style8 = Cells[row + 1, 0].GetStyle();
                style8.Font.Size = 13;
                style8.Font.Name = "Times New Roman";
                style8.ShrinkToFit = false;
                style8.IsTextWrapped = true;
                style8.HorizontalAlignment = TextAlignmentType.Left;
                style8.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style8.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style8.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style8.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);



                Aspose.Cells.Style styleSum;
                styleSum = Cells[row + 1, 0].GetStyle();
                styleSum.Font.Size = 13;
                styleSum.Font.Name = "Arial";
                styleSum.Font.IsBold = true;
                styleSum.ShrinkToFit = true;

                row++;
                foreach (DataGridViewRow drow in dgvSessionsExcels.Rows)
                {
                    if (drow.Index == dgvSessionsExcels.Rows.Count - 1) break;
                    Cells[row, 0].Value = dgvSessionsExcels.Rows[drow.Index].Cells[0].Value.ToString();
                    Cells[row, 1].Value = dgvSessionsExcels.Rows[drow.Index].Cells[1].Value.ToString();
                    Cells[row, 2].Value = dgvSessionsExcels.Rows[drow.Index].Cells[2].Value.ToString();
                    Cells[row, 3].Value = dgvSessionsExcels.Rows[drow.Index].Cells[3].Value.ToString();
                    Cells[row, 4].Value = dgvSessionsExcels.Rows[drow.Index].Cells[4].Value.ToString();
                    Cells[row, 5].Value = dgvSessionsExcels.Rows[drow.Index].Cells[5].Value.ToString();
                    Cells[row, 6].Value = dgvSessionsExcels.Rows[drow.Index].Cells[6].Value.ToString();
                    Cells[row, 7].Value = dgvSessionsExcels.Rows[drow.Index].Cells[7].Value.ToString();
                    Cells[row, 8].Value = dgvSessionsExcels.Rows[drow.Index].Cells[8].Value.ToString();
                    Cells[row, 9].Value = dgvSessionsExcels.Rows[drow.Index].Cells[9].Value.ToString();
                    Cells[row, 10].Value = dgvSessionsExcels.Rows[drow.Index].Cells[10].Value.ToString();
                    Cells[row, 11].Value = dgvSessionsExcels.Rows[drow.Index].Cells[11].Value.ToString();
                    Cells[row, 12].Value = dgvSessionsExcels.Rows[drow.Index].Cells[12].Value.ToString();
                    Cells[row, 13].Value = dgvSessionsExcels.Rows[drow.Index].Cells[13].Value.ToString();

                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style8);
                    Cells[row, 2].SetStyle(style8);
                    Cells[row, 3].SetStyle(style8);
                    Cells[row, 4].SetStyle(style8);
                    Cells[row, 5].SetStyle(style8);
                    Cells[row, 6].SetStyle(style8);
                    Cells[row, 7].SetStyle(style8);
                    Cells[row, 8].SetStyle(style8);
                    Cells[row, 9].SetStyle(style8);
                    Cells[row, 10].SetStyle(style8);
                    Cells[row, 11].SetStyle(style8);
                    Cells[row, 12].SetStyle(style8);
                    Cells[row, 13].SetStyle(style8);
                    row++;
                }

                Cells.SetColumnWidthPixel(0, 80);
                Cells.SetColumnWidthPixel(1, 80);
                Cells.SetColumnWidthPixel(2, 300);
                Cells.SetColumnWidthPixel(3, 300);
                Cells.SetColumnWidthPixel(4, 460);
                Cells.SetColumnWidthPixel(5, 260);
                Cells.SetColumnWidthPixel(6, 260);
                Cells.SetColumnWidthPixel(7, 260);
                Cells.SetColumnWidthPixel(8, 90);
                Cells.SetColumnWidthPixel(9, 90);
                Cells.SetColumnWidthPixel(10, 110);
                Cells.SetColumnWidthPixel(11, 300);
                Cells.SetColumnWidthPixel(12, 300);
                Cells.SetColumnWidthPixel(13, 300);

                string fileName = "C:\\Report_DAT\\CacPhienTrung_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + ".xls";
                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }
    }
}
