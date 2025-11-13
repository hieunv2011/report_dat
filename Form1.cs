using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;
using RestSharp.Serializers.NewtonsoftJson;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Net;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.Excel;
using static DAT_ToolReports.Form3;
using System.Globalization;

namespace DAT_ToolReports
{
    public partial class Form1 : Form
    {

        private RestClient client;
        private string accessToken;
        public bool bLogined = false;

        public class RestLoginReq
        {
            public string email;
            public string password;
        }
        public class LoginRes
        {
            public string access_token;
            //public string encryptedAccessToken;
            //public int userId;
            //public int expireInSeconds;
        }
        public class CourseRes
        {
            public int id;
            public string ma_khoa_hoc;
            public string ten_khoa_hoc;
            public string ma_hang_dao_tao;
            public int so_hoc_sinh;
            public string ngay_khai_giang;
            public string ngay_be_giang;
            public string ngay_sat_hach;
        }
        public class ResultCoursesRes
        {
            public List<CourseRes> items;
        }
        public List<CourseRes> ListCourses;
        public class TraineeRes
        {
            public int id;
            public string ma_dk;
            public string ho_va_ten;
            public string so_cmt;
            public string ngay_sinh;
            public string anh_chan_dung;
            public string hang_daotao;
            public string rfid_card;
            public int outdoor_hour;
            public int outdoor_distance;
            public int outdoor_session_count;
            public double synced_outdoor_hours;
            public double synced_outdoor_distance;
            public int auto_duration;
            public int night_duration;
        }
        public class ResultTraineeRes
        {
            public List<TraineeRes> items;
        }
        public class SessionRes
        {
            public string? created_date;
            public string? updated_date;
            public int? id;
            public string? session_id;
            public int? state;
            public string? start_time;
            public int? start_date;
            public double? start_lat;
            public double? start_lng;
            public string? start_address;
            public string? end_time;
            public double? end_lat;
            public double? end_lng;
            public string? end_address;
            public int? distance;
            public int? duration;
            public int? instructor_id;
            public string? instructor_name;
            public int? trainee_id;
            public string? trainee_name;
            public string? trainee_ma_dk; 
            public int? device_id;
            public string? device_serial;
            public int? vehicle_id;
            public string? vehicle_plate;
            public string? vehicle_hang;
            public int? faceid_failed_count;
            public int? faceid_success_count;
            public int? gps_count;
            public int? gps_failed_count;
            public Boolean? synced;
            public int sync_status;
            public string? sync_error;
            public string? ten_khoa_hoc;
        }
        public class ResultSessionRes
        {
            public List<SessionRes> items;
        }
        public List<SessionRes> Sessions;
        public class InforSessionReport
        {
            public int? Sessionid;
            public string MaPhienHoc;
            public string StartTime;
            public string StopTime;
            public string ThoiGianTH;
            public string QuangDuongTH;
            public string MaHocVien;
            public string HoTenHocVien;
            public string MaKhoaHoc;
            public string TenKhoaHoc;
            public string LoaiKhoaHoc;
            public string BienSoXe;
            public string HangXeTL;
        }
        public class VehicleRes
        {
            public string? created_date;
            public string? updated_date;
            public int id;
            public string? plate;
            public string? model;
            public string? manufacture_year;
            public string? hang;
            public string? notes;
            public string? gptl;
            public int? device_id;
            public int? branch_id;
            public string? branch_name;
            public int? customer_id;
            public string? customer_name;
            public string? device_name;
            public string? device_serial;
            public string? device_sim;
            public string? device_imei;
            public string? last_updated;
        }
        public class ResultVehicleRes
        {
            public List<VehicleRes> items;
        }
        public List<VehicleRes> Vehicles = new List<VehicleRes>();
        public class HocVienExcel
        {
            public int STT;
            public string MaDangKy;
            public string HoVaTen;
        }
        public List<HocVienExcel> HocVienExcels;

        string[] TrangThaiKiemTra = new string[] { "Vi phạm", "Chưa kiểm tra", "Không vi phạm", "3", "4", "5", "6", "7", "8", "9", "10", "11" };
        public Form1()
        {
            InitializeComponent();
            //if (DAT_ToolReports.Properties.Settings.Default.NightHour1 < 3) 
            //    DAT_ToolReports.Properties.Settings.Default.NightHour1 = 18;
            //if (DAT_ToolReports.Properties.Settings.Default.NightHour2 < 3) 
            //    DAT_ToolReports.Properties.Settings.Default.NightHour2 = 6;
            textEmail.Text = DAT_ToolReports.Properties.Settings.Default.Username;
            textPassword.Text = DAT_ToolReports.Properties.Settings.Default.Password;
            chkbDongBo.Checked = false;
            chkCheckNonOk.Checked = true;
            chkCheckOk.Checked = true;
            chkNonCheck.Checked = true;
            bLogined = false;
            client = new RestClient(DAT_ToolReports.Properties.Settings.Default.linkServer);
            client.UseNewtonsoftJson();
            HocVienExcels = new List<HocVienExcel>();

            if (!Directory.Exists("C:\\Report_DAT"))
                Directory.CreateDirectory("C:\\Report_DAT");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Boolean Retval = ExportWorkbookToPdf("C:\\Img_gen\\123456.xls", "C:\\Img_gen\\123456.pdf");
            PrintMyExcelFile("C:\\Img_gen\\123456.xls");

        }
        public bool ExportWorkbookToPdf(string workbookPath, string outputPath)
        {
            // If either required string is null or empty, stop and bail out
            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(outputPath))
            {
                return false;
            }

            // Create COM Objects
            Microsoft.Office.Interop.Excel.Application excelApplication;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook;

            // Create new instance of Excel
            excelApplication = new Microsoft.Office.Interop.Excel.Application();

            // Make the process invisible to the user
            excelApplication.ScreenUpdating = false;

            // Make the process silent
            excelApplication.DisplayAlerts = false;

            // Open the workbook that you wish to export to PDF
            excelWorkbook = excelApplication.Workbooks.Open(workbookPath);

            // If the workbook failed to open, stop, clean up, and bail out
            if (excelWorkbook == null)
            {
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;

                return false;
            }

            var exportSuccessful = true;
            try
            {
                // Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
                excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (System.Exception ex)
            {
                // Mark the export as failed for the return value...
                exportSuccessful = false;

                // Do something with any exceptions here, if you wish...
                // MessageBox.Show...        
            }
            finally
            {
                // Close the workbook, quit the Excel, and clean up regardless of the results...
                excelWorkbook.Close();
                excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }

            // You can use the following method to automatically open the PDF after export if you wish
            // Make sure that the file actually exists first...
            if (System.IO.File.Exists(outputPath))
            {
                System.Diagnostics.Process.Start(outputPath);
            }

            return exportSuccessful;
        }
        void PrintMyExcelFile(string FilePath)
        {
            //Excel.Application excelApp = new Excel.Application();
            Microsoft.Office.Interop.Excel.Application excelApplication;
            excelApplication = new Microsoft.Office.Interop.Excel.Application();

            // Open the Workbook:
            Microsoft.Office.Interop.Excel.Workbook wb = excelApplication.Workbooks.Open(
                FilePath,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Get the first worksheet.
            // (Excel uses base 1 indexing, not base 0.)
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            // Print out 1 copy to the default printer:
            ws.PrintOut(
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Cleanup:
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(ws);

            wb.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(wb);

            excelApplication.Quit();
            Marshal.FinalReleaseComObject(excelApplication);
        }

        void OpenMyExcelFile(string FilePath)
        {
            //Excel.Application excelApp = new Excel.Application();
            Microsoft.Office.Interop.Excel.Application excelApplication;
            excelApplication = new Microsoft.Office.Interop.Excel.Application();
            excelApplication.Visible = true;
            excelApplication.Workbooks.Open(FilePath);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            bLogined = false;
            if (textEmail.Text == "" && textPassword.Text == "")
            {
                return;
            }
            //txtData.AppendText("LOGIN:\n");
            var request = new RestRequest("/login", Method.POST, DataFormat.Json);
            var body = new RestLoginReq { email = textEmail.Text, password = textPassword.Text };
            request.AddJsonBody(body);
            var response = client.Post<LoginRes>(request);
            accessToken = response.Data.access_token;
            //txtData.AppendText(accessToken + "\r\n");
            client.Authenticator = new JwtAuthenticator(accessToken);
            MessageBox.Show(response.StatusCode.ToString());
            if (response.StatusCode.ToString() == "OK")
            {
                bLogined = true;
                DAT_ToolReports.Properties.Settings.Default.Username = textEmail.Text;
                DAT_ToolReports.Properties.Settings.Default.Password = textPassword.Text;
                DAT_ToolReports.Properties.Settings.Default.Save();
                IRestRequest requestCouse;
                IRestResponse<ResultCoursesRes> responseCouse;
                int nPage = 1;
                dgvCoures.Rows.Clear();
                ListCourses = new List<CourseRes>();
                while (true)
                {
                    requestCouse = new RestRequest("/courses", Method.GET).AddParameter("page_size", 50).AddParameter("page", nPage);
                    responseCouse = client.Get<ResultCoursesRes>(requestCouse);
                    ListCourses.AddRange(responseCouse.Data.items);
                    foreach (CourseRes course in responseCouse.Data.items)
                    {
                        //MessageBox.Show(course.id.ToString() + "---" + course.ma_khoa_hoc + "----" + course.ngay_sat_hach);
                        dgvCoures.Rows.Add(course.id.ToString(), course.ma_khoa_hoc, course.ten_khoa_hoc, course.ma_hang_dao_tao, course.so_hoc_sinh.ToString(), course.ngay_khai_giang, course.ngay_be_giang, course.ngay_sat_hach);
                    }
                    nPage++;
                    if (responseCouse.Data.items.Count < 50)
                    {
                        break;
                    }
                }

                //var request2 = new RestRequest("/courses", Method.GET);
                //var response2 = client.Get<ResultCoursesRes>(request2);
                //dgvCoures.Rows.Clear();
                //foreach (CourseRes course in response2.Data.items)
                //{
                //    //MessageBox.Show(course.id.ToString() + "---" + course.ma_khoa_hoc + "----" + course.ngay_sat_hach);
                //    dgvCoures.Rows.Add(course.id.ToString(), course.ma_khoa_hoc, course.ten_khoa_hoc, course.ma_hang_dao_tao, course.so_hoc_sinh.ToString(), course.ngay_khai_giang, course.ngay_be_giang, course.ngay_sat_hach);
                //}

                dgvVehicles.Rows.Clear();
                Vehicles.Clear();
                var request3 = new RestRequest("/vehicles", Method.GET).AddParameter("page_size", 500);
                var response3 = client.Get<ResultVehicleRes>(request3);
                int Stt = 1;
                Vehicles = response3.Data.items.ToList();
                foreach (VehicleRes vehicle in Vehicles)
                {
                    //MessageBox.Show(course.id.ToString() + "---" + course.ma_khoa_hoc + "----" + course.ngay_sat_hach);
                    dgvVehicles.Rows.Add((Stt++).ToString(), vehicle.plate, vehicle.model, vehicle.hang, vehicle.branch_name);
                }
                //Vehicles = response3.Data.items.ToList();
            }

        }
        String TenKhoaHoc = "";
        String HangDaoTao = "";
        private void dgvCoures_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (accessToken == "")
            {
                return;
            }

            int count = 1;
            int SoHV = Int32.Parse(dgvCoures.CurrentRow.Cells[4].Value.ToString());
            TenKhoaHoc = dgvCoures.CurrentRow.Cells[1].Value.ToString() + " ( " + dgvCoures.CurrentRow.Cells[2].Value.ToString() + " )";
            HangDaoTao = dgvCoures.CurrentRow.Cells[3].Value.ToString();
            dgwTrainees.Rows.Clear();
            IRestRequest request;
            IRestResponse<ResultTraineeRes> response;
            List<TraineeRes> lisTrainees = new List<TraineeRes>();
            if (SoHV <= 50)
            {
                request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 50);
                response = client.Get<ResultTraineeRes>(request);

                if (chkbSortID.Checked)
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_distance).ToList();

            }
            else
            {
                int Page = SoHV / 50 + 1;
                for (int k = 0; k < Page; k++)
                {
                    request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page", k + 1).AddParameter("page_size", 50);
                    response = client.Get<ResultTraineeRes>(request);

                    lisTrainees.AddRange(response.Data.items.ToList());

                }
                if (chkbSortID.Checked)
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_distance).ToList();
            }

            foreach (TraineeRes trainee in lisTrainees)
            {
                count++;
                //indexV = (int)Math.Floor((count - 2) / NumTraineeVehicles);
                //if (indexV > (Vehicles.Count - 1))
                //sPlate = "  ";
                //else
                //    sPlate = Vehicles[indexV].plate;
                dgwTrainees.Rows.Add((count - 1).ToString(), trainee.id.ToString(), trainee.ho_va_ten, trainee.ngay_sinh, trainee.synced_outdoor_hours.ToString(), trainee.synced_outdoor_distance.ToString(),
                    (trainee.outdoor_hour / 3600).ToString(), (trainee.outdoor_distance / 1000).ToString(), trainee.outdoor_session_count.ToString(), trainee.ma_dk, trainee.anh_chan_dung);
            }
        }

        private void dgwTrainees_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo ht = dgwTrainees.HitTest(e.X, e.Y);
            if ((ht.ColumnIndex >= 0) && (ht.RowIndex >= 0))
            {
                dgwTrainees.CurrentCell = dgwTrainees.Rows[ht.RowIndex].Cells[ht.ColumnIndex];
                dgwTrainees.ContextMenuStrip = MenuGetSessions;
            }
            else
            {
                dgwTrainees.ContextMenuStrip = null;
            }
        }
        string sTraineeID = "1";
        string fileNameImage = "";
        private void xemCácPhiênHọcToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("trainee_id", dgwTrainees.CurrentRow.Cells[1].Value.ToString()).AddParameter("page_size", 500).AddParameter("synced", 1).AddQueryParameter("status", "2");
                var response2 = client.Get<ResultSessionRes>(request2);
                sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("trainee_id", dgwTrainees.CurrentRow.Cells[1].Value.ToString()).AddQueryParameter("status", "2").AddParameter("page_size", 500);
                //var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("ho_va_ten", "75009-20240619101753307").AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }


            dgvSessions.Rows.Clear();
            DateTime StartTime;
            string ViPham = "";
            Boolean GetAll = true;
            if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                GetAll = true;
            else
                GetAll = false;
            foreach (SessionRes session in Sessions)
            {
                if ((GetAll == false) && (chkNonCheck.Checked == false) && (session.state == 0))
                    continue;
                if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (session.state < 0))
                    continue;
                if ((GetAll == false) && (chkCheckOk.Checked == false) && (session.state > 0))
                    continue;

                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                if (session.sync_status == 0)
                    ViPham = "Chưa kiểm tra";
                else if (session.sync_status > 0)
                    ViPham = "Không vi phạm";
                else
                    ViPham = session.sync_error;
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
            }
        }
        public static double Truncate(double value, int precision)
        {
            return Math.Truncate(value * Math.Pow(10, precision)) / Math.Pow(10, precision);
        }
        private void CreatFileExcelReport(string fileName, string TraineeName, string MaDK, string NgaySinh, string HangDT, string KhoaHoc, List<SessionRes> LstTmp)
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

                //worksheet.PageSetup.TopMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.LeftMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.RightMargin = Convert.ToDouble(36);

                Aspose.Cells.Cells Cells = worksheet.Cells;

                Aspose.Cells.Style style1;

                style1 = Cells["A1"].GetStyle();
                style1.VerticalAlignment = TextAlignmentType.Center;
                style1.HorizontalAlignment = TextAlignmentType.Center;
                style1.Font.Color = System.Drawing.Color.Black;
                style1.Font.IsBold = true;
                style1.Font.IsItalic = false;
                style1.Font.Size = 12;
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
                int idx;
                string st = DAT_ToolReports.Properties.Settings.Default.Company.ToUpper();
                int nPos = st.IndexOf(DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper());
                if (nPos > 0)
                    st = st.Substring(0, nPos).Trim();

                Cells.Merge(0, 0, 1, 5);
                Cells["A1"].Value = st;
                Cells["A1"].SetStyle(style1);
                Cells.Merge(1, 0, 1, 5);
                Cells["A2"].Value = DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper();
                Cells["A2"].SetStyle(style1);
                Cells.Merge(2, 0, 1, 5);
                Cells["A3"].Value = "***********";
                Cells["A3"].SetStyle(styleDate);

                Cells.Merge(0, 5, 1, 4);
                Cells["F1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                Cells["F1"].SetStyle(style1);
                Cells.Merge(1, 5, 1, 4);
                Cells["F2"].Value = "Độc lập - Tự do - Hạnh phúc";
                Cells["F2"].SetStyle(styleDate);
                Cells.Merge(2, 5, 1, 4);
                Cells["F3"].Value = "***********";
                Cells["F3"].SetStyle(styleDate);

                Aspose.Cells.Style style5;
                style5 = Cells["A1"].GetStyle();
                style5.Font.Size = 13;
                style5.Font.Name = "Times New Roman";
                style5.Font.IsBold = false;
                //style5.ShrinkToFit = true;
                style5.HorizontalAlignment = TextAlignmentType.Left;

                Aspose.Cells.Style style7;
                style7 = Cells["J4"].GetStyle();
                style7.VerticalAlignment = TextAlignmentType.Center;
                style7.HorizontalAlignment = TextAlignmentType.Center;
                style7.Font.Size = 13;
                style7.Font.Name = "Times New Roman";
                style7.Font.IsItalic = true;
                Cells.Merge(3, 5, 1, 4);
                Cells["F4"].Value = "    ";// DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells["F4"].SetStyle(style7);

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(5, 0, 1, 9);
                Cells["A6"].Value = "BÁO CÁO QUÁ TRÌNH ĐÀO TẠO CỦA HỌC VIÊN";
                Cells["A6"].SetStyle(styleTitle);

                Aspose.Cells.Style styleKhoaThi;
                styleKhoaThi = Cells["A7"].GetStyle();
                styleKhoaThi.VerticalAlignment = TextAlignmentType.Center;
                styleKhoaThi.HorizontalAlignment = TextAlignmentType.Center;
                styleKhoaThi.Font.Color = System.Drawing.Color.Black;
                styleKhoaThi.Font.IsBold = true;
                styleKhoaThi.Font.Size = 13;
                styleKhoaThi.Font.Name = "Times New Roman";
                Cells.Merge(6, 0, 1, 9);
                Cells["A7"].Value = "(Ngày báo cáo: " + "ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString() + ")";
                Cells["A7"].SetStyle(styleKhoaThi);

                //========================================
                Cells.Merge(8, 0, 1, 10);
                Cells[8, 0].Value = "I. Thông tin học viên";
                Cells[8, 0].SetStyle(styleKhoaThi);

                Cells.Merge(10, 7, 6, 2);
                //fileNameImage = "C:\\Img_gen\\521316.jpg";// + TraineeNumber.ToString() + ".xls";
                idx = worksheet.Pictures.Add(10, 7, fileNameImage);
                Aspose.Cells.Drawing.Picture pic = worksheet.Pictures[idx];
                double w = worksheet.Cells.GetColumnWidthInch(7) + worksheet.Cells.GetColumnWidthInch(8);
                double h = worksheet.Cells.GetRowHeightInch(10) * 6; //6 dòng
                pic.WidthInch = w;
                pic.HeightInch = h;

                int row = 10;

                Cells[row, 0].Value = "Họ và tên:";
                Cells[row, 0].SetStyle(style5);
                Cells[row, 2].Value = TraineeName;
                Cells[row, 2].SetStyle(style5);
                Cells[row + 1, 0].Value = "Mã học viên:";
                Cells[row + 1, 0].SetStyle(style5);
                Cells[row + 1, 2].Value = MaDK;
                Cells[row + 1, 2].SetStyle(style5);
                Cells[row + 2, 0].Value = "Ngày sinh:";
                Cells[row + 2, 0].SetStyle(style5);
                Cells[row + 2, 2].Value = NgaySinh;
                Cells[row + 2, 2].SetStyle(style5);
                Cells[row + 3, 0].Value = "Hạng đào tạo:";
                Cells[row + 3, 0].SetStyle(style5);
                Cells[row + 3, 2].Value = HangDT;
                Cells[row + 3, 2].SetStyle(style5);
                Cells[row + 4, 0].Value = "Khóa học:";
                Cells[row + 4, 0].SetStyle(style5);
                Cells[row + 4, 2].Value = KhoaHoc;
                Cells[row + 4, 2].SetStyle(style5);
                Cells[row + 5, 0].Value = "Cơ sở đào tạo:";
                Cells[row + 5, 0].SetStyle(style5);
                Cells[row + 5, 2].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                Cells[row + 5, 2].SetStyle(style5);

                Cells.Merge(row + 7, 0, 1, 10);
                Cells[row + 7, 0].Value = "II. Thông tin quá trình đào tạo";
                Cells[row + 7, 0].SetStyle(styleKhoaThi);

                Aspose.Cells.Range range = Cells.CreateRange(10, 0, 6, 9);
                range.SetOutlineBorders(CellBorderType.Thin, System.Drawing.Color.Black);

                row = 19;

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
                //Cells.Merge(row, 1, 1, 2);

                //Cells.Merge(row, 4, 1, 2);
                //Cells.Merge(row, 5, 1, 2);
                //Cells.Merge(row, 7, 1, 2);

                Cells[row, 0].Value = "STT";
                Cells[row, 1].Value = "Phiên đào tạo";
                Cells[row, 2].Value = "Biển số xe tập lái";
                Cells[row, 3].Value = "Hạng xe tập lái";
                Cells[row, 4].Value = "Bắt đầu";
                Cells[row, 5].Value = "Kết thúc";
                Cells[row, 6].Value = "Thời gian đào tạo";
                Cells[row, 7].Value = "Số giờ đêm";
                Cells[row, 8].Value = "Quãng đường đào tạo";
                //Cells[row, 9].Value = "Vi phạm";

                Cells[row, 0].SetStyle(styleHeader);
                Cells[row, 1].SetStyle(styleHeader);
                Cells[row, 2].SetStyle(styleHeader);
                Cells[row, 3].SetStyle(styleHeader);
                Cells[row, 4].SetStyle(styleHeader);
                Cells[row, 5].SetStyle(styleHeader);
                Cells[row, 6].SetStyle(styleHeader);
                Cells[row, 7].SetStyle(styleHeader);
                Cells[row, 8].SetStyle(styleHeader);
                //Cells[row, 9].SetStyle(styleHeader);

                Aspose.Cells.Style style4;
                style4 = Cells[row + 1, 0].GetStyle();
                style4.Font.Size = 13;
                style4.Font.Name = "Times New Roman";
                style4.ShrinkToFit = true;
                style4.HorizontalAlignment = TextAlignmentType.Center;
                style4.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Aspose.Cells.Style style8;
                style8 = Cells[row + 1, 0].GetStyle();
                style8.Font.Size = 13;
                style8.Font.Name = "Times New Roman";
                style8.ShrinkToFit = true;
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
                int Offset = 1;
                int nCountTime = 0;
                int nCountDistance = 0;
                int CountTimeNight = 0;
                int CountTongTimeNight = 0;
                int CountTongTimeAT = 0;
                string sPhut = "";
                string sHangXe = "";
                Boolean bHangAT = false;
                string CheckinH, CheckinM, CheckoutH, CheckoutM, TimelearnH, TimelearnM, TongTGH, TongTGM;
                Boolean GetAll = true;
                if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                    GetAll = true;
                else
                    GetAll = false;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    if ((GetAll == false) && (chkNonCheck.Checked == false) && (LstTmp[i].state == 0))
                        continue;
                    if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (LstTmp[i].state < 0))
                        continue;
                    if ((GetAll == false) && (chkCheckOk.Checked == false) && (LstTmp[i].state > 0))
                        continue;
                    CountTimeNight = 0;
                    SessionRes dateAttendanceTmp = LstTmp[i];
                    nCountTime = (int)(nCountTime + LstTmp[i].duration);
                    nCountDistance = (int)(nCountDistance + LstTmp[i].distance);
                    Cells[row, 0].Value = Offset.ToString();
                    Cells[row, 1].Value = LstTmp[i].session_id;//.ToShortDateString();
                    Cells[row, 2].Value = LstTmp[i].vehicle_plate;
                    //string PlateXe = new String(LstTmp[i].vehicle_plate.Where(Char.IsLetterOrDigit).ToArray());
                    try
                    {
                        sHangXe = LstTmp[i].vehicle_hang;//Vehicles.Find(e => e.plate.ToUpper() == PlateXe.ToUpper()).hang;
                        if ((sHangXe != null) && (sHangXe.Trim().Length > 2))
                        {
                            CountTongTimeAT = CountTongTimeAT + (int)LstTmp[i].duration;
                            bHangAT = true;
                        }
                        else
                        {
                            bHangAT = false;
                            sHangXe = HangDT;
                        }
                    }
                    catch (Exception ex)
                    {
                        bHangAT = false;
                        sHangXe = HangDT;
                    }
                    Cells[row, 3].Value = sHangXe;
                    DateTime StartTime = DateTime.ParseExact(LstTmp[i].start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    Cells[row, 4].Value = StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(); //StartTime.Day.ToString() + "/" + StartTime.Month.ToString() + "/" + StartTime.Year.ToString() + " " +
                        //StartTime.Hour.ToString() + ":" + StartTime.Minute.ToString() + ":" + StartTime.Second.ToString();
                    DateTime StopTime = StartTime.AddSeconds(Convert.ToDouble(LstTmp[i].duration));
                    Cells[row, 5].Value = StopTime.ToShortDateString() + " " + StopTime.ToLongTimeString(); //StopTime.Day.ToString() + "/" + StopTime.Month.ToString() + "/" + StopTime.Year.ToString() + " " +
                        //StopTime.Hour.ToString() + ":" + StopTime.Minute.ToString() + ":" + StopTime.Second.ToString();
                    sPhut = ((LstTmp[i].duration / 60) % 60).ToString();
                    if (sPhut.Length < 2) sPhut = "0" + sPhut;
                    Cells[row, 6].Value = ((LstTmp[i].duration / 60) / 60).ToString() + ":" + sPhut;
                    //Cells[row, 6].Value = Truncate(((double)LstTmp[i].duration / 3600), 3).ToString().Replace(".", ",");

                    int iStartTimeMinute = StartTime.Hour * 60 + StartTime.Minute;
                    int iNightTime1 = DAT_ToolReports.Properties.Settings.Default.NightHour1 * 60 + DAT_ToolReports.Properties.Settings.Default.NightMinute1;
                    int iNightTime2 = DAT_ToolReports.Properties.Settings.Default.NightHour2 * 60 + DAT_ToolReports.Properties.Settings.Default.NightMinute2;
                    if ((DAT_ToolReports.Properties.Settings.Default.NightByStart == false) || ((iStartTimeMinute >= iNightTime1) || (iStartTimeMinute < iNightTime2)))
                    {
                        if ((!bHangAT) || (DAT_ToolReports.Properties.Settings.Default.CountATforNight) || (HangDT.Trim().Length > 2))
                        {
                            if (iStartTimeMinute < iNightTime2)
                            {
                                if ((StartTime.AddSeconds((double)LstTmp[i].duration).Hour * 60 + StartTime.AddSeconds((double)LstTmp[i].duration).Minute) <= iNightTime2) //ket thuc truoc 6h
                                {
                                    CountTimeNight = CountTimeNight + (int)LstTmp[i].duration;
                                }
                                else //ket thuc sau 6 
                                {
                                    CountTimeNight = CountTimeNight + iNightTime2 * 60 - StartTime.Hour * 3600 - StartTime.Minute * 60 - StartTime.Second;
                                }
                            }
                            else if (iStartTimeMinute < iNightTime1)
                            {
                                if ((StartTime.AddSeconds((double)LstTmp[i].duration).Hour * 60 + StartTime.AddSeconds((double)LstTmp[i].duration).Minute) >= iNightTime1) //ket thuc sau 18h
                                {
                                    CountTimeNight = CountTimeNight + (int)LstTmp[i].duration + StartTime.Hour * 3600 + StartTime.Minute * 60 + StartTime.Second - iNightTime1 * 60;
                                }
                            }
                            else
                                CountTimeNight = CountTimeNight + (int)LstTmp[i].duration;
                            if (CountTimeNight > 0)
                            {
                                sPhut = ((CountTimeNight / 60) % 60).ToString();
                                if (sPhut.Length < 2) sPhut = "0" + sPhut;
                                Cells[row, 7].Value = ((CountTimeNight / 60) / 60).ToString() + ":" + sPhut;
                                //Cells[row, 7].Value = Truncate(((double)CountTimeNight / 3600), 3).ToString().Replace(".", ",");
                                CountTongTimeNight = CountTongTimeNight + CountTimeNight;
                            }
                        }
                    }
                    Cells[row, 8].Value = Truncate(((double)LstTmp[i].distance / 1000), 3).ToString().Replace(".", ",");
                    //if (LstTmp[i].sync_status == 0)
                    //    Cells[row, 9].Value = "Chưa kiểm tra";
                    //else if (LstTmp[i].sync_status > 0)
                    //    Cells[row, 9].Value = "Không vi phạm";
                    //else
                    //    Cells[row, 9].Value = LstTmp[i].sync_error;

                    //Cells.Merge(row, 1, 1, 2);
                    //Cells.Merge(row, 4, 1, 2);
                    //Cells.Merge(row, 5, 1, 2);
                    //Cells.Merge(row, 7, 1, 2);

                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style8);
                    Cells[row, 2].SetStyle(style8);
                    Cells[row, 3].SetStyle(style4);
                    Cells[row, 4].SetStyle(style8);
                    Cells[row, 5].SetStyle(style8);
                    Cells[row, 6].SetStyle(style4);
                    Cells[row, 7].SetStyle(style4);
                    Cells[row, 8].SetStyle(style4);
                    //Cells[row, 9].SetStyle(style4);
                    row++;
                    Offset++;
                }
                TongTGH = (nCountTime / 60).ToString();
                if (TongTGH.Length < 2) TongTGH = "0" + TongTGH;
                TongTGM = (nCountTime % 60).ToString();
                if (TongTGM.Length < 2) TongTGM = "0" + TongTGM;

                styleSum.HorizontalAlignment = TextAlignmentType.Center;
                styleSum.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Cells.Merge(row, 0, 1, 6);
                //Cells.Merge(row, 7, 1, 2);

                Cells[row, 0].Value = "Tổng: ";// + sSoGio + "(giờ)";
                sPhut = ((nCountTime / 60) % 60).ToString();
                if (sPhut.Length < 2) sPhut = "0" + sPhut;
                Cells[row, 6].Value = ((nCountTime / 60) / 60).ToString() + ":" + sPhut;
                //Cells[row, 6].Value = Truncate(((double)nCountTime / 3600), 3).ToString().Replace(".", ",");
                sPhut = ((CountTongTimeNight / 60) % 60).ToString();
                if (sPhut.Length < 2) sPhut = "0" + sPhut;
                Cells[row, 7].Value = ((CountTongTimeNight / 60) / 60).ToString() + ":" + sPhut;
                //Cells[row, 7].Value = Truncate(((double)CountTongTimeNight / 3600), 3).ToString().Replace(".", ",");
                Cells[row, 8].Value = Truncate(((double)nCountDistance / 1000), 3).ToString().Replace(".", ",");

                Cells[row, 0].SetStyle(styleSum);
                Cells[row, 1].SetStyle(styleSum);
                Cells[row, 2].SetStyle(styleSum);
                Cells[row, 3].SetStyle(styleSum);
                Cells[row, 4].SetStyle(styleSum);
                Cells[row, 5].SetStyle(styleSum);
                Cells[row, 6].SetStyle(styleSum);
                Cells[row, 7].SetStyle(styleSum);
                Cells[row, 8].SetStyle(styleSum);

                row++;

                Cells.Merge(row, 0, 1, 6);
                Cells.Merge(row, 6, 1, 3);
                Cells[row, 0].Value = "Đủ điều kiện thi";
                Cells[row, 6].Value = "Đạt";

                Cells[row, 0].SetStyle(styleSum);
                Cells[row, 1].SetStyle(styleSum);
                Cells[row, 2].SetStyle(styleSum);
                Cells[row, 3].SetStyle(styleSum);
                Cells[row, 4].SetStyle(styleSum);
                Cells[row, 5].SetStyle(styleSum);
                Cells[row, 6].SetStyle(styleSum);
                Cells[row, 7].SetStyle(styleSum);
                Cells[row, 8].SetStyle(styleSum);
                //Cells[row, 9].SetStyle(styleSum);

                row++;

                Aspose.Cells.Style style6;
                style6 = Cells[row + 2, 0].GetStyle();
                style6.Font.IsBold = true;
                style6.Font.Size = 13;
                style6.Font.Name = "Times New Roman";
                style6.HorizontalAlignment = TextAlignmentType.Center;

                Aspose.Cells.Style style9;
                style9 = Cells[row + 2, 0].GetStyle();
                style9.Font.IsBold = true;
                style9.Font.Size = 13;
                style9.Font.Name = "Times New Roman";
                style9.HorizontalAlignment = TextAlignmentType.Right;

                row++;
                Cells.Merge(row, 3, 1, 2);
                Cells[row, 3].Value = "Tổng hợp kết quả";
                Cells[row, 3].SetStyle(style6);
                Cells.Merge(row, 5, 1, 2);
                Cells[row, 5].Value = "Quãng đường đào tạo";
                Cells[row, 5].SetStyle(style9);
                Cells[row, 7].Value = Truncate(((double)nCountDistance / 1000), 2).ToString().Replace(".", ",");
                Cells[row, 7].SetStyle(style9);
                row++;
                Cells.Merge(row, 5, 1, 2);
                Cells[row, 5].Value = "Thời gian đào tạo";
                Cells[row, 5].SetStyle(style9);
                sPhut = ((nCountTime / 60) % 60).ToString();
                if (sPhut.Length < 2) sPhut = "0" + sPhut;
                Cells[row, 7].Value = ((nCountTime / 60) / 60).ToString() + ":" + sPhut;
                //Cells[row, 7].Value = Truncate(((double)nCountTime / 3600), 2).ToString().Replace(".", ",");
                Cells[row, 7].SetStyle(style9);
                row++;
                Cells.Merge(row, 5, 1, 2);
                Cells[row, 5].Value = "Số giờ đêm";
                Cells[row, 5].SetStyle(style9);
                sPhut = ((CountTongTimeNight / 60) % 60).ToString();
                if (sPhut.Length < 2) sPhut = "0" + sPhut;
                Cells[row, 7].Value = ((CountTongTimeNight / 60) / 60).ToString() + ":" + sPhut;
                //Cells[row, 7].Value = Truncate(((double)CountTongTimeNight / 3600), 2).ToString().Replace(".", ",");
                Cells[row, 7].SetStyle(style9);
                row++;
                Cells.Merge(row, 5, 1, 2);
                Cells[row, 5].Value = "Số giờ tự động";
                Cells[row, 5].SetStyle(style9);
                sPhut = ((CountTongTimeAT / 60) % 60).ToString();
                if (sPhut.Length < 2) sPhut = "0" + sPhut;
                Cells[row, 7].Value = ((CountTongTimeAT / 60) / 60).ToString() + ":" + sPhut;
                //Cells[row, 7].Value = Truncate(((double)CountTongTimeAT / 3600), 2).ToString().Replace(".", ",");
                Cells[row, 7].SetStyle(style9);
                row++;

                Cells.Merge(row, 5, 1, 4);
                Cells[row, 5].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells[row, 5].SetStyle(style7);



                //row++;
                ////Cells.Merge(row, 0, 1, 4);
                ////Cells[row, 0].Value = "Xác nhận của học viên";// "Trưởng phòng đào tạo";
                //Cells[row, 0].SetStyle(style6);
                //Cells.Merge(row, 4, 1, 6); //Cells.Merge(row, 5, 1, 4);
                //Cells[row, 5].Value = "Đại diện cơ sở đào tạo";// "Trưởng phòng đào tạo";
                //Cells[row, 5].SetStyle(style6);

                row++;

                Cells.Merge(row, 0, 1, 3);
                Cells.Merge(row, 3, 1, 3);
                Cells.Merge(row, 6, 1, 3);

                Cells[row, 0].Value = "Xác nhận của trung tâm";
                Cells[row, 3].Value = "Giám sát";
                Cells[row, 6].Value = "Chữ ký học viên";
                Cells[row, 0].SetStyle(style6);
                Cells[row, 3].SetStyle(style6);
                Cells[row, 6].SetStyle(style6);

                row = row + 6;
                Cells.Merge(row, 3, 1, 3);
                Cells[row, 3].Value = "";// "Trần Ngọc Hoàng";
                Cells[row, 3].SetStyle(style6);


                //Cells[row, 0].Value = "Giáo viên                 Học viên";
                //Cells[row, 4].Value = "Người lập biểu         Phòng đào tạo           Giám đốc";
                //Cells[row, 0].SetStyle(style6);
                //Cells[row, 4].SetStyle(style6);


                Cells.SetColumnWidthPixel(0, 50);
                Cells.SetColumnWidthPixel(1, 300);
                Cells.SetColumnWidthPixel(2, 100);
                Cells.SetColumnWidthPixel(3, 60);
                Cells.SetColumnWidthPixel(4, 190);
                Cells.SetColumnWidthPixel(5, 190);
                Cells.SetColumnWidthPixel(6, 80);
                Cells.SetColumnWidthPixel(7, 80);
                Cells.SetColumnWidthPixel(8, 80);
                //Cells.SetColumnWidthPixel(9, 500);

                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }

        private void CreatFileExcelReportCouseSession(string fileName, string HangDT, string KhoaHoc, List<SessionRes> LstTmp)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                object misValue = System.Reflection.Missing.Value;

                worksheet.PageSetup.Orientation = PageOrientationType.Landscape;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0;
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.TopMargin = 1;
                worksheet.PageSetup.BottomMargin = 1;
                worksheet.PageSetup.LeftMargin = 1;
                worksheet.PageSetup.RightMargin = 0.3;

                //worksheet.PageSetup.TopMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.LeftMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.RightMargin = Convert.ToDouble(36);

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
                int idx;
                string st = DAT_ToolReports.Properties.Settings.Default.Company.ToUpper();
                int nPos = st.IndexOf(DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper());
                if (nPos > 0)
                    st = st.Substring(0, nPos).Trim();

                Cells.Merge(0, 0, 1, 6);
                Cells["A1"].Value = st;
                Cells["A1"].SetStyle(style1);
                Cells.Merge(1, 0, 1, 6);
                Cells["A2"].Value = DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper();
                Cells["A2"].SetStyle(style1);
                Cells.Merge(2, 0, 1, 6);
                Cells["A3"].Value = "***********";
                Cells["A3"].SetStyle(styleDate);

                Cells.Merge(0, 9, 1, 4);
                Cells["J1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                Cells["J1"].SetStyle(style1);
                Cells.Merge(1, 9, 1, 4);
                Cells["J2"].Value = "Độc lập - Tự do - Hạnh phúc";
                Cells["J2"].SetStyle(styleDate);
                Cells.Merge(2, 6, 1, 4);
                Cells["J3"].Value = "***********";
                Cells["J3"].SetStyle(styleDate);

                Aspose.Cells.Style style5;
                style5 = Cells["A1"].GetStyle();
                style5.Font.Size = 13;
                style5.Font.Name = "Times New Roman";
                style5.Font.IsBold = false;
                //style5.ShrinkToFit = true;
                style5.HorizontalAlignment = TextAlignmentType.Left;

                Aspose.Cells.Style style7;
                style7 = Cells["J4"].GetStyle();
                style7.VerticalAlignment = TextAlignmentType.Center;
                style7.HorizontalAlignment = TextAlignmentType.Center;
                style7.Font.Size = 13;
                style7.Font.Name = "Times New Roman";
                style7.Font.IsItalic = true;
                Cells.Merge(3, 5, 1, 4);
                Cells["F4"].Value = "    ";// DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells["F4"].SetStyle(style7);

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(5, 0, 1, 13);
                Cells["A6"].Value = "BÁO CÁO PHIÊN HỌC";
                Cells["A6"].SetStyle(styleTitle);

                Aspose.Cells.Style styleKhoaThi;
                styleKhoaThi = Cells["A7"].GetStyle();
                styleKhoaThi.VerticalAlignment = TextAlignmentType.Center;
                styleKhoaThi.HorizontalAlignment = TextAlignmentType.Center;
                styleKhoaThi.Font.Color = System.Drawing.Color.Black;
                styleKhoaThi.Font.IsBold = true;
                styleKhoaThi.Font.Size = 13;
                styleKhoaThi.Font.Name = "Times New Roman";
                Cells.Merge(6, 0, 1, 13);
                Cells["A7"].Value = "(Ngày báo cáo: " + "ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString() + ")";
                Cells["A7"].SetStyle(styleKhoaThi);

                //========================================
                Cells.Merge(8, 0, 1, 13);
                Cells[8, 0].Value = "I. Thông tin Khóa học";
                Cells[8, 0].SetStyle(styleKhoaThi);

                //Cells.Merge(10, 7, 6, 2);
                ////fileNameImage = "C:\\Img_gen\\521316.jpg";// + TraineeNumber.ToString() + ".xls";
                //idx = worksheet.Pictures.Add(10, 7, fileNameImage);
                //Picture pic = worksheet.Pictures[idx];
                //double w = worksheet.Cells.GetColumnWidthInch(7) + worksheet.Cells.GetColumnWidthInch(8);
                //double h = worksheet.Cells.GetRowHeightInch(10) * 6; //6 dòng
                //pic.WidthInch = w;
                //pic.HeightInch = h;

                int row = 10;

                Cells[row, 0].Value = "Trung tâm đào tạo:";
                Cells[row, 0].SetStyle(style5);
                Cells[row, 2].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                Cells[row, 2].SetStyle(style5);
                Cells[row + 1, 0].Value = "Đơn vị truyền dữ liệu:";
                Cells[row + 1, 0].SetStyle(style5);
                Cells[row + 1, 2].Value = "Công ty CP Công nghệ Sát hạch Toàn Phương";
                Cells[row + 1, 2].SetStyle(style5);
                Cells[row + 2, 0].Value = "Khóa học:";
                Cells[row + 2, 0].SetStyle(style5);
                Cells[row + 2, 2].Value = KhoaHoc;
                Cells[row + 2, 2].SetStyle(style5);
                Cells[row + 3, 0].Value = "Hạng đào tạo:";
                Cells[row + 3, 0].SetStyle(style5);
                Cells[row + 3, 2].Value = HangDT;
                Cells[row + 3, 2].SetStyle(style5);
                Cells[row + 4, 0].Value = "Thời gian xuất dữ liệu:";
                Cells[row + 4, 0].SetStyle(style5);
                Cells[row + 4, 2].Value = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString() + " " +
                    DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":" + DateTime.Now.Second.ToString();
                Cells[row + 4, 2].SetStyle(style5);
                //Cells[row + 5, 0].Value = "Cơ sở đào tạo:";
                //Cells[row + 5, 0].SetStyle(style5);
                //Cells[row + 5, 2].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                //Cells[row + 5, 2].SetStyle(style5);

                Cells.Merge(row + 6, 0, 1, 13);
                Cells[row + 6, 0].Value = "II. Thông tin quá trình đào tạo";
                Cells[row + 6, 0].SetStyle(styleKhoaThi);

                Aspose.Cells.Range range = Cells.CreateRange(10, 0, 5, 13);
                range.SetOutlineBorders(CellBorderType.Thin, System.Drawing.Color.Black);

                row = 18;

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
                Cells.Merge(row, 1, 1, 2);

                //Cells.Merge(row, 4, 1, 2);
                //Cells.Merge(row, 5, 1, 2);
                //Cells.Merge(row, 7, 1, 2);

                Cells[row, 0].Value = "STT";
                Cells[row, 1].Value = "Mã phiên học";
                Cells[row, 3].Value = "Trạng thái";
                Cells[row, 4].Value = "Thời gian truyền";
                Cells[row, 5].Value = "Thời gian phiên học";
                Cells[row, 6].Value = "Mã học viên";
                Cells[row, 7].Value = "Tên học viên";
                Cells[row, 8].Value = "Hạng xe tập lái";
                Cells[row, 9].Value = "Thời gian đào tạo(h)";
                Cells[row, 10].Value = "Quãng đường đào tạo(Km)";
                Cells[row, 11].Value = "Biển xố xe tập lái";
                Cells[row, 12].Value = "Vi phạm";

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
                style8.ShrinkToFit = true;
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
                int Offset = 1;
                string TrangThai = "";
                string ViPham = "";
                string sHangXe = "";
                //string CheckinH, CheckinM, CheckoutH, CheckoutM, TimelearnH, TimelearnM, TongTGH, TongTGM;
                Boolean GetAll = true;
                if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                    GetAll = true;
                else
                    GetAll = false;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    if ((GetAll == false) && (chkNonCheck.Checked == false) && (LstTmp[i].sync_status == 0))
                        continue;
                    if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (LstTmp[i].sync_status < 0))
                        continue;
                    if ((GetAll == false) && (chkCheckOk.Checked == false) && (LstTmp[i].sync_status > 0))
                        continue;

                    if (LstTmp[i].sync_status == 0)
                    {
                        TrangThai = "Chưa kiểm tra";
                        ViPham = "Chưa xác định được";
                    }
                    else if (LstTmp[i].sync_status > 0)
                    {
                        TrangThai = "Không vi phạm";
                        ViPham = LstTmp[i].sync_error;
                    }
                    else
                    {
                        TrangThai = "Vi Phạm";
                        ViPham = LstTmp[i].sync_error;
                    }

                    Cells[row, 0].Value = Offset.ToString();
                    Cells[row, 1].Value = LstTmp[i].session_id;//.ToShortDateString();
                    Cells[row, 3].Value = TrangThai;
                    //string PlateXe = new String(LstTmp[i].vehicle_plate.Where(Char.IsLetterOrDigit).ToArray());
                    try
                    {
                        sHangXe = LstTmp[i].vehicle_hang;//Vehicles.Find(e => e.plate.ToUpper() == PlateXe.ToUpper()).hang;
                    }
                    catch (Exception ex)
                    {
                        sHangXe = HangDT;
                    }
                    //Cells[row, 4].Value = "ThoigianTruyen";
                    DateTime StartTime = DateTime.ParseExact(LstTmp[i].start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    Cells[row, 5].Value = StartTime.Day.ToString() + "/" + StartTime.Month.ToString() + "/" + StartTime.Year.ToString() + " " +
                        StartTime.Hour.ToString() + ":" + StartTime.Minute.ToString() + ":" + StartTime.Second.ToString();


                    Cells[row, 6].Value = LstTmp[i].trainee_ma_dk;//.trainee_id.ToString();
                    Cells[row, 7].Value = LstTmp[i].trainee_name;
                    Cells[row, 8].Value = sHangXe;
                    Cells[row, 9].Value = Truncate(((double)LstTmp[i].duration / 3600), 2).ToString().Replace(".", ",");
                    Cells[row, 10].Value = Truncate(((double)LstTmp[i].distance / 1000), 2).ToString().Replace(".", ",");
                    Cells[row, 11].Value = LstTmp[i].vehicle_plate;
                    Cells[row, 12].Value = ViPham;

                    Cells.Merge(row, 1, 1, 2);


                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style8);
                    Cells[row, 2].SetStyle(style8);
                    Cells[row, 3].SetStyle(style8);
                    Cells[row, 4].SetStyle(style4);
                    Cells[row, 5].SetStyle(style8);
                    Cells[row, 6].SetStyle(style4);
                    Cells[row, 7].SetStyle(style4);
                    Cells[row, 8].SetStyle(style4);
                    Cells[row, 9].SetStyle(style4);
                    Cells[row, 10].SetStyle(style4);
                    Cells[row, 11].SetStyle(style4);
                    Cells[row, 12].SetStyle(style4);
                    row++;
                    Offset++;
                }


                styleSum.HorizontalAlignment = TextAlignmentType.Center;
                styleSum.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Cells.Merge(row, 0, 1, 6);
                //Cells.Merge(row, 7, 1, 2);


                row++;

                Aspose.Cells.Style style6;
                style6 = Cells[row + 2, 0].GetStyle();
                style6.Font.IsBold = true;
                style6.Font.Size = 13;
                style6.Font.Name = "Times New Roman";
                style6.HorizontalAlignment = TextAlignmentType.Center;

                Aspose.Cells.Style style9;
                style9 = Cells[row + 2, 0].GetStyle();
                style9.Font.IsBold = true;
                style9.Font.Size = 13;
                style9.Font.Name = "Times New Roman";
                style9.HorizontalAlignment = TextAlignmentType.Right;

                Cells.SetColumnWidthPixel(0, 100);
                Cells.SetColumnWidthPixel(1, 180);
                Cells.SetColumnWidthPixel(2, 120);
                Cells.SetColumnWidthPixel(3, 100);
                Cells.SetColumnWidthPixel(4, 100);
                Cells.SetColumnWidthPixel(5, 190);
                Cells.SetColumnWidthPixel(6, 100);
                Cells.SetColumnWidthPixel(7, 300);
                Cells.SetColumnWidthPixel(8, 100);
                Cells.SetColumnWidthPixel(9, 100);
                Cells.SetColumnWidthPixel(10, 100);
                Cells.SetColumnWidthPixel(11, 100);
                Cells.SetColumnWidthPixel(12, 500);

                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }

        private Bitmap DrawFilledRectangle(int x, int y)
        {
            Bitmap bmp = new Bitmap(x, y);
            using (Graphics graph = Graphics.FromImage(bmp))
            {
                System.Drawing.Rectangle ImageSize = new System.Drawing.Rectangle(0, 0, x, y);
                graph.FillRectangle(Brushes.White, ImageSize);
            }
            return bmp;
        }

        private void xuấtRaFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("trainee_id", dgwTrainees.CurrentRow.Cells[1].Value.ToString()).AddParameter("page_size", 500).AddParameter("synced", 1).AddQueryParameter("status", "2");
                var response2 = client.Get<ResultSessionRes>(request2);
                sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("trainee_id", dgwTrainees.CurrentRow.Cells[1].Value.ToString()).AddQueryParameter("status", "2").AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }
            WebClient webClient = new WebClient();
            string LinkFile = "";
            int row = dgwTrainees.CurrentRow.Index;
            if (dgwTrainees[10, row].Value is null)
            {
                LinkFile = "NULL";
            }
            else
            {
                LinkFile = dgwTrainees[10, row].Value.ToString();
            }
            if (LinkFile.Length > 5)
            {
                fileNameImage = "C:\\Report_DAT\\" + dgwTrainees[1, row].Value.ToString() + ".jpg";
                try
                {
                    webClient.DownloadFile(LinkFile, fileNameImage);
                }
                catch (Exception)
                {
                    fileNameImage = "C:\\Report_DAT\\1234.jpg";
                }
            }
            else
            {
                fileNameImage = "C:\\Report_DAT\\1234.jpg";
            }
            if (!(System.IO.File.Exists(fileNameImage)))
            {
                Bitmap bmp = DrawFilledRectangle(680, 480);
                bmp.Save(fileNameImage, ImageFormat.Jpeg);
            }

            string fileName = "C:\\Report_DAT\\" + dgwTrainees[1, row].Value.ToString() + ".xls";
            CreatFileExcelReport(fileName, dgwTrainees.CurrentRow.Cells[2].Value.ToString(), dgwTrainees.CurrentRow.Cells[9].Value.ToString(), dgwTrainees.CurrentRow.Cells[3].Value.ToString(), HangDaoTao, TenKhoaHoc, Sessions);
            dgvSessions.Rows.Clear();
            DateTime StartTime;
            string ViPham = "";
            Boolean GetAll = true;
            if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                GetAll = true;
            else
                GetAll = false;
            foreach (SessionRes session in Sessions)
            {
                if ((GetAll == false) && (chkNonCheck.Checked == false) && (session.sync_status == 0))
                    continue;
                if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (session.sync_status < 0))
                    continue;
                if ((GetAll == false) && (chkCheckOk.Checked == false) && (session.sync_status > 0))
                    continue;

                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                if (session.sync_status == 0)
                    ViPham = "Chưa kiểm tra";
                else if (session.sync_status > 0)
                    ViPham = "Không vi phạm";
                else
                    ViPham = session.sync_error;
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
            }
            OpenMyExcelFile(fileName);
        }

        private void inBáoCáoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int row = dgwTrainees.CurrentRow.Index;
            string fileName = "C:\\Report_DAT\\" + dgwTrainees[1, row].Value.ToString() + ".xls";
            if (System.IO.File.Exists(fileName))
            {
                PrintMyExcelFile(fileName);
            }
            else
            {
                if (chkbDongBo.Checked)
                {
                    var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("trainee_id", dgwTrainees.CurrentRow.Cells[1].Value.ToString()).AddParameter("page_size", 500).AddParameter("synced", 1).AddQueryParameter("status", "2");
                    var response2 = client.Get<ResultSessionRes>(request2);
                    sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                    Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
                }
                else
                {
                    var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("trainee_id", dgwTrainees.CurrentRow.Cells[1].Value.ToString()).AddQueryParameter("status", "2").AddParameter("page_size", 500);
                    var response3 = client.Get<ResultSessionRes>(request3);
                    sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                    Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
                }
                WebClient webClient = new WebClient();
                string LinkFile = "";
                int row2 = dgwTrainees.CurrentRow.Index;
                LinkFile = dgwTrainees[10, row2].Value.ToString();
                fileNameImage = "C:\\Report_DAT\\" + dgwTrainees[1, row].Value.ToString() + ".jpg";
                webClient.DownloadFile(LinkFile, fileNameImage);
                string fileName2 = "C:\\Report_DAT\\" + dgwTrainees[1, row].Value.ToString() + ".xls";
                CreatFileExcelReport(fileName2, dgwTrainees.CurrentRow.Cells[2].Value.ToString(), dgwTrainees.CurrentRow.Cells[9].Value.ToString(), dgwTrainees.CurrentRow.Cells[3].Value.ToString(), HangDaoTao, TenKhoaHoc, Sessions);
                dgvSessions.Rows.Clear();
                DateTime StartTime;
                foreach (SessionRes session in Sessions)
                {
                    StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                    dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                        Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                        session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString());
                }
                if (System.IO.File.Exists(fileName2))
                {
                    PrintMyExcelFile(fileName2);
                }
            }

        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void xemCácPhiênHọcToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            string monthfrom, dayfrom, monthto, dayto;
            monthfrom = dtpFrom.Value.Month.ToString();
            if (monthfrom.Length < 2) monthfrom = "0" + monthfrom;
            dayfrom = dtpFrom.Value.Day.ToString();
            if (dayfrom.Length < 2) dayfrom = "0" + dayfrom;
            monthto = dtpTo.Value.Month.ToString();
            if (monthto.Length < 2) monthto = "0" + monthto;
            dayto = dtpTo.Value.Day.ToString();
            if (dayto.Length < 2) dayto = "0" + dayto;


            string from_time = dtpFrom.Value.Year.ToString() + "-" + monthfrom + "-" + dayfrom + "T00:00:00";
            string to_time = dtpTo.Value.Year.ToString() + "-" + monthto + "-" + dayto + "T23:59:59";
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("plate", dgvVehicles.CurrentRow.Cells[1].Value.ToString())
                    .AddQueryParameter("from_date", from_time).AddQueryParameter("to_date", to_time)
                    .AddParameter("page_size", 500).AddParameter("synced", 1);
                var response2 = client.Get<ResultSessionRes>(request2);
                //sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("plate", dgvVehicles.CurrentRow.Cells[1].Value.ToString())
                    .AddQueryParameter("from_date", from_time).AddQueryParameter("to_date", to_time)
                    .AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                //sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }


            dgvSessions.Rows.Clear();
            DateTime StartTime;
            string ViPham = "";
            Boolean GetAll = true;
            if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                GetAll = true;
            else
                GetAll = false;
            foreach (SessionRes session in Sessions)
            {
                if ((GetAll == false) && (chkNonCheck.Checked == false) && (session.sync_status == 0))
                    continue;
                if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (session.sync_status < 0))
                    continue;
                if ((GetAll == false) && (chkCheckOk.Checked == false) && (session.sync_status > 0))
                    continue;

                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                if (session.sync_status == 0)
                    ViPham = "Chưa kiểm tra";
                else if (session.sync_status > 0)
                    ViPham = "Không vi phạm";
                else
                    ViPham = session.sync_error;
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
            }

        }

        private void dgvVehicles_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo ht = dgvVehicles.HitTest(e.X, e.Y);
            if ((ht.ColumnIndex >= 0) && (ht.RowIndex >= 0))
            {
                dgvVehicles.CurrentCell = dgvVehicles.Rows[ht.RowIndex].Cells[ht.ColumnIndex];
                dgvVehicles.ContextMenuStrip = MenuGetSessionsVehicle;
            }
            else
            {
                dgvVehicles.ContextMenuStrip = null;
            }
        }

        private void CreatFileExcelReportVehicle(string fileName, string Plate, string HangDT, string ChuSoHuu, string CSDT, string KhoangThoiGian, List<SessionRes> LstTmp)
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

                //worksheet.PageSetup.TopMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.LeftMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.RightMargin = Convert.ToDouble(36);

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
                int idx;
                string st = DAT_ToolReports.Properties.Settings.Default.Company.ToUpper();
                int nPos = st.IndexOf(DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper());
                if (nPos > 0)
                    st = st.Substring(0, nPos).Trim();

                Cells.Merge(0, 0, 1, 5);
                Cells["A1"].Value = st;
                Cells["A1"].SetStyle(style1);
                Cells.Merge(1, 0, 1, 5);
                Cells["A2"].Value = DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper();
                Cells["A2"].SetStyle(style1);
                Cells.Merge(2, 0, 1, 5);
                Cells["A3"].Value = "***********";
                Cells["A3"].SetStyle(styleDate);

                Cells.Merge(0, 5, 1, 4);
                Cells["F1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                Cells["F1"].SetStyle(style1);
                Cells.Merge(1, 5, 1, 4);
                Cells["F2"].Value = "Độc lập - Tự do - Hạnh phúc";
                Cells["F2"].SetStyle(styleDate);
                Cells.Merge(2, 5, 1, 4);
                Cells["F3"].Value = "***********";
                Cells["F3"].SetStyle(styleDate);

                Aspose.Cells.Style style5;
                style5 = Cells["A1"].GetStyle();
                style5.Font.Size = 13;
                style5.Font.Name = "Times New Roman";
                style5.Font.IsBold = false;
                //style5.ShrinkToFit = true;
                style5.HorizontalAlignment = TextAlignmentType.Left;

                Aspose.Cells.Style style7;
                style7 = Cells["J4"].GetStyle();
                style7.VerticalAlignment = TextAlignmentType.Center;
                style7.HorizontalAlignment = TextAlignmentType.Center;
                style7.Font.Size = 13;
                style7.Font.Name = "Times New Roman";
                style7.Font.IsItalic = true;
                Cells.Merge(3, 5, 1, 4);
                Cells["F4"].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells["F4"].SetStyle(style7);

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(5, 0, 1, 9);
                Cells["A6"].Value = "BÁO CÁO QUÁ TRÌNH ĐÀO TẠO CỦA XE TẬP LÁI";
                Cells["A6"].SetStyle(styleTitle);

                Aspose.Cells.Style styleKhoaThi;
                styleKhoaThi = Cells["A7"].GetStyle();
                styleKhoaThi.VerticalAlignment = TextAlignmentType.Center;
                styleKhoaThi.HorizontalAlignment = TextAlignmentType.Center;
                styleKhoaThi.Font.Color = System.Drawing.Color.Black;
                styleKhoaThi.Font.IsBold = true;
                styleKhoaThi.Font.Size = 13;
                styleKhoaThi.Font.Name = "Times New Roman";
                Cells.Merge(6, 0, 1, 9);
                Cells["A7"].Value = "(Ngày báo cáo: " + "ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString() + ")";
                Cells["A7"].SetStyle(styleKhoaThi);

                //========================================
                Cells.Merge(8, 0, 1, 9);
                Cells[8, 0].Value = "I. Thông tin xe tập lái";
                Cells[8, 0].SetStyle(styleKhoaThi);

                //Cells.Merge(10, 7, 6, 2);
                ////fileNameImage = "C:\\Img_gen\\521316.jpg";// + TraineeNumber.ToString() + ".xls";
                //idx = worksheet.Pictures.Add(10, 7, fileNameImage);
                //Picture pic = worksheet.Pictures[idx];
                //double w = worksheet.Cells.GetColumnWidthInch(7) + worksheet.Cells.GetColumnWidthInch(8);
                //double h = worksheet.Cells.GetRowHeightInch(10) * 6; //6 dòng
                //pic.WidthInch = w;
                //pic.HeightInch = h;

                int row = 10;

                Cells[row, 0].Value = "Biển số xe:";
                Cells[row, 0].SetStyle(style5);
                Cells[row, 2].Value = Plate;
                Cells[row, 2].SetStyle(style5);

                Cells[row + 1, 0].Value = "Hạng xe:";
                Cells[row + 1, 0].SetStyle(style5);
                Cells[row + 1, 2].Value = HangDT;
                Cells[row + 1, 2].SetStyle(style5);

                Cells[row + 2, 0].Value = "Chủ sở hữu:";
                Cells[row + 2, 0].SetStyle(style5);
                Cells[row + 2, 2].Value = ChuSoHuu;
                Cells[row + 2, 2].SetStyle(style5);

                Cells[row + 3, 0].Value = "Cơ sở đào tạo:";
                Cells[row + 3, 0].SetStyle(style5);
                Cells[row + 3, 2].Value = CSDT;
                Cells[row + 3, 2].SetStyle(style5);

                Cells[row + 4, 0].Value = KhoangThoiGian;
                Cells[row + 4, 0].SetStyle(style5);
                //Cells[row + 4, 2].Value = KhoaHoc;
                //Cells[row + 4, 2].SetStyle(style5);

                //Cells[row + 5, 0].Value = "Cơ sở đào tạo:";
                //Cells[row + 5, 0].SetStyle(style5);
                //Cells[row + 5, 2].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                //Cells[row + 5, 2].SetStyle(style5);

                Cells.Merge(row + 7, 0, 1, 9);
                Cells[row + 7, 0].Value = "II. Thông tin quá trình đào tạo";
                Cells[row + 7, 0].SetStyle(styleKhoaThi);

                Aspose.Cells.Range range = Cells.CreateRange(10, 0, 5, 9);
                range.SetOutlineBorders(CellBorderType.Thin, System.Drawing.Color.Black);

                row = 18;

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
                Cells.Merge(row, 1, 1, 2);
                Cells.Merge(row, 4, 1, 2);
                //Cells.Merge(row, 5, 1, 2);
                Cells.Merge(row, 7, 1, 2);
                Cells[row, 0].Value = "STT";
                Cells[row, 1].Value = "Phiên đào tạo";
                Cells[row, 3].Value = "Học viên tập lái";
                Cells[row, 4].Value = "Ngày đào tạo";
                Cells[row, 6].Value = "Thời gian đào tạo";
                Cells[row, 7].Value = "Quãng đường đào tạo";

                Cells[row, 0].SetStyle(styleHeader);
                Cells[row, 1].SetStyle(styleHeader);
                Cells[row, 2].SetStyle(styleHeader);
                Cells[row, 3].SetStyle(styleHeader);
                Cells[row, 4].SetStyle(styleHeader);
                Cells[row, 5].SetStyle(styleHeader);
                Cells[row, 6].SetStyle(styleHeader);
                Cells[row, 7].SetStyle(styleHeader);
                Cells[row, 8].SetStyle(styleHeader);

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
                int Offset = 1;
                int nCountTime = 0;
                int nCountDistance = 0;
                string CheckinH, CheckinM, CheckoutH, CheckoutM, TimelearnH, TimelearnM, TongTGH, TongTGM;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    SessionRes dateAttendanceTmp = LstTmp[i];
                    nCountTime = (int)(nCountTime + LstTmp[i].duration);
                    nCountDistance = (int)(nCountDistance + LstTmp[i].distance);
                    Cells[row, 0].Value = Offset.ToString();
                    Cells[row, 1].Value = LstTmp[i].session_id;//.ToShortDateString();
                    Cells[row, 3].Value = LstTmp[i].trainee_name;
                    DateTime StartTime = DateTime.ParseExact(LstTmp[i].start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    Cells[row, 4].Value = StartTime.ToShortDateString() + "-" + StartTime.ToShortTimeString();
                    Cells[row, 6].Value = Truncate(((double)LstTmp[i].duration / 3600), 2).ToString();
                    Cells[row, 7].Value = Truncate(((double)LstTmp[i].distance / 1000), 2).ToString();


                    Cells.Merge(row, 1, 1, 2);
                    Cells.Merge(row, 4, 1, 2);
                    //Cells.Merge(row, 5, 1, 2);
                    Cells.Merge(row, 7, 1, 2);
                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style8);
                    Cells[row, 2].SetStyle(style8);
                    Cells[row, 3].SetStyle(style8);
                    Cells[row, 4].SetStyle(style8);
                    Cells[row, 5].SetStyle(style4);
                    Cells[row, 6].SetStyle(style4);
                    Cells[row, 7].SetStyle(style4);
                    Cells[row, 8].SetStyle(style4);
                    row++;
                    Offset++;
                }
                TongTGH = (nCountTime / 60).ToString();
                if (TongTGH.Length < 2) TongTGH = "0" + TongTGH;
                TongTGM = (nCountTime % 60).ToString();
                if (TongTGM.Length < 2) TongTGM = "0" + TongTGM;

                styleSum.HorizontalAlignment = TextAlignmentType.Center;
                styleSum.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                styleSum.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Cells.Merge(row, 0, 1, 6);
                Cells.Merge(row, 7, 1, 2);
                Cells[row, 0].Value = "Tổng: ";// + sSoGio + "(giờ)";
                Cells[row, 6].Value = ((nCountTime / 60) / 60).ToString() + ":" + ((nCountTime / 60) % 60).ToString();
                Cells[row, 7].Value = ((double)nCountDistance / 1000).ToString();

                Cells[row, 0].SetStyle(styleSum);
                Cells[row, 1].SetStyle(styleSum);
                Cells[row, 2].SetStyle(styleSum);
                Cells[row, 3].SetStyle(styleSum);
                Cells[row, 4].SetStyle(styleSum);
                Cells[row, 5].SetStyle(styleSum);
                Cells[row, 6].SetStyle(styleSum);
                Cells[row, 7].SetStyle(styleSum);
                Cells[row, 8].SetStyle(styleSum);

                row++;

                //Range range = Cells.CreateRange(1 + 1, 0, row - 1 - 2, 9);
                //range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

                Cells.Merge(row, 0, 1, 9);
                //Cells[row, 0].Value = LstTmp.Count.ToString();
                //Cells[row, 0].SetStyle(styleSum);
                row++;


                Aspose.Cells.Style style6;
                style6 = Cells[row + 2, 0].GetStyle();
                style6.Font.IsBold = true;
                style6.Font.Size = 13;
                style6.Font.Name = "Times New Roman";
                style6.HorizontalAlignment = TextAlignmentType.Center;

                Cells[row + 2, 0].Value = "  ";
                Cells[row + 3, 0].Value = "  ";
                Cells.Merge(row + 2, 0, 1, 4);
                Cells.Merge(row + 3, 0, 1, 4);
                Cells[row + 2, 0].SetStyle(style6);
                Cells[row + 3, 0].SetStyle(style7);

                //Cells[row + 2, 4].Value = "Tổ trưởng sát hạch";
                //Cells.Merge(row + 2, 4, 1, 3);
                //Cells[row + 2, 4].SetStyle(style6);
                Cells.Merge(row + 1, 6, 1, 3);
                Cells[row + 1, 6].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells[row + 1, 6].SetStyle(style7);

                Cells[row + 2, 6].Value = "Trưởng phòng đào tạo";
                Cells[row + 3, 6].Value = "(ký tên)";
                Cells.Merge(row + 2, 6, 1, 3);
                Cells.Merge(row + 3, 6, 1, 3);
                Cells[row + 2, 6].SetStyle(style6);
                Cells[row + 3, 6].SetStyle(style7);

                Cells.SetColumnWidthPixel(0, 100);
                Cells.SetColumnWidthPixel(1, 180);
                Cells.SetColumnWidthPixel(2, 120);
                Cells.SetColumnWidthPixel(3, 190);
                //worksheet.AutoFitColumn(3);
                Cells.SetColumnWidthPixel(4, 10);
                Cells.SetColumnWidthPixel(5, 200);
                Cells.SetColumnWidthPixel(6, 145);
                Cells.SetColumnWidthPixel(7, 100);
                Cells.SetColumnWidthPixel(8, 45);

                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }

        private void xuấtRaFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string monthfrom, dayfrom, monthto, dayto;
            monthfrom = dtpFrom.Value.Month.ToString();
            if (monthfrom.Length < 2) monthfrom = "0" + monthfrom;
            dayfrom = dtpFrom.Value.Day.ToString();
            if (dayfrom.Length < 2) dayfrom = "0" + dayfrom;
            monthto = dtpTo.Value.Month.ToString();
            if (monthto.Length < 2) monthto = "0" + monthto;
            dayto = dtpTo.Value.Day.ToString();
            if (dayto.Length < 2) dayto = "0" + dayto;


            string from_time = dtpFrom.Value.Year.ToString() + "-" + monthfrom + "-" + dayfrom + "T00:00:00";
            string to_time = dtpTo.Value.Year.ToString() + "-" + monthto + "-" + dayto + "T23:59:59";
            string Khoangthoigian = "Từ ngày " + dayfrom + "/" + monthfrom + "/" + dtpFrom.Value.Year.ToString() + " đến ngày " + dayto + "/" + monthto + "/" + dtpTo.Value.Year.ToString();
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("plate", dgvVehicles.CurrentRow.Cells[1].Value.ToString())
                    .AddQueryParameter("from_date", from_time).AddQueryParameter("to_date", to_time)
                    .AddParameter("page_size", 500).AddParameter("synced", 1);
                var response2 = client.Get<ResultSessionRes>(request2);
                //sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("plate", dgvVehicles.CurrentRow.Cells[1].Value.ToString())
                    .AddQueryParameter("from_date", from_time).AddQueryParameter("to_date", to_time)
                    .AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                //sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }


            int row = dgvVehicles.CurrentRow.Index;
            string fileName = "C:\\Report_DAT\\" + dgvVehicles[1, row].Value.ToString() + ".xls";
            CreatFileExcelReportVehicle(fileName, dgvVehicles.CurrentRow.Cells[1].Value.ToString(), dgvVehicles.CurrentRow.Cells[3].Value.ToString(), dgvVehicles.CurrentRow.Cells[4].Value.ToString(), dgvVehicles.CurrentRow.Cells[4].Value.ToString(), Khoangthoigian, Sessions);
            dgvSessions.Rows.Clear();
            DateTime StartTime;
            string ViPham = "";
            Boolean GetAll = true;
            if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                GetAll = true;
            else
                GetAll = false;
            foreach (SessionRes session in Sessions)
            {
                if ((GetAll == false) && (chkNonCheck.Checked == false) && (session.sync_status == 0))
                    continue;
                if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (session.sync_status < 0))
                    continue;
                if ((GetAll == false) && (chkCheckOk.Checked == false) && (session.sync_status > 0))
                    continue;

                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                if (session.sync_status == 0)
                    ViPham = "Chưa kiểm tra";
                else if (session.sync_status > 0)
                    ViPham = "Không vi phạm";
                else
                    ViPham = session.sync_error;
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
            }
            OpenMyExcelFile(fileName);
        }

        private void inBáoCáoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string monthfrom, dayfrom, monthto, dayto;
            monthfrom = dtpFrom.Value.Month.ToString();
            if (monthfrom.Length < 2) monthfrom = "0" + monthfrom;
            dayfrom = dtpFrom.Value.Day.ToString();
            if (dayfrom.Length < 2) dayfrom = "0" + dayfrom;
            monthto = dtpTo.Value.Month.ToString();
            if (monthto.Length < 2) monthto = "0" + monthto;
            dayto = dtpTo.Value.Day.ToString();
            if (dayto.Length < 2) dayto = "0" + dayto;


            string from_time = dtpFrom.Value.Year.ToString() + "-" + monthfrom + "-" + dayfrom + "T00:00:00";
            string to_time = dtpTo.Value.Year.ToString() + "-" + monthto + "-" + dayto + "T23:59:59";
            string Khoangthoigian = "Từ ngày " + dayfrom + "/" + monthfrom + "/" + dtpFrom.Value.Year.ToString() + " đến ngày " + dayto + "/" + monthto + "/" + dtpTo.Value.Year.ToString();
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("plate", dgvVehicles.CurrentRow.Cells[1].Value.ToString())
                    .AddQueryParameter("from_date", from_time).AddQueryParameter("to_date", to_time)
                    .AddParameter("page_size", 500).AddParameter("synced", 1);
                var response2 = client.Get<ResultSessionRes>(request2);
                //sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("plate", dgvVehicles.CurrentRow.Cells[1].Value.ToString())
                    .AddQueryParameter("from_date", from_time).AddQueryParameter("to_date", to_time)
                    .AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                //sTraineeID = dgwTrainees.CurrentRow.Cells[1].Value.ToString();
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }


            int row = dgvVehicles.CurrentRow.Index;
            string fileName = "C:\\Report_DAT\\" + dgvVehicles[1, row].Value.ToString() + ".xls";
            CreatFileExcelReportVehicle(fileName, dgvVehicles.CurrentRow.Cells[1].Value.ToString(), dgvVehicles.CurrentRow.Cells[3].Value.ToString(), dgvVehicles.CurrentRow.Cells[4].Value.ToString(), dgvVehicles.CurrentRow.Cells[4].Value.ToString(), Khoangthoigian, Sessions);
            dgvSessions.Rows.Clear();
            DateTime StartTime;
            foreach (SessionRes session in Sessions)
            {
                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString());
            }

            if (System.IO.File.Exists(fileName))
            {
                PrintMyExcelFile(fileName);
            }

        }

        private void xemDanhSáchHọcViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (accessToken == "")
            {
                return;
            }
            int count = 1;
            int SoHV = Int32.Parse(dgvCoures.CurrentRow.Cells[4].Value.ToString());
            TenKhoaHoc = dgvCoures.CurrentRow.Cells[1].Value.ToString() + " ( " + dgvCoures.CurrentRow.Cells[2].Value.ToString() + " )";
            HangDaoTao = dgvCoures.CurrentRow.Cells[3].Value.ToString();
            dgwTrainees.Rows.Clear();
            IRestRequest request;
            IRestResponse<ResultTraineeRes> response;
            List<TraineeRes> lisTrainees = new List<TraineeRes>();
            if (SoHV <= 50)
            {
                request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 50);
                response = client.Get<ResultTraineeRes>(request);

                if (chkbSortID.Checked)
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_distance).ToList();

            }
            else
            {
                int Page = SoHV / 50 + 1;
                for (int k = 0; k < Page; k++)
                {
                    request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page", k + 1).AddParameter("page_size", 50);
                    response = client.Get<ResultTraineeRes>(request);

                    lisTrainees.AddRange(response.Data.items.ToList());

                }
                if (chkbSortID.Checked)
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_distance).ToList();
            }

            foreach (TraineeRes trainee in lisTrainees)
            {
                count++;
                dgwTrainees.Rows.Add((count - 1).ToString(), trainee.id.ToString(), trainee.ho_va_ten, trainee.ngay_sinh, trainee.synced_outdoor_hours.ToString(), trainee.synced_outdoor_distance.ToString(),
                    (trainee.outdoor_hour / 3600).ToString(), (trainee.outdoor_distance / 1000).ToString(), trainee.outdoor_session_count.ToString(), trainee.ma_dk, trainee.anh_chan_dung);
            }
        }

        private void dgvCoures_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo ht = dgvCoures.HitTest(e.X, e.Y);
            if ((ht.ColumnIndex >= 0) && (ht.RowIndex >= 0))
            {
                dgvCoures.CurrentCell = dgvCoures.Rows[ht.RowIndex].Cells[ht.ColumnIndex];
                dgvCoures.ContextMenuStrip = MenuGetTrainees;
            }
            else
            {
                dgvCoures.ContextMenuStrip = null;
            }

        }

        private void CreatFileExcelReportCouse(string fileName, string MaKhoaHoc, string HangDT, string NgayKG, string NgayBG, string CSDT, List<TraineeRes> LstTmp)
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

                //worksheet.PageSetup.TopMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.LeftMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.RightMargin = Convert.ToDouble(36);

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
                int idx;
                string st = DAT_ToolReports.Properties.Settings.Default.Company.ToUpper();
                int nPos = st.IndexOf(DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper());
                if (nPos > 0)
                    st = st.Substring(0, nPos).Trim();

                Cells.Merge(0, 0, 1, 5);
                Cells["A1"].Value = st;
                Cells["A1"].SetStyle(style1);
                Cells.Merge(1, 0, 1, 5);
                Cells["A2"].Value = DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper();
                Cells["A2"].SetStyle(style1);
                Cells.Merge(2, 0, 1, 5);
                Cells["A3"].Value = "***********";
                Cells["A3"].SetStyle(styleDate);

                Cells.Merge(0, 5, 1, 4);
                Cells["F1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                Cells["F1"].SetStyle(style1);
                Cells.Merge(1, 5, 1, 4);
                Cells["F2"].Value = "Độc lập - Tự do - Hạnh phúc";
                Cells["F2"].SetStyle(styleDate);
                Cells.Merge(2, 5, 1, 4);
                Cells["F3"].Value = "***********";
                Cells["F3"].SetStyle(styleDate);

                Aspose.Cells.Style style5;
                style5 = Cells["A1"].GetStyle();
                style5.Font.Size = 13;
                style5.Font.Name = "Times New Roman";
                style5.Font.IsBold = false;
                //style5.ShrinkToFit = true;
                style5.HorizontalAlignment = TextAlignmentType.Left;

                Aspose.Cells.Style style7;
                style7 = Cells["J4"].GetStyle();
                style7.VerticalAlignment = TextAlignmentType.Center;
                style7.HorizontalAlignment = TextAlignmentType.Center;
                style7.Font.Size = 13;
                style7.Font.Name = "Times New Roman";
                style7.Font.IsItalic = true;
                Cells.Merge(3, 5, 1, 4);
                Cells["F4"].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells["F4"].SetStyle(style7);

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(5, 0, 1, 9);
                Cells["A6"].Value = "BÁO CÁO KẾT QUẢ THỰC HÀNH LÁI XE CỦA KHÓA HỌC";
                Cells["A6"].SetStyle(styleTitle);

                Aspose.Cells.Style styleKhoaThi;
                styleKhoaThi = Cells["A7"].GetStyle();
                styleKhoaThi.VerticalAlignment = TextAlignmentType.Center;
                styleKhoaThi.HorizontalAlignment = TextAlignmentType.Center;
                styleKhoaThi.Font.Color = System.Drawing.Color.Black;
                styleKhoaThi.Font.IsBold = true;
                styleKhoaThi.Font.Size = 13;
                styleKhoaThi.Font.Name = "Times New Roman";
                Cells.Merge(6, 0, 1, 9);
                Cells["A7"].Value = "(Ngày báo cáo: " + "ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString() + ")";
                Cells["A7"].SetStyle(styleKhoaThi);

                //========================================
                Cells.Merge(8, 0, 1, 9);
                Cells[8, 0].Value = "I. Thông tin khóa học";
                Cells[8, 0].SetStyle(styleKhoaThi);

                //Cells.Merge(10, 7, 6, 2);
                ////fileNameImage = "C:\\Img_gen\\521316.jpg";// + TraineeNumber.ToString() + ".xls";
                //idx = worksheet.Pictures.Add(10, 7, fileNameImage);
                //Picture pic = worksheet.Pictures[idx];
                //double w = worksheet.Cells.GetColumnWidthInch(7) + worksheet.Cells.GetColumnWidthInch(8);
                //double h = worksheet.Cells.GetRowHeightInch(10) * 6; //6 dòng
                //pic.WidthInch = w;
                //pic.HeightInch = h;

                int row = 10;

                Cells[row, 0].Value = "Mã khóa học:";
                Cells[row, 0].SetStyle(style5);
                Cells[row, 2].Value = MaKhoaHoc;
                Cells[row, 2].SetStyle(style5);
                Cells[row + 1, 0].Value = "Hạng đào tạo:";
                Cells[row + 1, 0].SetStyle(style5);
                Cells[row + 1, 2].Value = HangDT;
                Cells[row + 1, 2].SetStyle(style5);
                Cells[row + 2, 0].Value = "Ngày khai giảng:";
                Cells[row + 2, 0].SetStyle(style5);
                Cells[row + 2, 2].Value = NgayKG.Substring(8, 2) + "/" + NgayKG.Substring(5, 2) + "/" + NgayKG.Substring(0, 4);
                Cells[row + 2, 2].SetStyle(style5);
                Cells[row + 3, 0].Value = "Ngày bế giảng:";
                Cells[row + 3, 0].SetStyle(style5);
                Cells[row + 3, 2].Value = NgayBG.Substring(8, 2) + "/" + NgayBG.Substring(5, 2) + "/" + NgayBG.Substring(0, 4);
                Cells[row + 3, 2].SetStyle(style5);
                Cells[row + 4, 0].Value = "Cơ sở đào tạo:";
                Cells[row + 4, 0].SetStyle(style5);
                Cells[row + 4, 2].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                Cells[row + 4, 2].SetStyle(style5);

                //Cells[row + 5, 0].Value = "Cơ sở đào tạo:";
                //Cells[row + 5, 0].SetStyle(style5);
                //Cells[row + 5, 2].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                //Cells[row + 5, 2].SetStyle(style5);

                Cells.Merge(row + 7, 0, 1, 9);
                Cells[row + 7, 0].Value = "II. Thông tin quá trình đào tạo";
                Cells[row + 7, 0].SetStyle(styleKhoaThi);

                Aspose.Cells.Range range = Cells.CreateRange(10, 0, 5, 9);
                range.SetOutlineBorders(CellBorderType.Thin, System.Drawing.Color.Black);

                row = 18;

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
                Cells.Merge(row, 1, 1, 2);
                Cells.Merge(row, 4, 1, 2);
                //Cells.Merge(row, 5, 1, 2);
                //Cells.Merge(row, 7, 1, 2);
                Cells[row, 0].Value = "STT";
                Cells[row, 1].Value = "Họ và tên";
                Cells[row, 3].Value = "Ngày sinh";
                Cells[row, 4].Value = "Mã học viên";
                Cells[row, 6].Value = "Thời gian đào tạo";
                Cells[row, 7].Value = "Quãng đường đào tạo";
                Cells[row, 8].Value = "Ghi chú";

                Cells[row, 0].SetStyle(styleHeader);
                Cells[row, 1].SetStyle(styleHeader);
                Cells[row, 2].SetStyle(styleHeader);
                Cells[row, 3].SetStyle(styleHeader);
                Cells[row, 4].SetStyle(styleHeader);
                Cells[row, 5].SetStyle(styleHeader);
                Cells[row, 6].SetStyle(styleHeader);
                Cells[row, 7].SetStyle(styleHeader);
                Cells[row, 8].SetStyle(styleHeader);

                Aspose.Cells.Style style4;
                style4 = Cells[row + 1, 0].GetStyle();
                style4.Font.Size = 13;
                style4.Font.Name = "Times New Roman";
                style4.ShrinkToFit = true;
                style4.HorizontalAlignment = TextAlignmentType.Center;
                style4.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Aspose.Cells.Style style8;
                style8 = Cells[row + 1, 0].GetStyle();
                style8.Font.Size = 13;
                style8.Font.Name = "Times New Roman";
                style8.ShrinkToFit = true;
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
                int Offset = 1;
                string CheckinH, CheckinM, CheckoutH, CheckoutM, TimelearnH, TimelearnM, TongTGH, TongTGM;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    TraineeRes dateAttendanceTmp = LstTmp[i];
                    Cells[row, 0].Value = Offset.ToString();
                    Cells[row, 1].Value = LstTmp[i].ho_va_ten;
                    Cells[row, 3].Value = LstTmp[i].ngay_sinh.Substring(8, 2) + "/" + LstTmp[i].ngay_sinh.Substring(5, 2) + "/" + LstTmp[i].ngay_sinh.Substring(0, 4);
                    Cells[row, 4].Value = LstTmp[i].ma_dk;
                    Cells[row, 6].Value = ((LstTmp[i].outdoor_hour / 60) / 60).ToString() + ":" + ((LstTmp[i].outdoor_hour / 60) % 60).ToString();
                    Cells[row, 7].Value = ((double)LstTmp[i].outdoor_distance / 1000).ToString();
                    Cells[row, 8].Value = " ";


                    Cells.Merge(row, 1, 1, 2);
                    Cells.Merge(row, 4, 1, 2);
                    //Cells.Merge(row, 5, 1, 2);
                    //Cells.Merge(row, 7, 1, 2);
                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style8);
                    Cells[row, 2].SetStyle(style8);
                    Cells[row, 3].SetStyle(style8);
                    Cells[row, 4].SetStyle(style8);
                    Cells[row, 5].SetStyle(style4);
                    Cells[row, 6].SetStyle(style4);
                    Cells[row, 7].SetStyle(style4);
                    Cells[row, 8].SetStyle(style4);
                    row++;
                    Offset++;
                }


                row++;

                //Range range = Cells.CreateRange(1 + 1, 0, row - 1 - 2, 9);
                //range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

                Cells.Merge(row, 0, 1, 9);
                //Cells[row, 0].Value = LstTmp.Count.ToString();
                //Cells[row, 0].SetStyle(styleSum);
                row++;


                Aspose.Cells.Style style6;
                style6 = Cells[row + 2, 0].GetStyle();
                style6.Font.IsBold = true;
                style6.Font.Size = 13;
                style6.Font.Name = "Times New Roman";
                style6.HorizontalAlignment = TextAlignmentType.Center;

                Cells[row + 2, 0].Value = "  ";
                Cells[row + 3, 0].Value = "  ";
                Cells.Merge(row + 2, 0, 1, 4);
                Cells.Merge(row + 3, 0, 1, 4);
                Cells[row + 2, 0].SetStyle(style6);
                Cells[row + 3, 0].SetStyle(style7);

                Cells.Merge(row + 1, 6, 1, 3);
                Cells[row + 1, 6].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells[row + 1, 6].SetStyle(style7);

                Cells[row + 2, 6].Value = "Trưởng phòng đào tạo";
                Cells[row + 3, 6].Value = "(ký tên)";
                Cells.Merge(row + 2, 6, 1, 3);
                Cells.Merge(row + 3, 6, 1, 3);
                Cells[row + 2, 6].SetStyle(style6);
                Cells[row + 3, 6].SetStyle(style7);

                Cells.SetColumnWidthPixel(0, 100);
                Cells.SetColumnWidthPixel(1, 180);
                Cells.SetColumnWidthPixel(2, 120);
                Cells.SetColumnWidthPixel(3, 100);
                //worksheet.AutoFitColumn(3);
                Cells.SetColumnWidthPixel(4, 100);
                Cells.SetColumnWidthPixel(5, 120);
                Cells.SetColumnWidthPixel(6, 140);
                Cells.SetColumnWidthPixel(7, 140);
                Cells.SetColumnWidthPixel(8, 90);

                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }

        private void CreatFileExcelReportCouse_FromExcel(string fileName, string sThongTinFile, List<TraineeRes> LstTmp)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                object misValue = System.Reflection.Missing.Value;

                worksheet.PageSetup.Orientation = PageOrientationType.Landscape;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0;
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.TopMargin = 1;
                worksheet.PageSetup.BottomMargin = 1;
                worksheet.PageSetup.LeftMargin = 1;
                worksheet.PageSetup.RightMargin = 0.3;

                //worksheet.PageSetup.TopMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.LeftMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.RightMargin = Convert.ToDouble(36);

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
                int idx;
                string st = DAT_ToolReports.Properties.Settings.Default.Company.ToUpper();
                int nPos = st.IndexOf(DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper());
                if (nPos > 0)
                    st = st.Substring(0, nPos).Trim();

                Cells.Merge(0, 0, 1, 6);
                Cells["A1"].Value = st;
                Cells["A1"].SetStyle(style1);
                Cells.Merge(1, 0, 1, 6);
                Cells["A2"].Value = DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper();
                Cells["A2"].SetStyle(style1);
                Cells.Merge(2, 0, 1, 6);
                Cells["A3"].Value = "***********";
                Cells["A3"].SetStyle(styleDate);

                Cells.Merge(0, 6, 1, 5);
                Cells["G1"].Value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
                Cells["G1"].SetStyle(style1);
                Cells.Merge(1, 6, 1, 5);
                Cells["G2"].Value = "Độc lập - Tự do - Hạnh phúc";
                Cells["G2"].SetStyle(styleDate);
                Cells.Merge(2, 6, 1, 5);
                Cells["G3"].Value = "***********";
                Cells["G3"].SetStyle(styleDate);

                Aspose.Cells.Style style5;
                style5 = Cells["A1"].GetStyle();
                style5.Font.Size = 13;
                style5.Font.Name = "Times New Roman";
                style5.Font.IsBold = false;
                //style5.ShrinkToFit = true;
                style5.HorizontalAlignment = TextAlignmentType.Left;

                Aspose.Cells.Style style7;
                style7 = Cells["J4"].GetStyle();
                style7.VerticalAlignment = TextAlignmentType.Center;
                style7.HorizontalAlignment = TextAlignmentType.Center;
                style7.Font.Size = 13;
                style7.Font.Name = "Times New Roman";
                style7.Font.IsItalic = true;
                Cells.Merge(3, 6, 1, 5);
                Cells["G4"].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                Cells["G4"].SetStyle(style7);

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(5, 0, 1, 11);
                Cells["A6"].Value = "BÁO CÁO KẾT QUẢ ĐÀO TẠO THỰC HÀNH LÁI XE TRÊN ĐƯỜNG GIAO THÔNG";
                Cells["A6"].SetStyle(styleTitle);

                Aspose.Cells.Style styleKhoaThi;
                styleKhoaThi = Cells["A7"].GetStyle();
                styleKhoaThi.VerticalAlignment = TextAlignmentType.Center;
                styleKhoaThi.HorizontalAlignment = TextAlignmentType.Center;
                styleKhoaThi.Font.Color = System.Drawing.Color.Black;
                styleKhoaThi.Font.IsBold = true;
                styleKhoaThi.Font.Size = 13;
                styleKhoaThi.Font.Name = "Times New Roman";
                Cells.Merge(6, 0, 1, 11);
                Cells["A7"].Value = "(Ngày báo cáo: " + "ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString() + ")";
                Cells["A7"].SetStyle(styleKhoaThi);

                //========================================
                Cells.Merge(8, 0, 1, 11);
                Cells[8, 0].Value = "THEO DANH SÁCH THÍ SINH SÁT HẠCH: " + sThongTinFile;
                Cells[8, 0].SetStyle(styleKhoaThi);



                int row = 10;


                Cells.Merge(row + 1, 0, 1, 11);
                Cells[row + 1, 0].Value = "II. Thông tin quá trình đào tạo";
                Cells[row + 1, 0].SetStyle(styleKhoaThi);

                Aspose.Cells.Range range = Cells.CreateRange(10, 0, 5, 11);
                range.SetOutlineBorders(CellBorderType.Thin, System.Drawing.Color.Black);

                row = 12;

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

                Cells.Merge(row, 1, 1, 2);
                Cells.Merge(row, 3, 1, 2);

                Cells[row, 0].Value = "STT";
                Cells[row, 1].Value = "Mã học viên";
                Cells[row, 3].Value = "Họ và tên";
                Cells[row, 5].Value = "Ngày sinh";
                Cells[row, 6].Value = "Hạng";
                Cells[row, 7].Value = "Thời gian đào tạo";
                Cells[row, 8].Value = "Quãng đường đào tạo";
                Cells[row, 9].Value = "Thời gian học số tự động";
                Cells[row, 10].Value = "Thời gian học ban đêm";

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

                Aspose.Cells.Style style4;
                style4 = Cells[row + 1, 0].GetStyle();
                style4.Font.Size = 13;
                style4.Font.Name = "Times New Roman";
                style4.ShrinkToFit = true;
                style4.HorizontalAlignment = TextAlignmentType.Center;
                style4.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);

                Aspose.Cells.Style style8;
                style8 = Cells[row + 1, 0].GetStyle();
                style8.Font.Size = 13;
                style8.Font.Name = "Times New Roman";
                style8.ShrinkToFit = true;
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
                int Offset = 1;
                string CheckinH, CheckinM, CheckoutH, CheckoutM, TimelearnH, TimelearnM, TongTGH, TongTGM;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    TraineeRes dateAttendanceTmp = LstTmp[i];
                    Cells[row, 0].Value = Offset.ToString();
                    Cells[row, 1].Value = LstTmp[i].ma_dk;
                    Cells[row, 3].Value = LstTmp[i].ho_va_ten;
                    Cells[row, 5].Value = LstTmp[i].ngay_sinh.Substring(8, 2) + "/" + LstTmp[i].ngay_sinh.Substring(5, 2) + "/" + LstTmp[i].ngay_sinh.Substring(0, 4);
                    Cells[row, 6].Value = LstTmp[i].hang_daotao;
                    Cells[row, 7].Value = ((LstTmp[i].outdoor_hour / 60) / 60).ToString() + ":" + ((LstTmp[i].outdoor_hour / 60) % 60).ToString();
                    Cells[row, 8].Value = ((double)LstTmp[i].outdoor_distance / 1000).ToString();
                    Cells[row, 9].Value = ((LstTmp[i].auto_duration / 60) / 60).ToString() + ":" + ((LstTmp[i].auto_duration / 60) % 60).ToString();
                    Cells[row, 10].Value = ((LstTmp[i].night_duration / 60) / 60).ToString() + ":" + ((LstTmp[i].night_duration / 60) % 60).ToString();


                    Cells.Merge(row, 1, 1, 2);
                    Cells.Merge(row, 3, 1, 2);

                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style8);
                    Cells[row, 2].SetStyle(style8);
                    Cells[row, 3].SetStyle(style8);
                    Cells[row, 4].SetStyle(style8);
                    Cells[row, 5].SetStyle(style4);
                    Cells[row, 6].SetStyle(style4);
                    Cells[row, 7].SetStyle(style4);
                    Cells[row, 8].SetStyle(style4);
                    Cells[row, 9].SetStyle(style4);
                    Cells[row, 10].SetStyle(style4);
                    row++;
                    Offset++;
                }


                row++;

                //Range range = Cells.CreateRange(1 + 1, 0, row - 1 - 2, 9);
                //range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

                Cells.Merge(row, 0, 1, 11);
                //Cells[row, 0].Value = LstTmp.Count.ToString();
                //Cells[row, 0].SetStyle(styleSum);
                row++;


                Aspose.Cells.Style style6;
                style6 = Cells[row + 2, 0].GetStyle();
                style6.Font.IsBold = true;
                style6.Font.Size = 13;
                style6.Font.Name = "Times New Roman";
                style6.HorizontalAlignment = TextAlignmentType.Center;

                Cells[row + 2, 0].Value = "  ";
                Cells[row + 3, 0].Value = "  ";
                Cells.Merge(row + 2, 0, 1, 4);
                Cells.Merge(row + 3, 0, 1, 4);
                Cells[row + 2, 0].SetStyle(style6);
                Cells[row + 3, 0].SetStyle(style7);

                //Cells.Merge(row + 1, 6, 1, 3);
                //Cells[row + 1, 6].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                //Cells[row + 1, 6].SetStyle(style7);

                Cells[row + 1, 6].Value = "Trưởng phòng đào tạo";
                Cells[row + 2, 6].Value = "(ký tên)";
                Cells.Merge(row + 1, 6, 1, 3);
                Cells.Merge(row + 2, 6, 1, 3);
                Cells[row + 1, 6].SetStyle(style6);
                Cells[row + 2, 6].SetStyle(style7);

                Cells.SetColumnWidthPixel(0, 50);
                Cells.SetColumnWidthPixel(1, 160);
                Cells.SetColumnWidthPixel(2, 120);
                Cells.SetColumnWidthPixel(3, 150);
                //worksheet.AutoFitColumn(3);
                Cells.SetColumnWidthPixel(4, 130);
                Cells.SetColumnWidthPixel(5, 120);
                Cells.SetColumnWidthPixel(6, 80);
                Cells.SetColumnWidthPixel(7, 140);
                Cells.SetColumnWidthPixel(8, 140);
                Cells.SetColumnWidthPixel(9, 140);
                Cells.SetColumnWidthPixel(10, 140);

                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }

        private void CreatFileExcelReportSession_FromExcel(string fileName, string sThongTinFile, List<InforSessionReport> LstTmp)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                object misValue = System.Reflection.Missing.Value;

                worksheet.PageSetup.Orientation = PageOrientationType.Landscape;
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0;
                worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                worksheet.PageSetup.TopMargin = 1;
                worksheet.PageSetup.BottomMargin = 1;
                worksheet.PageSetup.LeftMargin = 1;
                worksheet.PageSetup.RightMargin = 0.3;

                //worksheet.PageSetup.TopMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.LeftMargin = Convert.ToDouble(36);
                //worksheet.PageSetup.RightMargin = Convert.ToDouble(36);

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
                //style1.Number = 3;


                Aspose.Cells.Style styleDate;
                styleDate = Cells["A1"].GetStyle();
                styleDate.VerticalAlignment = TextAlignmentType.Center;
                styleDate.HorizontalAlignment = TextAlignmentType.Center;
                styleDate.Font.Color = System.Drawing.Color.Black;
                styleDate.Font.IsBold = true;
                styleDate.Font.IsItalic = true;
                styleDate.Font.Size = 13;
                styleDate.Font.Name = "Times New Roman";
                int idx;
                string st = DAT_ToolReports.Properties.Settings.Default.Company.ToUpper();
                int nPos = st.IndexOf(DAT_ToolReports.Properties.Settings.Default.Centre.ToUpper());

                

                Aspose.Cells.Style styleTitle;
                styleTitle = Cells["F1"].GetStyle();
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
                styleTitle.HorizontalAlignment = TextAlignmentType.Center;
                styleTitle.Font.Color = System.Drawing.Color.Black;
                styleTitle.Font.IsBold = true;
                styleTitle.Font.IsItalic = false;
                styleTitle.Font.Size = 15;
                styleTitle.Font.Name = "Times New Roman";

                Cells.Merge(1, 0, 1, 11);
                Cells["A1"].Value = "BÁO CÁO KẾT QUẢ ĐÀO TẠO THỰC HÀNH LÁI XE TRÊN ĐƯỜNG GIAO THÔNG";
                Cells["A1"].SetStyle(styleTitle);

                Aspose.Cells.Style styleKhoaThi;
                styleKhoaThi = Cells["A7"].GetStyle();
                styleKhoaThi.VerticalAlignment = TextAlignmentType.Center;
                styleKhoaThi.HorizontalAlignment = TextAlignmentType.Center;
                styleKhoaThi.Font.Color = System.Drawing.Color.Black;
                styleKhoaThi.Font.IsBold = true;
                styleKhoaThi.Font.Size = 13;
                styleKhoaThi.Font.Name = "Times New Roman";
                Cells.Merge(3, 0, 1, 11);
                Cells["A3"].Value = "(Ngày báo cáo: " + "ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString() + ")";
                Cells["A3"].SetStyle(styleKhoaThi);

                //========================================
                Cells.Merge(5, 0, 1, 11);
                Cells[5, 0].Value = "THEO DANH SÁCH THÍ SINH SÁT HẠCH: " + sThongTinFile;
                Cells[5, 0].SetStyle(styleKhoaThi);



                int row = 7;




                //Aspose.Cells.Range range = Cells.CreateRange(10, 0, 5, 11);
                //range.SetOutlineBorders(CellBorderType.Thin, System.Drawing.Color.Black);

                //row = 12;

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

                //Cells.Merge(row, 1, 1, 2);
                //Cells.Merge(row, 3, 1, 2);

                Cells[row, 0].Value = "Mã phiên học";
                Cells[row, 1].Value = "Thời gian bắt đầu phiên học";
                Cells[row, 2].Value = "Thời gian kết thúc phiên học";
                Cells[row, 3].Value = "Thời gian thực hành";
                Cells[row, 4].Value = "Quãng đường thực hành";
                Cells[row, 5].Value = "Mã học viên";
                Cells[row, 6].Value = "Họ và tên học viên";
                Cells[row, 7].Value = "Mã khóa học";
                Cells[row, 8].Value = "Tên khóa học";
                Cells[row, 9].Value = "Loại khóa học";
                Cells[row, 10].Value = "Biển số xe";
                Cells[row, 11].Value = "Hạng xe tập lái";
                Cells[row, 12].Value = "Sở GTVT Quản lý";
                Cells[row, 13].Value = "Cơ sở đào tạo";
                Cells[row, 14].Value = "Đơn vị truyền dữ liệu";

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
                Cells[row, 14].SetStyle(styleHeader);
                
                Aspose.Cells.Style style4;
                style4 = Cells[row + 1, 0].GetStyle();
                style4.Font.Size = 13;
                style4.Font.Name = "Times New Roman";
                style4.ShrinkToFit = true;
                style4.HorizontalAlignment = TextAlignmentType.Center;
                style4.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
                style4.Number = 4;

                Aspose.Cells.Style style8;
                style8 = Cells[row + 1, 0].GetStyle();
                style8.Font.Size = 13;
                style8.Font.Name = "Times New Roman";
                style8.ShrinkToFit = true;
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
                int Offset = 1;
                string CheckinH, CheckinM, CheckoutH, CheckoutM, TimelearnH, TimelearnM, TongTGH, TongTGM;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    InforSessionReport dateAttendanceTmp = LstTmp[i];
                    Cells[row, 3].Value = LstTmp[i].ThoiGianTH;
                    Cells[row, 4].Value = LstTmp[i].QuangDuongTH;
                    
                    Cells[row, 3].SetStyle(style4);
                    Cells[row, 4].SetStyle(style4);
                    row++;
                    Offset++;
                }
                workbook.Worksheets[0].Cells.ConvertStringToNumericValue();

                row = 8;
                for (int i = 0; i < LstTmp.Count; i++)
                {
                    InforSessionReport dateAttendanceTmp = LstTmp[i];
                    Cells[row, 0].Value = LstTmp[i].MaPhienHoc;
                    Cells[row, 1].Value = LstTmp[i].StartTime;
                    Cells[row, 2].Value = LstTmp[i].StopTime;
                    //Cells[row, 3].Value = LstTmp[i].ThoiGianTH;
                    //Cells[row, 4].Value = LstTmp[i].QuangDuongTH;
                    Cells[row, 5].Value = LstTmp[i].MaHocVien;
                    Cells[row, 6].Value = LstTmp[i].HoTenHocVien;
                    Cells[row, 7].Value = LstTmp[i].MaKhoaHoc;
                    Cells[row, 8].Value = LstTmp[i].TenKhoaHoc;
                    Cells[row, 9].Value = LstTmp[i].LoaiKhoaHoc;
                    Cells[row, 10].Value = LstTmp[i].BienSoXe;
                    Cells[row, 11].Value = LstTmp[i].HangXeTL;
                    Cells[row, 12].Value = DAT_ToolReports.Properties.Settings.Default.Company;
                    Cells[row, 13].Value = DAT_ToolReports.Properties.Settings.Default.Centre;
                    Cells[row, 14].Value = "Công ty CP Công Nghệ Sát Hạch Toàn Phương";


                    Cells[row, 0].SetStyle(style4);
                    Cells[row, 1].SetStyle(style4);
                    Cells[row, 2].SetStyle(style4);
                    //Cells[row, 3].SetStyle(style4);
                    //Cells[row, 4].SetStyle(style4);
                    Cells[row, 5].SetStyle(style4);
                    Cells[row, 6].SetStyle(style4);
                    Cells[row, 7].SetStyle(style4);
                    Cells[row, 8].SetStyle(style4);
                    Cells[row, 9].SetStyle(style4);
                    Cells[row, 10].SetStyle(style4);
                    Cells[row, 11].SetStyle(style4);
                    Cells[row, 12].SetStyle(style4);
                    Cells[row, 13].SetStyle(style4);
                    Cells[row, 14].SetStyle(style4);

                    row++;
                    Offset++;
                }

                row++;

                //Range range = Cells.CreateRange(1 + 1, 0, row - 1 - 2, 9);
                //range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

                Cells.Merge(row, 0, 1, 11);
                //Cells[row, 0].Value = LstTmp.Count.ToString();
                //Cells[row, 0].SetStyle(styleSum);
                row++;


                Aspose.Cells.Style style6;
                style6 = Cells[row + 2, 0].GetStyle();
                style6.Font.IsBold = true;
                style6.Font.Size = 13;
                style6.Font.Name = "Times New Roman";
                style6.HorizontalAlignment = TextAlignmentType.Center;

                Cells[row + 2, 0].Value = "  ";
                Cells[row + 3, 0].Value = "  ";
                Cells.Merge(row + 2, 0, 1, 4);
                Cells.Merge(row + 3, 0, 1, 4);
                Cells[row + 2, 0].SetStyle(style6);
                Cells[row + 3, 0].SetStyle(style6);

                //Cells.Merge(row + 1, 6, 1, 3);
                //Cells[row + 1, 6].Value = DAT_ToolReports.Properties.Settings.Default.Province + ", ngày " + DateTime.Now.Day.ToString() + " tháng " + DateTime.Now.Month.ToString() + " năm " + DateTime.Now.Year.ToString();
                //Cells[row + 1, 6].SetStyle(style7);

                Cells[row + 1, 6].Value = "Trưởng phòng đào tạo";
                Cells[row + 2, 6].Value = "(ký tên)";
                Cells.Merge(row + 1, 6, 1, 3);
                Cells.Merge(row + 2, 6, 1, 3);
                Cells[row + 1, 6].SetStyle(style6);
                Cells[row + 2, 6].SetStyle(style6);

                Cells.SetColumnWidthPixel(0, 350);
                Cells.SetColumnWidthPixel(1, 160);
                Cells.SetColumnWidthPixel(2, 160);
                Cells.SetColumnWidthPixel(3, 80);
                Cells.SetColumnWidthPixel(4, 80);
                Cells.SetColumnWidthPixel(5, 200);
                Cells.SetColumnWidthPixel(6, 200);
                Cells.SetColumnWidthPixel(7, 140);
                Cells.SetColumnWidthPixel(8, 90);
                Cells.SetColumnWidthPixel(9, 90);
                Cells.SetColumnWidthPixel(10, 90);
                Cells.SetColumnWidthPixel(11, 90);
                Cells.SetColumnWidthPixel(12, 240);
                Cells.SetColumnWidthPixel(13, 240);
                Cells.SetColumnWidthPixel(14, 240);

                workbook.Save(fileName);

                MessageBox.Show("File " + fileName + " đã được tạo ra thành công", "Thông báo");
            }
            catch (SystemException se)
            {
                MessageBox.Show("Lỗi xuất file XLS.\n" + se.Message, "Thông báo");
            }
            Cursor.Current = Cursors.Default;
        }

        private void xuấtDanhSáchRaFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (accessToken == "")
            {
                return;
            }
            int count = 1;
            int SoHV = Int32.Parse(dgvCoures.CurrentRow.Cells[4].Value.ToString());
            TenKhoaHoc = dgvCoures.CurrentRow.Cells[2].Value.ToString();
            HangDaoTao = dgvCoures.CurrentRow.Cells[3].Value.ToString();
            dgwTrainees.Rows.Clear();
            IRestRequest request;
            IRestResponse<ResultTraineeRes> response;
            List<TraineeRes> lisTrainees = new List<TraineeRes>();
            if (SoHV <= 50)
            {
                request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 50);
                response = client.Get<ResultTraineeRes>(request);

                if (chkbSortID.Checked)
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_distance).ToList();

            }
            else
            {
                int Page = SoHV / 50 + 1;
                for (int k = 0; k < Page; k++)
                {
                    request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page", k + 1).AddParameter("page_size", 50);
                    response = client.Get<ResultTraineeRes>(request);

                    lisTrainees.AddRange(response.Data.items.ToList());

                }
                if (chkbSortID.Checked)
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_distance).ToList();
            }

            foreach (TraineeRes trainee in lisTrainees)
            {
                count++;
                dgwTrainees.Rows.Add((count - 1).ToString(), trainee.id.ToString(), trainee.ho_va_ten, trainee.ngay_sinh, trainee.synced_outdoor_hours.ToString(), trainee.synced_outdoor_distance.ToString(),
                    (trainee.outdoor_hour / 3600).ToString(), (trainee.outdoor_distance / 1000).ToString(), trainee.outdoor_session_count.ToString(), trainee.ma_dk, trainee.anh_chan_dung);
            }



            int row = dgvCoures.CurrentRow.Index;
            string fileName = "C:\\Report_DAT\\" + dgvCoures[1, row].Value.ToString() + ".xls";
            CreatFileExcelReportCouse(fileName, dgvCoures.CurrentRow.Cells[1].Value.ToString(), dgvCoures.CurrentRow.Cells[3].Value.ToString(), dgvCoures.CurrentRow.Cells[5].Value.ToString(), dgvCoures.CurrentRow.Cells[6].Value.ToString(), "Ko dung den", lisTrainees);
            OpenMyExcelFile(fileName);

        }

        private void inDanhSáchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (accessToken == "")
            {
                return;
            }
            int count = 1;
            int SoHV = Int32.Parse(dgvCoures.CurrentRow.Cells[4].Value.ToString());
            TenKhoaHoc = dgvCoures.CurrentRow.Cells[2].Value.ToString();
            HangDaoTao = dgvCoures.CurrentRow.Cells[3].Value.ToString();
            dgwTrainees.Rows.Clear();
            IRestRequest request;
            IRestResponse<ResultTraineeRes> response;
            List<TraineeRes> lisTrainees = new List<TraineeRes>();
            if (SoHV <= 50)
            {
                request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 50);
                response = client.Get<ResultTraineeRes>(request);

                if (chkbSortID.Checked)
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = response.Data.items.OrderByDescending(item => item.outdoor_distance).ToList();

            }
            else
            {
                int Page = SoHV / 50 + 1;
                for (int k = 0; k < Page; k++)
                {
                    request = new RestRequest("/trainees", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page", k + 1).AddParameter("page_size", 50);
                    response = client.Get<ResultTraineeRes>(request);

                    lisTrainees.AddRange(response.Data.items.ToList());

                }
                if (chkbSortID.Checked)
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_hour).ToList();
                else
                    lisTrainees = lisTrainees.OrderByDescending(item => item.outdoor_distance).ToList();
            }
            foreach (TraineeRes trainee in lisTrainees)
            {
                count++;
                //indexV = (int)Math.Floor((count - 2) / NumTraineeVehicles);
                //if (indexV > (Vehicles.Count - 1))
                //else
                //    sPlate = Vehicles[indexV].plate;
                dgwTrainees.Rows.Add((count - 1).ToString(), trainee.id.ToString(), trainee.ho_va_ten, trainee.ngay_sinh, trainee.synced_outdoor_hours.ToString(), trainee.synced_outdoor_distance.ToString(),
                    (trainee.outdoor_hour / 3600).ToString(), (trainee.outdoor_distance / 1000).ToString(), trainee.outdoor_session_count.ToString(), trainee.ma_dk, trainee.anh_chan_dung);
            }
            int row = dgvCoures.CurrentRow.Index;
            string fileName = "C:\\Report_DAT\\" + dgvCoures[1, row].Value.ToString() + ".xls";
            CreatFileExcelReportCouse(fileName, dgvCoures.CurrentRow.Cells[1].Value.ToString(), dgvCoures.CurrentRow.Cells[3].Value.ToString(), dgvCoures.CurrentRow.Cells[5].Value.ToString(), dgvCoures.CurrentRow.Cells[6].Value.ToString(), "Ko dung den", lisTrainees);
            if (System.IO.File.Exists(fileName))
            {
                PrintMyExcelFile(fileName);
            }
        }

        private void btnFindCouse_Click(object sender, EventArgs e)
        {
            string searchValue = txtFind.Text.Trim();
            if (searchValue.Length < 1)
            {
                MessageBox.Show("Chưa nhập thông tin tìm kiếm");
                return;
            }
            dgvCoures.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dgvCoures.Rows)
                {
                    if (row.Index == (dgvCoures.Rows.Count - 1))
                    {
                        MessageBox.Show("Không tìm thấy");
                        break;
                    }
                    if (row.Cells[2].Value.ToString().ToUpper().Equals(searchValue.ToUpper()))
                    {
                        row.Selected = true;
                        dgvCoures.CurrentCell = dgvCoures.Rows[row.Index].Cells[0];
                        break;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnFindVehicle_Click(object sender, EventArgs e)
        {
            string searchValue = txtFind.Text.Trim();
            if (searchValue.Length < 1)
            {
                MessageBox.Show("Chưa nhập thông tin tìm kiếm");
                return;
            }
            dgvVehicles.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dgvVehicles.Rows)
                {
                    if (row.Index == (dgvVehicles.Rows.Count - 1))
                    {
                        MessageBox.Show("Không tìm thấy");
                        break;
                    }
                    if (row.Cells[1].Value.ToString().ToUpper().Equals(searchValue.ToUpper()))
                    {
                        row.Selected = true;
                        dgvVehicles.CurrentCell = dgvVehicles.Rows[row.Index].Cells[0];
                        break;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }

        }

        private void btnFindTrainee_Click(object sender, EventArgs e)
        {
            string searchValue = txtFind.Text.Trim();
            if (searchValue.Length < 1)
            {
                MessageBox.Show("Chưa nhập thông tin tìm kiếm");
                return;
            }
            dgwTrainees.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dgwTrainees.Rows)
                {
                    if (row.Index == (dgwTrainees.Rows.Count - 1))
                    {
                        MessageBox.Show("Không tìm thấy");
                        break;
                    }
                    if (row.Cells[2].Value.ToString().ToUpper().Contains(searchValue.ToUpper()))
                    {
                        row.Selected = true;
                        dgwTrainees.CurrentCell = dgwTrainees.Rows[row.Index].Cells[0];
                        break;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }

        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            //Form3 f3 = new Form3();//mo form duyet excel cua Bo
            //f3.ShowDialog();

            //mo excel cua csdt

            if (!bLogined)
            {
                MessageBox.Show("Bạn chưa đăng nhập");
                return;
            }
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls",
                FilterIndex = 1,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtLogs.Text = "";
                Cursor.Current = Cursors.WaitCursor;
                HocVienExcels.Clear();

                Microsoft.Office.Interop.Excel.Application excelApplication;
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.Visible = false;
                //string fileName = "C:\\sampleExcelFile.xlsx";

                //open the workbook
                string FileBaoCaoName = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
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
                for (int row = 2; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    if (row == 194)
                        row = 194;
                    HocVienExcel items = new HocVienExcel();
                    if (valueArray[row, 1] is null)
                        break;
                    items.STT = Int32.Parse(valueArray[row, 1].ToString());
                    if (items.STT != (row - 1))
                        break;
                    items.MaDangKy = valueArray[row, 2].ToString();
                    items.HoVaTen = valueArray[row, 3].ToString();

                    HocVienExcels.Add(items);
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
                //HocVienExcels = HocVienExcels.OrderBy(item => item.MaHocVien).ThenBy(item => item.ThoiGianPhienHoc).ToList();
                MessageBox.Show("Open excel so hoc vien: " + HocVienExcels.Count.ToString());

                dgwTrainees.Rows.Clear();
                IRestRequest request;
                IRestResponse<ResultTraineeRes> response;
                List<TraineeRes> lisTrainees = new List<TraineeRes>();
                for (int k = 0; k < HocVienExcels.Count; k++)
                {
                    //if ((k % 100) == 0) 
                    //    txtLogs.Text = ""; //reset
                    txtLogs.AppendText(k.ToString() + "-");// + HocVienExcels.Count.ToString() + "__");
                    //request = new RestRequest("/trainees", Method.GET).AddQueryParameter("name", "17005-20230804145808700").AddParameter("page_size", 50);
                    request = new RestRequest("/trainees", Method.GET).AddQueryParameter("name", HocVienExcels[k].MaDangKy).AddParameter("page_size", 50);
                    response = client.Get<ResultTraineeRes>(request);

                    lisTrainees.AddRange(response.Data.items.ToList());
                }
                string ConvertNgaySinh = "";
                //int count = 1;
                foreach (TraineeRes trainee in lisTrainees)
                {
                    string STT = HocVienExcels.SingleOrDefault(x => x.MaDangKy == trainee.ma_dk).STT.ToString();
                    ConvertNgaySinh = trainee.ngay_sinh.Substring(8, 2) + "/" + trainee.ngay_sinh.Substring(5, 2) + "/" + trainee.ngay_sinh.Substring(0, 4);
                    dgwTrainees.Rows.Add(STT, trainee.id.ToString(), trainee.ho_va_ten, ConvertNgaySinh, trainee.synced_outdoor_hours.ToString(), trainee.synced_outdoor_distance.ToString(),
                        (trainee.outdoor_hour / 3600).ToString(), (trainee.outdoor_distance / 1000).ToString(), trainee.outdoor_session_count.ToString(), trainee.ma_dk, trainee.anh_chan_dung);
                }
                string fileName = "C:\\Report_DAT\\BaoCaoTuExcel_" + FileBaoCaoName + ".xls";
                CreatFileExcelReportCouse_FromExcel(fileName, FileBaoCaoName, lisTrainees);
                OpenMyExcelFile(fileName);
            }
        }
        private void xemDanhSáchPhiênHọcToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (accessToken == "")
            {
                return;
            }
            int count = 1;
            int SoHV = Int32.Parse(dgvCoures.CurrentRow.Cells[4].Value.ToString());
            TenKhoaHoc = dgvCoures.CurrentRow.Cells[1].Value.ToString() + " ( " + dgvCoures.CurrentRow.Cells[2].Value.ToString() + " )";
            HangDaoTao = dgvCoures.CurrentRow.Cells[3].Value.ToString();
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 500).AddParameter("synced", 1);
                var response2 = client.Get<ResultSessionRes>(request2);
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }

            Sessions = Sessions.OrderBy(item => item.trainee_name).ThenBy(item => item.start_time).ToList();
            dgvSessions.Rows.Clear();
            DateTime StartTime;
            string ViPham = "";
            Boolean GetAll = true;
            if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                GetAll = true;
            else
                GetAll = false;
            foreach (SessionRes session in Sessions)
            {
                if ((GetAll == false) && (chkNonCheck.Checked == false) && (session.sync_status == 0))
                    continue;
                if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (session.sync_status < 0))
                    continue;
                if ((GetAll == false) && (chkCheckOk.Checked == false) && (session.sync_status > 0))
                    continue;

                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                if (session.sync_status == 0)
                    ViPham = "Chưa kiểm tra";
                else if (session.sync_status > 0)
                    ViPham = "Không vi phạm";
                else
                    ViPham = session.sync_error;
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
            }
        }

        private void inDanhSáchPhiênHọcToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TenKhoaHoc = dgvCoures.CurrentRow.Cells[1].Value.ToString() + " ( " + dgvCoures.CurrentRow.Cells[2].Value.ToString() + " )";
            HangDaoTao = dgvCoures.CurrentRow.Cells[3].Value.ToString();
            if (chkbDongBo.Checked)
            {
                var request2 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 500).AddParameter("synced", 1);
                var response2 = client.Get<ResultSessionRes>(request2);
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response2.Content);
            }
            else
            {
                var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("course_id", dgvCoures.CurrentRow.Cells[0].Value.ToString()).AddParameter("page_size", 500);
                var response3 = client.Get<ResultSessionRes>(request3);
                Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);
            }
            Sessions = Sessions.OrderBy(item => item.trainee_name).ThenBy(item => item.start_time).ToList();
            int row = dgvCoures.CurrentRow.Index;
            string Ngay = DateTime.Now.Day.ToString();
            if (Ngay.Length < 2) Ngay = "0" + Ngay;
            string Thang = DateTime.Now.Month.ToString();
            if (Thang.Length < 2) Thang = "0" + Thang;
            Thang = Ngay + Thang + DateTime.Now.Year.ToString();
            string fileName = "C:\\Report_DAT\\" + dgvCoures[1, row].Value.ToString() + "_" + Thang + ".xls";
            CreatFileExcelReportCouseSession(fileName, HangDaoTao, TenKhoaHoc, Sessions);
            dgvSessions.Rows.Clear();
            DateTime StartTime;
            string ViPham = "";
            Boolean GetAll = true;
            if ((chkNonCheck.Checked == false) && (chkCheckOk.Checked == false) && (chkCheckNonOk.Checked == false))
                GetAll = true;
            else
                GetAll = false;
            foreach (SessionRes session in Sessions)
            {
                if ((GetAll == false) && (chkNonCheck.Checked == false) && (session.sync_status == 0))
                    continue;
                if ((GetAll == false) && (chkCheckNonOk.Checked == false) && (session.sync_status < 0))
                    continue;
                if ((GetAll == false) && (chkCheckOk.Checked == false) && (session.sync_status > 0))
                    continue;

                StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                if (session.sync_status == 0)
                    ViPham = "Chưa kiểm tra";
                else if (session.sync_status > 0)
                    ViPham = "Không vi phạm";
                else
                    ViPham = session.sync_error;
                //dgvSessions.Rows.Add(session.id.ToString(), session.session_id, session.start_time, session.duration.ToString(), session.distance.ToString(), (session.faceid_failed_count + session.faceid_success_count).ToString());
                dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
            }
            OpenMyExcelFile(fileName);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnOpenExcelCT_Click(object sender, EventArgs e)
        {
            if (!bLogined)
            {
                MessageBox.Show("Bạn chưa đăng nhập");
                return;
            }
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls",
                FilterIndex = 1,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtLogs.Text = "";
                Cursor.Current = Cursors.WaitCursor;
                HocVienExcels.Clear();

                Microsoft.Office.Interop.Excel.Application excelApplication;
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.Visible = false;
                //string fileName = "C:\\sampleExcelFile.xlsx";

                //open the workbook
                string FileBaoCaoName = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
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
                for (int row = 2; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    HocVienExcel items = new HocVienExcel();
                    if (valueArray[row, 1] is null)
                        break;
                    items.STT = Int32.Parse(valueArray[row, 1].ToString());
                    if (items.STT != (row - 1))
                        break;
                    items.MaDangKy = valueArray[row, 2].ToString();
                    items.HoVaTen = valueArray[row, 3].ToString();

                    HocVienExcels.Add(items);
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
                //HocVienExcels = HocVienExcels.OrderBy(item => item.MaHocVien).ThenBy(item => item.ThoiGianPhienHoc).ToList();
                MessageBox.Show("Open excel so hoc vien: " + HocVienExcels.Count.ToString());

                dgvSessions.Rows.Clear();

                List<SessionRes> SessionsFromExel = new List<SessionRes>();
                for (int l = 0; l < HocVienExcels.Count; l++)
                {
                    //if ((k % 100) == 0) 
                    //    txtLogs.Text = ""; //reset
                    txtLogs.AppendText(l.ToString() + "-");// + HocVienExcels.Count.ToString() + "__");
                    //request = new RestRequest("/trainees", Method.GET).AddQueryParameter("name", HocVienExcels[k].MaDangKy).AddParameter("page_size", 50);
                    //response = client.Get<ResultTraineeRes>(request);
                    //lisTrainees.AddRange(response.Data.items.ToList());
                    var request3 = new RestRequest("/outdoor-sessions", Method.GET).AddQueryParameter("ho_va_ten", HocVienExcels[l].MaDangKy).AddQueryParameter("status", "2").AddParameter("page_size", 500);
                    var response3 = client.Get<ResultSessionRes>(request3);
                    Sessions = JsonConvert.DeserializeObject<List<SessionRes>>(response3.Content);

                    SessionsFromExel.AddRange(Sessions);
                }
                //string ConvertNgaySinh = "";
                DateTime StartTime, EndTime;
                int count = 1;
                List<InforSessionReport> InforSessionReports = new List<InforSessionReport>();
                
                foreach (SessionRes session in SessionsFromExel)
                {
                    //string STT = HocVienExcels.SingleOrDefault(x => x.MaDangKy == trainee.ma_dk).STT.ToString();
                    //ConvertNgaySinh = trainee.ngay_sinh.Substring(8, 2) + "/" + trainee.ngay_sinh.Substring(5, 2) + "/" + trainee.ngay_sinh.Substring(0, 4);

                    StartTime = DateTime.ParseExact(session.start_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    EndTime = DateTime.ParseExact(session.end_time.Substring(0, 19).Replace('T', ' '), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    dgvSessions.Rows.Add(session.session_id, session.trainee_name, StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString(),
                    Truncate(((double)session.duration / 3600), 2).ToString(), Truncate(((double)session.distance / 1000), 2).ToString(), session.vehicle_plate,
                    session.faceid_success_count.ToString() + "/" + (session.faceid_failed_count + session.faceid_success_count).ToString(), session.synced.ToString(), ViPham);
                    
                    InforSessionReport InforSessionReportItem = new InforSessionReport();
                    InforSessionReportItem.Sessionid = session.id;
                    InforSessionReportItem.MaPhienHoc = session.session_id;
                    InforSessionReportItem.StartTime = StartTime.ToShortDateString() + " " + StartTime.ToLongTimeString();
                    InforSessionReportItem.StopTime = EndTime.ToShortDateString() + " " + EndTime.ToLongTimeString();
                    InforSessionReportItem.ThoiGianTH = Truncate(((double)session.duration / 3600), 3).ToString();//.Replace(".", ",");
                    InforSessionReportItem.QuangDuongTH = Truncate(((double)session.distance / 1000), 3).ToString();//.Replace(".",",");
                    InforSessionReportItem.BienSoXe = session.vehicle_plate;
                    InforSessionReportItem.HangXeTL = session.vehicle_hang;
                    InforSessionReportItem.MaHocVien = session.trainee_ma_dk;
                    InforSessionReportItem.HoTenHocVien = session.trainee_name;
                    InforSessionReportItem.TenKhoaHoc = session.ten_khoa_hoc;
                    InforSessionReportItem.MaKhoaHoc = "KXD";
                    InforSessionReportItem.LoaiKhoaHoc = "KXD";
                    foreach (CourseRes course in ListCourses)
                    {
                        if (course.ten_khoa_hoc.Trim().ToUpper() == session.ten_khoa_hoc.Trim().ToUpper())
                        {
                            InforSessionReportItem.MaKhoaHoc = course.ma_khoa_hoc;
                            InforSessionReportItem.LoaiKhoaHoc = course.ma_hang_dao_tao;
                            break;
                        }
                    }

                    InforSessionReports.Add(InforSessionReportItem);
                }
                string fileName = "C:\\Report_DAT\\BaoCaoCTTuExcel_" + FileBaoCaoName + ".xls";
                CreatFileExcelReportSession_FromExcel(fileName, FileBaoCaoName, InforSessionReports);
                OpenMyExcelFile(fileName);
            }
        }
    }
}