using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Diagnostics;

namespace WordAutomation_01
{
    public partial class Form1 : Form
    {
        string FFrom = "";
        string DiHoacVe = "";
        string FTo = "";
        string EnglishDays = "";
        string VietnameseDays = "";
        string NoiDi = "", NoiDen = "";
        string TenTaiXe = "", SDT_TaiXe = "", TenPhuXe = "", SDT_PhuXe = "", SoXe = "", SoRomooc = "";
        string FSoXe = "", FRomooc = "", FTaiXe = "", FSDT_TaiXe = "", FPhuXe = "", FSDT_PhuXe = "";
        string path = Environment.CurrentDirectory.ToString();
        string conn_string = @"Data Source=.\SQLEXPRESS;Initial Catalog=AnhKhoa;Integrated Security=True";
        SqlConnection conn = new SqlConnection();

        Object oMissing = System.Reflection.Missing.Value;
        // if you want your document to be saved as pdf
        //Object format = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
        Word.Application oWord = new Word.Application();
        Word.Document oWordDoc = new Word.Document();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Click(object sender, EventArgs e)
        {
            status_idle();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Database_Connect();
            Danh_Sach_Tai_Xe();
            Danh_Sach_Tuyen_Duong();
            Preview_Word_File();
        }

        public void str()
        {

        }

        private void cbb_TaiXe_TextChanged(object sender, EventArgs e)
        {
            Preview_Word_File();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            oWordDoc.Close();
            oWord.Quit();
        }

        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            oWordDoc.Close();
            oWord.Quit();
        }

        public void status_idle()
        {
            Danh_Sach_Tai_Xe();
            Danh_Sach_Tuyen_Duong();
        }

        public void Database_Connect()
        {
            conn.ConnectionString = conn_string;
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Không kết nối được Cơ Sở Dữ Liệu.");
            }
            finally
            {
                conn.Close();
            }
        }

        public void Danh_Sach_Tai_Xe()
        {

            SqlCommand cmd_sql = new SqlCommand();
            try
            {
                conn.ConnectionString = conn_string;
                conn.Open();
                cmd_sql.Connection = conn;
                string load_data = @"SELECT * FROM Tai_Xe";
                cmd_sql.CommandText = load_data;

                DbDataReader reader = cmd_sql.ExecuteReader();
                if (reader.HasRows)
                {
                    while(reader.Read())
                    {
                        TenTaiXe = reader.GetString(1);
                        cbb_TaiXe.Items.Add(TenTaiXe);

                        SDT_TaiXe = reader.GetString(2);
                        cbb_SDT_TaiXe.Items.Add(SDT_TaiXe);

                        TenPhuXe = reader.GetString(3);
                        cbb_PhuXe.Items.Add(TenPhuXe);

                        SDT_PhuXe = reader.GetString(4);
                        cbb_SDT_PhuXe.Items.Add(SDT_PhuXe);

                        SoXe = reader.GetString(5);
                        cbb_SoXe.Items.Add(SoXe);

                        SoRomooc = reader.GetString(6);
                        cbb_Romooc.Items.Add(SoRomooc);
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection." + ex);
            }
            finally
            {
                cmd_sql.Dispose();
                conn.Close();
            }
        }

        public void Danh_Sach_Tuyen_Duong()
        {

            SqlCommand cmd_sql = new SqlCommand();
            try
            {
                conn.ConnectionString = conn_string;
                conn.Open();
                cmd_sql.Connection = conn;
                string load_data = @"SELECT * FROM Tuyen_Duong";
                cmd_sql.CommandText = load_data;

                DbDataReader reader = cmd_sql.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        NoiDi = reader.GetString(1);
                        cbb_From.Items.Add(NoiDi);

                        NoiDen = reader.GetString(2);
                        cbb_To.Items.Add(NoiDen);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection." + ex);
            }
            finally
            {
                cmd_sql.Dispose();
                conn.Close();
            }
        }

        public void SetRange_TaiXe()
        {
            SqlCommand cmd_sql = new SqlCommand();
            try
            {
                conn.ConnectionString = conn_string;
                conn.Open();
                cmd_sql.Connection = conn;
                string load_data = "SELECT * FROM Tai_Xe WHERE TenTaiXe = N'"+cbb_TaiXe.Text+"'";
                cmd_sql.CommandText = load_data;

                DbDataReader reader = cmd_sql.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        SDT_TaiXe = reader.GetString(2);
                        cbb_SDT_TaiXe.Text = SDT_TaiXe;

                        TenPhuXe = reader.GetString(3);
                        cbb_PhuXe.Text = TenPhuXe;

                        SDT_PhuXe = reader.GetString(4);
                        cbb_SDT_PhuXe.Text = SDT_PhuXe;

                        SoXe = reader.GetString(5);
                        cbb_SoXe.Text = SoXe;

                        SoRomooc = reader.GetString(6);
                        cbb_Romooc.Text = SoRomooc;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection." + ex);
            }
            finally
            {
                cmd_sql.Dispose();
                conn.Close();
            }
        }

        public void Preview_Word_File()
        {
            lblCongTy.Text = "CÔNG TY TNHH NGUYỄN ANH KHOA \n" +
                        "115 Đường 768, Ấp 1, Xã Tân An, Huyện Vĩnh Cửu, Tỉnh Đồng Nai \n" +
                        "Tel : 0902.497380";
            lblCHXH.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM \n" +
                            "Độc Lập – Tự Do – Hạnh Phúc";

            lblTieude.Text = "DANH SÁCH XE CÔNG TY NGUYỄN ANH KHOA VẬN CHUYỂN HÀNG \n" + 
                                "CHO CÔNG TY CỔ PHẦN TICO";
            GetDaysOfWeekToVietnamese();
            string Ngay_label = datetime_DayStart.Value.Day.ToString();
            string Thang_label = datetime_DayStart.Value.Month.ToString();
            string Nam_label = datetime_DayStart.Value.Year.ToString();
            lblNgayThangNam.Text = VietnameseDays + ", Ngày " + Ngay_label + ", Tháng " + Thang_label + ", Năm " + Nam_label;

            lblDong1A.Text = "Kính Gửi: CHI  NHÁNH CÔNG TY CỔ PHẦN TICO";
            lblDong1B.Text = "Địa chỉ: 83 / 2B Khu Phố 1B ,Phường An Phú ,Thị Xã Thuận An, Tỉnh Bình Dương.";
            lblDong1C.Text = "MST: 0300769124001";
            lblDong1D.Text = "Hôm nay, Công Ty TNHH Nguyễn Anh Khoa xin gửi đến Quý Công Ty danh sách xe vận chuyển hàng cho Quý Công Ty như sau:";

            FFrom = cbb_From.Text;
            if (FFrom == "ABS")
            {
                FFrom = "NHÀ MÁY " + FFrom;
                DiHoacVe = "ĐI";
            }
            else
                DiHoacVe = "VỀ";
            FTo = cbb_To.Text; FSoXe = cbb_SoXe.Text; FRomooc = cbb_Romooc.Text; FTaiXe = cbb_TaiXe.Text;
            FSDT_TaiXe = cbb_SDT_TaiXe.Text; FPhuXe = cbb_PhuXe.Text; FSDT_PhuXe = cbb_SDT_PhuXe.Text;
            
            lblDong02A.Text = "XE VẬN CHUYỂN HÀNG TỪ "+ FFrom + " " + DiHoacVe + " " + FTo;
            Dong02B1.Text = "Số xe: " + FSoXe + " | " + "Số Rơ-moóc: " + FRomooc;
            Dong02B2.Text = "Tài xế: " + FTaiXe + " | " + "SĐT: " + FSDT_TaiXe;
            Dong02B3.Text = "Phụ xế: " + FPhuXe + " | " + "SĐT: " + FSDT_PhuXe;
            string t = "";
        }

        public void GetDaysOfWeekToVietnamese()
        {
            EnglishDays = datetime_DayStart.Value.DayOfWeek.ToString();
            switch (EnglishDays)
            {
                case "Sunday":
                    VietnameseDays = "Chủ Nhật";
                    break;
                case "Monday":
                    VietnameseDays = "Thứ Hai";
                    break;
                case "Tuesday":
                    VietnameseDays = "Thứ Ba";
                    break;
                case "Wednesday":
                    VietnameseDays = "Thứ Tư";
                    break;
                case "Thursday":
                    VietnameseDays = "Thứ Năm";
                    break;
                case "Friday":
                    VietnameseDays = "Thứ Sáu";
                    break;
                case "Saturday":
                    VietnameseDays = "Thứ Bảy";
                    break;
            }

        }

        public void general_word()
        {
            // path of the Word Template document
            Object oTemplatePath = path +@"\CTY_AK.dotx";
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            int iFields = 0;

            foreach (Word.Field myMergeField in oWordDoc.Fields)
            {
                iFields++;
                Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();

                    if (fieldName == "Day01")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        GetDaysOfWeekToVietnamese();
                        oWord.Selection.TypeText(VietnameseDays);
                    }
                    if (fieldName == "Day02")
                    {
                        myMergeField.Select();
                        string Day02 = datetime_DayStart.Value.Day.ToString();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(Day02);
                    }
                    if (fieldName == "Month")
                    {
                        myMergeField.Select();
                        string FMonth = datetime_DayStart.Value.Month.ToString();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FMonth);
                    }
                    if (fieldName == "Year")
                    {
                        myMergeField.Select();
                        string FYear = datetime_DayStart.Value.Year.ToString();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FYear);
                    }
                    FFrom = cbb_From.Text;
                    if (fieldName == "From")
                    {
                        myMergeField.Select();

                        // check whether the control text is empty
                        if (FFrom == "ABS")
                        {
                            FFrom = "NHÀ MÁY " + FFrom;
                            oWord.Selection.TypeText(FFrom);
                        }
                        else
                        {
                            oWord.Selection.TypeText(FFrom);
                        }
                    }
                    if (fieldName == "DiHoacVe")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        if (FFrom == "ABS")
                        {
                            DiHoacVe = "ĐI";
                            oWord.Selection.TypeText(DiHoacVe);
                        }
                        else
                        {
                            DiHoacVe = "VỀ";
                            oWord.Selection.TypeText(DiHoacVe);
                        }
                    }
                    FTo = cbb_To.Text;
                    if (fieldName == "To")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FTo);
                    }

                    //DONG 02
                    FSoXe = cbb_SoXe.Text;
                    if (fieldName == "FSoXe")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FSoXe);
                    }
                    FRomooc = cbb_Romooc.Text;
                    if (fieldName == "FRomooc")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FRomooc);
                    }
                    FTaiXe = cbb_TaiXe.Text;
                    if (fieldName == "FTaiXe")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FTaiXe);
                    }
                    FSDT_TaiXe = cbb_SDT_TaiXe.Text;
                    if (fieldName == "FSDT_TaiXe")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FSDT_TaiXe);
                    }
                    FPhuXe = cbb_PhuXe.Text;
                    if (fieldName == "FPhuXe")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FPhuXe);
                    }
                    FSDT_PhuXe = cbb_SDT_PhuXe.Text;
                    if (fieldName == "FSDT_PhuXe")
                    {
                        myMergeField.Select();
                        // check whether the control text is empty
                        oWord.Selection.TypeText(FSDT_PhuXe);
                    }
                }
            }

            //string s = path + @"\W";
            string s = path;
            string short_time = DateTime.Now.ToString("HH_mm_ss");
            object savePath = s + "\\Danh sach xe chay ngay " + datetime_DayStart.Value.ToString("dd_MM_yyyy__")+short_time+ ".docx";
            
            oWordDoc.SaveAs(ref savePath, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            
            if (MessageBox.Show(@"Export word file complete. Do you want open folder? ", "Export Word", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start("explorer.exe", s);
            }

        }

        private void btn_Send_Click(object sender, EventArgs e)
        {
            general_word();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            object doNotSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            oWordDoc.Close(ref doNotSaveChanges, ref oMissing, ref oMissing);
            oWord.Quit(ref doNotSaveChanges, ref oMissing, ref oMissing);
        }

            private void cbb_TaiXe_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetRange_TaiXe();
            Preview_Word_File();
        }
    }
}
