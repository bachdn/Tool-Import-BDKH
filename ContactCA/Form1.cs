using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.IO;
using System.Data.OleDb;
using System.Globalization;
using Import_KHBD.ServiceReference1;
using Microsoft.Win32;

namespace Import_KHBD
{


    public partial class fromContact : Form
    {

        public static ServiceReference1.USD_WebServiceSoapClient ca;
        public static string ObjectResult, ObjectHandle,getUID;
        public static int SID;
        public static double kehoachbd_seconds,kehoachbd_bs_seconds;
        public static string DateNow;

        public static XDocument xDoc;
        
        public fromContact()
        {
            InitializeComponent();

        }

        public List<string> lstCACheckImport;
        /// <summary>
        /// Kiểm tra và ghi log đầu vào chương trình. Đọc các tham số từ file app.config
        /// </summary>
        public bool Check_Config()
        {

            #region Kết nối CA
            try
            {
                ca = new ServiceReference1.USD_WebServiceSoapClient();
                xDoc = new XDocument();
                SID = ca.login(txtUserName.Text,txtPassword.Text);
                WriteOLog("Dang nhap CA thanh cong");
               
            }
            catch
            {
                WriteOLog("Dang nhap CA khong thah cong");
                return false;
            }
            return true;
            #endregion
        }

        ///  Ghi log Chuong trinh vào file log.txt trong thư mục bin
        public void WriteOLog(string value)
        {
            using (StreamWriter write = new StreamWriter(txtFolderLog.Text + "\\log.txt", true))
            {
                write.WriteLine(value);
            }
        }

        private int checkEmpty()
        {
            if (string.IsNullOrEmpty(txtFileExcel.Text) || string.IsNullOrEmpty(txtFolderLog.Text) || string.IsNullOrEmpty(txtUserName.Text) || string.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("Enter requement field.", "Message");

                return 0;
            }
            return 1;
        }

        private void btnKetNoi_Click(object sender, EventArgs e)
        {
            if (checkEmpty() == 0) return;
            if (Check_Config())
            {
                MessageBox.Show("Connect CA success!", "Message");
                WriteOLog("Ket noi CA thanh cong!");
            }
            else
            {
                MessageBox.Show("Connect CA fail!", "Message");
                WriteOLog("Ket noi CA that bai!");
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtFileExcel.Text = openFileDialog1.FileName;
                // bindDataSheetExcel(txtFileExcel.Text);
            }

        }

        

        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkEmpty() == 0) return;
                ca = new ServiceReference1.USD_WebServiceSoapClient();
                xDoc = new XDocument();
                SID = ca.login(txtUserName.Text, txtPassword.Text);
                getUID = ca.getHandleForUserid(SID, txtUserName.Text);
               
                DateNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss tt");
                WriteOLog("\t ----------------- Start - Date Time: " + DateNow + "-----------------");
                //DataTable tblContact = ReadExcel(txtFileExcel.Text.Trim());
                //DataTable db = InsertCSV(txtFileExcel.Text.Trim());

                DataTable rd = tbReadEx(txtFileExcel.Text.Trim(), "ListMaintenance");
                InsertContact(rd);
            }
            catch
            {
                MessageBox.Show("Connect CA fail!", "Message");
            }

        }
        public DataTable tbReadEx(string strFilePath, string strSheet)
        {
            DataTable tblDataExcel = new DataTable();
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text\"", strFilePath);
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0\"", strFilePath);
            using (OleDbConnection dbConnection = new OleDbConnection(strConn))
            {

                try
                {
                    using (OleDbDataAdapter dbAdapter = new OleDbDataAdapter("SELECT * FROM [" + strSheet + "$]", dbConnection)) //rename sheet if required!
                        dbAdapter.Fill(tblDataExcel);
                }
                catch (Exception)
                {
                    //throw;
                    return null;
                }

            }
            return tblDataExcel;
               
        }


        private void InsertContact(DataTable tbl)
        {
            int iSuccess = 0;
            int iTrung = 0;
            int iFail = 0;
            try
            {
                if (tbl != null && tbl.Rows.Count > 0)
                {
                    for (int i = 0; i < tbl.Rows.Count; i++)
                    {
                        string terminalID = tbl.Rows[i][5].ToString().Trim();
                        string assignee = tbl.Rows[i][2].ToString().Trim();
                        string group = tbl.Rows[i][3].ToString().Trim();
                        string kehoachbd = tbl.Rows[i][1].ToString();
                        string kehoachbd_bs = DateTime.Now.ToString();
                        string ky_bd = tbl.Rows[i][0].ToString().Trim();
                        string affected_enduser = string.Empty;
                        switch (group) 
                            {
                                case "An Giang":
                                    group = "BGT_AG_TECH";
                                    break;
                                case "Bà Rịa-Vũng Tàu":
                                    group = "BGT_BV_TECH";
                                    break;
                                case "Bạc Liêu":
                                    group = "BGT_BL_TECH";
                                    break;
                                case "Bắc Kạn":
                                    group = "BGT_BK_TECH";
                                    break;
                                case "Bắc Giang":
                                    group = "BGT_BG_TECH";
                                    break;
                                case "Bắc Ninh":
                                    group = "BGT_BN_TECH";
                                    break;
                                case "Bến Tre":
                                    group = "BGT_BT_TECH";
                                    break;
                                case "Bình Dương":
                                    group = "BGT_BD_TECH";
                                    break;
                                case "Bình Định":
                                    group = "BGT_BĐ_TECH";
                                    break;
                                case "Bình Phước":
                                    group = "BGT_BP_TECH";
                                    break;
                                case "Bình Thuận":
                                    group = "BGT_BTh_TECH";
                                    break;
                                case "Cà Mau":
                                    group = "BGT_CM_TECH";
                                    break;
                                case "Cao Bằng":
                                    group = "BGT_CB_TECH";
                                    break;
                                case "Cần Thơ":
                                    group = "BGT_CT_TECH";
                                    break;
                                case "Đà Nẵng":
                                    group = "BGT_ĐNA_TECH";
                                    affected_enduser = "5A9E8EF20972554286CCE745ED5A80D8";
                                    break;
                                case "Đắk Lắk":
                                    group = "BGT_ĐL_TECH";
                                    break;
                                case "Đắk Nông":
                                    group = "BGT_ĐNo_TECH";
                                    break;
                                case "Điện Biên":
                                    group = "BGT_ĐB_TECH";
                                    break;
                                case "Đồng Nai":
                                    group = "BGT_ĐN_TECH";
                                    break;
                                case "Đồng Tháp":
                                    group = "BGT_ĐT_TECH";
                                    break;
                                case "Gia Lai":
                                    group = "BGT_GL_TECH";
                                    break;
                                case "Hà Giang":
                                    group = "BGT_HG_TECH";
                                    break;
                                case "Hà Nam":
                                    group = "BGT_HNA_TECH";
                                    affected_enduser = "FAC114F88B7A924FAB76E3C3ADE719A2";
                                    break;
                                case "Hà Nội":
                                    group = "BGT_HAN_TECH";
                                    affected_enduser = "FAC114F88B7A924FAB76E3C3ADE719A2";
                                    break;
                                case "Hà Tĩnh":
                                    group = "BGT_HT_TECH";
                                    break;
                                case "Hải Dương":
                                    group = "BGT_HD_TECH";
                                    break;
                                case "Hải Phòng":
                                    group = "BGT_HP_TECH";
                                    affected_enduser = "90E56D8C3E08DA42BF0F77967D5E26A3";
                                    break;
                                case "Hậu Giang":
                                    group = "BGT_HG_TECH";
                                    break;
                                case "Hòa Bình":
                                    group = "BGT_HB_TECH";
                                    break;
                                case "TP Hồ Chí Minh":
                                    group = "BGT_HCM_TECH";
                                    affected_enduser = "EEB7DB4F6368E04FBE4296ECEC6C46CF";
                                    break;
                                case "Hưng Yên":
                                    group = "BGT_HY_TECH";
                                    break;
                                case "Khánh Hoà":
                                    group = "BGT_KH_TECH";
                                    break;
                                case "Kiên Giang":
                                    group = "BGT_KG_TECH";
                                    break;
                                case "Kon Tum":
                                    group = "BGT_KT_TECH";
                                    break;
                                case "Lai Châu":
                                    group = "BGT_LC_TECH";
                                    break;
                                case "Lạng Sơn":
                                    group = "BGT_LS_TECH";
                                    break;
                                case "Lào Cai":
                                    group = "BGT_LCa_TECH";
                                    break;
                                case "Lâm Đồng":
                                    group = "BGT_LĐ_TECH";
                                    break;
                                case "Long An":
                                    group = "BGT_LA_TECH";
                                    break;
                                case "Nam Định":
                                    group = "BGT_NĐ_TECH";
                                    break;
                                case "Nghệ An":
                                    group = "BGT_NA_TECH";
                                    affected_enduser = "FAC114F88B7A924FAB76E3C3ADE719A2";
                                    break;
                                case "Ninh Bình":
                                    group = "BGT_NB_TECH";
                                    break;
                                case "Ninh Thuận":
                                    group = "BGT_NT_TECH";
                                    break;
                                case "Phú Thọ":
                                    group = "BGT_PT_TECH";
                                    break;
                                case "Phú Yên":
                                    group = "BGT_PY_TECH";
                                    break;
                                case "Quảng Bình":
                                    group = "BGT_QB_TECH";
                                    break;
                                case "Quảng Nam":
                                    group = "BGT_QNa_TECH";
                                    break;
                                case "Quảng Ngãi":
                                    group = "BGT_QNg_TECH";
                                    break;
                                case "Quảng Ninh":
                                    group = "BGT_QN_TECH";
                                    break;
                                case "Quảng Trị":
                                    group = "BGT_QT_TECH";
                                    break;
                                case "Sóc Trăng":
                                    group = "BGT_ST_TECH";
                                    break;
                                case "Sơn La":
                                    group = "BGT_SL_TECH";
                                    break;
                                case "Tây Ninh":
                                    group = "BGT_TN_TECH";
                                    affected_enduser = "EEB7DB4F6368E04FBE4296ECEC6C46CF";
                                    break;
                                case "Thái Bình":
                                    group = "BGT_TB_TECH";
                                    break;
                                case "Thái Nguyên":
                                    group = "BGT_TNg_TECH";
                                    break;
                                case "Thanh Hoá":
                                    group = "BGT_TH_TECH";
                                    break;
                                case "Thừa Thiên Huế":
                                    group = "BGT_TTH_TECH";
                                    affected_enduser = "5A9E8EF20972554286CCE745ED5A80D8";
                                    break;
                                case "Tiền Giang":
                                    group = "BGT_TG_TECH";
                                    break;
                                case "Trà Vinh":
                                    group = "BGT_TV_TECH";
                                    break;
                                case "Tuyên Quang":
                                    group = "BGT_TQ_TECH";
                                    break;
                                case "Vĩnh Long":
                                    group = "BGT_VL_TECH";
                                    break;
                                case "Vĩnh Phúc":
                                    group = "BGT_VP_TECH";
                                    break;
                                case "Yên Bái":
                                    group = "BGT_YB_TECH";
                                    break;

                            }
                        DateTime date = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                        if (string.IsNullOrEmpty(kehoachbd))
                        {
                            kehoachbd_seconds = 0;
                        }
                        else
                        {
                            DateTime dt2 = Convert.ToDateTime(kehoachbd);
                            TimeSpan diff = dt2 - date;
                            kehoachbd_seconds = diff.TotalSeconds+1800;
                        }

                        if (string.IsNullOrEmpty(kehoachbd_bs))
                        {
                            kehoachbd_bs_seconds = 0;
                        }
                        else
                        {
                            DateTime dt2 = Convert.ToDateTime(kehoachbd_bs);
                            TimeSpan diff = dt2 - date;
                            kehoachbd_bs_seconds = diff.TotalSeconds-25272;
                        }
                        
                        List<string> lstCATerminal = GetDataTerminal(terminalID);
                        List<string> lstCAAssignee = GetDataAssignee(assignee);
                        List<string> lstCaGroupAssignee = GetDataGroup(group);
                        if (lstCATerminal != null && lstCATerminal.Count > 0)
                        {
                             lstCACheckImport =
                                GetDataSummary(lstCATerminal[2].ToString(), ky_bd);
                        }
                        if (lstCAAssignee != null && lstCAAssignee.Count > 0)
                        { 
                            assignee=lstCAAssignee[0].ToString();
                        }
                        else 
                        {
                            assignee=string.Empty;
                        }
                        if (lstCATerminal != null && lstCATerminal.Count > 0 && kehoachbd_seconds > 0 && lstCACheckImport.Count == 0 && lstCaGroupAssignee != null && lstCaGroupAssignee.Count>0) 
                        {
                        
                            try
                            {
                                string[] attributes = { };
                                string newTicketHandler = string.Empty;
                                string newTicketNumber = "ref_num";
                                string duplication_id = string.Empty;
                                string returnUserData = string.Empty;
                                string returnAppData = string.Empty;
                                string check_khbdbs_empty;

                                if (kehoachbd_bs_seconds==0)
                                {
                                    check_khbdbs_empty = string.Empty;
                                }
                                else
                                {
                                    check_khbdbs_empty = kehoachbd_bs_seconds.ToString();
                                }
                                
                                //ca.createObject(SID, "cr", Vals, attributes, ref  ObjectResult, ref ObjectHandle);

                                ca.createTicket(SID, "Bảo dưỡng chu kỳ " + ky_bd + "_" + lstCATerminal[2].ToString(), "I", getUID, string.Empty, duplication_id, ref newTicketHandler, ref newTicketNumber, ref returnUserData,
                                   ref returnAppData);
                                string ticket = newTicketNumber;
                                
                               
                                iSuccess++;
                                List<string> lstCATicket = GetDataPersistent(ticket);
                                lstCATicket[0].ToString();
                                string[] Vals = { "affected_resource", lstCATerminal[0].ToString(), "priority", "pri:502", "category", "pcat:402404", "zbgt_support_type", "400001","zbgt_processing_unit", "400001", "zbgt_maintenance_date", kehoachbd_seconds.ToString(), "summary", "Kỳ " + ky_bd + "_" + lstCATerminal[2].ToString(), "status", "OP", "assignee", assignee, "group", lstCaGroupAssignee[0].ToString(), "customer", affected_enduser, "zbgt_open_case_date", kehoachbd_bs_seconds.ToString() };
                                ca.updateObject(SID,newTicketHandler, Vals, attributes);
                                
                            }
                       catch
                            {
                                iFail++;
                            }
                        }
                        else
                        {
                           
                            if (!string.IsNullOrEmpty(terminalID) && lstCATerminal.Count==0) 
                            {
                                WriteOLog(terminalID + " Khong co terminal trong he thong");
                                iFail++;
                            }
                            else if (!string.IsNullOrEmpty(terminalID) && (lstCACheckImport!=null && lstCACheckImport.Count>0))
                            {
                                WriteOLog(terminalID+" Da import");
                                iTrung++;
                            }
                            else if (kehoachbd_seconds == 0 && !string.IsNullOrEmpty(terminalID))
                            {
                                    WriteOLog(terminalID+ " Khong co thong tin ke hoach bao duong");
                                iFail++;
                            }
                        }
                    }
                    if (iTrung > 0)
                        WriteOLog("Trung " + (iTrung) + " terminals.");

                    WriteOLog("insert thanh cong " + (iSuccess) + " tickets.");

                    if (iFail > 0)
                        WriteOLog("insert Khong thanh cong " + (iFail) + " tickets.");

                    string strThongBao = string.Empty;
                    strThongBao = "Insert success " + iSuccess + " tickets.";
                    if (iTrung > 0)
                        strThongBao += "\n Dupplicate " + iTrung + " tickets";
                    if (iFail > 0)
                        strThongBao += "\n Insert fail " + iFail + " tickets";

                    MessageBox.Show(strThongBao, "Message");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("gap loi trong qua trinh inssert", "Message");
            }

        }


        

        public static List<string> GetDataPcat(string pcat_sym)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "sym", "id" };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, "pcat", "sym like '" + pcat_sym + "'", -1, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static List<string> GetDataTerminal(string terminal_id)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "id", "zhw_id", "zcode" };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, "zterminal", "zcode = '" + terminal_id + "'", -1, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

       //public static List<string> GetDataContactProvince(string province)
       // {
       //     try
       //     {
       //         List<string> lst = new List<string>();
       //         string[] attr = { "id", "province_v2" };
       //         XDocument xml = new XDocument();
       //         string UDSObj = ca.doSelect(SID, "zlocation_area", "province_v2 = '" + province + "'", -1, attr);
       //         xml = XDocument.Parse(UDSObj);

       //         foreach (XElement element in xml.Descendants("UDSObject"))
       //         {
       //             foreach (XElement EAttr in element.Descendants("Attribute"))
       //             {
       //                 lst.Add(EAttr.Element("AttrValue").Value);
       //             }
       //         }
       //         return lst;
       //     }
       //     catch (Exception ex)
       //     {
       //         return null;
       //     }
       // }

        public static List<string> GetDataPersistent(string ticket)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "persistent_id", "ref_num","id" };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, "cr", "ref_num = '" + ticket + "'",10, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static List<string> GetDataSummary(string terminalid, string summary)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "persistent_id", "summary" };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, "cr", "summary = '" +"Kỳ "+ summary+"_"+terminalid + "'", 10, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static List<string> GetDataAssignee(string userid)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "id", "userid","combo_name" };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, "agt", "userid = '" + userid + "'", 10, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static List<string> GetDataGroup(string group_name)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "id","last_name" };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, "grp", "last_name = '" + group_name + "'", 10, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string txt = folderBrowserDialog1.SelectedPath.ToString();
                txtFolderLog.Text = txt;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            DateNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss tt");
            WriteOLog("\t ----------------- Exit - Date Time: " + DateNow + "-----------------");
            this.Close();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void txtSummary_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtUserName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtFileExcel_TextChanged(object sender, EventArgs e)
        {

        }


    }
}
