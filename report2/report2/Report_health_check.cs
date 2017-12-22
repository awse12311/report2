using Microsoft.Reporting.WinForms;
using MySql.Data.MySqlClient;
using report2;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace report2
{
    public partial class Report_health_check : Form
    {
        string sql,sql2, CheckPrintKind = "R00003";
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        //傳值
        private string _UserName;
        public string UserName
        {
            get { return _UserName; }
            set { _UserName = value; }
        }
        public void setUserValue()
        {
            this.Text += "   使用者:" + UserName;
        }

        

        public Report_health_check()
        {
            InitializeComponent();
            this.isemployee0.Checked = true;
            this.isemployee1.Checked = true;
            this.dtFromDate.Enabled = false;
            this.dtToDate.Enabled = false;
            this.btnLoad.Enabled = false;
            this.Emp_no.Checked = true;
            this.checkBox_sid.Checked = true;
            this.checkBox_name.Checked = true;
            
            dirNameText.Enabled = false;
            dirNameLord.Enabled = false;

        }
        //停用元件
        private void close()
        {
            reportViewer1.Enabled = false;
            sceneNoLoad.Enabled = false;
            nameLoad.Enabled = false;
            idLoad.Enabled = false;
            projectNoLoad.Enabled = false;
            companyLoad.Enabled = false;
            btnLoad.Enabled = false;
            arraydate.Enabled = false;
            project_search.Enabled = false;
            listBox1.Enabled = false;
            dirNameText.Enabled = false;
            dirNameLord.Enabled = false;
            projectNameLoad.Enabled = false;
            company_search.Enabled = false;
            deptNameLord.Enabled = false;
            this.Text = "資料讀取中...";
        }
        //開啟元件
        private void open()
        {
            reportViewer1.Enabled = true;
            sceneNoLoad.Enabled = true;
            nameLoad.Enabled = true;
            idLoad.Enabled = true;
            projectNoLoad.Enabled = true;
            companyLoad.Enabled = true;
            projectNameLoad.Enabled = true;
            deptNameLord.Enabled = true;

            arraydate.Enabled = true;
            project_search.Enabled = true;
            if (arraydate.Checked)
                btnLoad.Enabled = true;
            listBox1.Enabled = true;
            dirNameText.Enabled = true;
            dirNameLord.Enabled = true;
            company_search.Enabled = true;
            this.Text = ReportSelect.Text + "   使用者:" + UserName;
        }
        
        
        /// 查詢按鈕
        private void btnLoad_Click(object sender, EventArgs e)
        {

            //日期篩選查詢
            sql = reportSelect_SqlString("date", getFilterString().Remove(0, 3));

            Refresh(sql);

        }
        private void sceneNoLoad_Click(object sender, EventArgs e)
        {
            //場次查詢
            sql = reportSelect_SqlString("sceneNo", sceneNoText.Text);

            Refresh(sql);

        }
        private void nameLoad_Click(object sender, EventArgs e)
        {
            //姓名查詢
            sql = reportSelect_SqlString("name", nameText.Text);

            Refresh(sql);

        }
        private void idLoad_Click(object sender, EventArgs e)
        {
            //身分證字號查詢
            sql = reportSelect_SqlString("id", idText.Text);

            Refresh(sql);

        }
        private void projectNoLoad_Click(object sender, EventArgs e)
        {
            
            //專案查詢
            sql = reportSelect_SqlString("projectname", projectNoText.Text);

            Refresh(sql);

        }
        private void companyLoad_Click(object sender, EventArgs e)
        {
            //公司查詢
            sql = reportSelect_SqlString("company", companyText.Text);

            Refresh(sql);
        }
        private void projectNameLoad_Click(object sender, EventArgs e)
        {
            //公司專案查詢
            sql = reportSelect_SqlString("projectname2",projectNameText.Text.Split(':').GetValue(0).ToString());

            Refresh(sql);
        }
        private void deptNameLord_Click(object sender, EventArgs e)
        {
            //公司專案部門查詢
            sql = reportSelect_SqlString("deptName", projectNameText.Text.Split(':').GetValue(0).ToString());

            Refresh(sql);
        }

        
        ///回傳sql指令
        private string reportSelect_SqlString(string str,string data)
        {
            
            SqlString.sqlString1 sqlString1 = new SqlString.sqlString1();
            SqlString.sqlString2 sqlString2 = new SqlString.sqlString2();
            if (ReportSelect.Text == "勞工一般體格及健康檢查紀錄" || ReportSelect.Text== "從事供膳作業員工健康檢查紀錄表")
            {
                
                return sqlString1.getSqlString1(str, CheckPrintKind, data, getFilterString(), getSortString(), deptNameSelect.Text,ReportSelect.Text);
            }
            else if(ReportSelect.Text == "健康檢查紀錄表A3" || ReportSelect.Text == "健康檢查紀錄表A4")
            {
                return sqlString2.getSqlString2(str, CheckPrintKind, data, getFilterString(), getSortString(), deptNameSelect.Text, ReportSelect.Text,getPayKindString());
            }
            return sqlString1.getSqlString1(str, CheckPrintKind,data, getFilterString(), getSortString(), deptNameSelect.Text, ReportSelect.Text);
        }
        //專案關鍵字查詢
        private void project_search_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select `Name` from project WHERE  isfinish='1' and `Name` like '%" + projectNoText.Text + "%';", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataReader projNoRead = command.ExecuteReader();
                    while (projNoRead.Read())
                    {
                        listBox1.Items.Add(projNoRead.GetString(0));
                    }
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料量過大，請重新查詢\n" + ex.Data);

            }
        }
        //公司查詢
        private void company_search_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select Name from customer where Name like '%"+companyText.Text+"%';", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataReader companyRead = command.ExecuteReader();
                    while (companyRead.Read())
                    {
                        listBox2.Items.Add(companyRead.GetString(0));
                    }
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料量過大，請重新查詢\n" + ex.Data);

            }
        }
        //報表選擇
        private void ReportSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt1.Clear();
            dirNameLord.Enabled = false;
            switch (ReportSelect.Text)
            {
                
                case "勞工一般體格及健康檢查紀錄" :
                    reportViewer1.Reset();
                    reportViewer1.LocalReport.EnableExternalImages = true;
                    reportViewer1.LocalReport.ReportEmbeddedResource = "report2.Report.Report_General_health_check.rdlc";
                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dt1));
                    reportViewer1.RefreshReport();
                    this.Text = ReportSelect.Text + "   使用者:" + UserName;
                    //此項目無itextsharp批次
                    export_single.Enabled = true;
                    export_batch.Enabled = true;
                    A4_Item_add.Visible = false;
                    CheckPrintKind = "R00003";
                    break;
                case "從事供膳作業員工健康檢查紀錄表":
                    reportViewer1.Reset();
                    reportViewer1.LocalReport.EnableExternalImages = true;
                    reportViewer1.LocalReport.ReportEmbeddedResource = "report2.Report.Report_FeedStaff_health_check.rdlc";
                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dt1));
                    reportViewer1.RefreshReport();
                    this.Text = ReportSelect.Text + "   使用者:" + UserName;
                    //此項目無itextsharp批次
                    export_single.Enabled = true;
                    export_batch.Enabled = true;
                    A4_Item_add.Visible = false;
                    CheckPrintKind = "R00068";
                    break;
                case "健康檢查紀錄表A3":
                    reportViewer1.Reset();
                    reportViewer1.LocalReport.ReportEmbeddedResource = "report2.Report.Report_health_check_A3.rdlc";
                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dt1));
                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet2", dt1));
                    reportViewer1.RefreshReport();
                    this.Text = ReportSelect.Text + "   使用者:" + UserName;
                    //此項目尚未建置完成
                    export_single.Enabled = true;
                    export_batch.Enabled = true;
                    A4_Item_add.Visible = false;
                    CheckPrintKind = "R00001";
                    break;
                case "健康檢查紀錄表A4":
                    reportViewer1.Reset();
                    reportViewer1.LocalReport.ReportEmbeddedResource = "report2.Report.Reprot_health_check_A4.rdlc";
                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dt1));
                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet2", dt1));
                    reportViewer1.RefreshReport();
                    this.Text = ReportSelect.Text + "   使用者:" + UserName;
                    //此項目批次輸出已建置完成
                    export_single.Enabled = true;
                    export_batch.Enabled = true;
                    A4_Item_add.Visible = true;
                    CheckPrintKind = "";
                    break;
            }

        }
        //report refresh 報表刷新
        private async void Refresh(string sql)
        {
            close();

            //非同步線程
            await Task.Run(() => dt1 = populate(dt1, sql));
            reportViewer1.LocalReport.DataSources.Clear();
            //ReportDataSource rpts = ;
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dt1));
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet2", dt1));

            reportViewer1.RefreshReport();
            open();
        }
        //連線 抓取資料 填入dataset
        private DataTable populate(DataTable dt, string sql)
        {
            dt1.Clear();
            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                //string connectionstring = "server=localhost;database=backup1;uid=root;pwd=123qwe;";

                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand(sql, cnn);
                    command.CommandTimeout = 0;
                    MySqlDataAdapter myDA = new MySqlDataAdapter(command);
                    myDA.Fill(dt);
                    cnn.Close();
                    return dt;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料查詢有誤，請重新查詢\n" + ex.Data);
                return dt;
            }


        }
        
        //排序字串
        private string getSortString()
        {
            string SortString = "";
            if (DeptEmp_no.Checked)
            {
                SortString = "dept_name,CAST(hh.emp_no AS DECIMAL)";
            }
            else if (Emp_no.Checked)
            {
                SortString = "CAST(hh.emp_no AS DECIMAL)";
            }
            else if (name_sort.Checked)
            {
                SortString = "Name";
            }
            else if (deptName.Checked)
            {
                SortString = "dept_name,Name";
            }
            return SortString;
        }
        //篩選字串
        private string getFilterString()
        {
            string FilterString = "";
            //Form Data
            string From = dtFromDate.Value.Date.ToString("yyyy/MM/dd");//from date
            string To = dtToDate.Value.Date.ToString("yyyy/MM/dd");//To date
            //isemployee
            if (isemployee0.Checked == true && isemployee1.Checked == true)
            {
                FilterString = "";
            }
            else if (isemployee0.Checked == true && isemployee1.Checked == false)
            {
                FilterString = "and hh.isemployee = '0'";
            }
            else if (isemployee0.Checked == false && isemployee1.Checked == true)
            {
                FilterString = "and hh.isemployee = '1'";
            }
            else if (isemployee0.Checked == false && isemployee1.Checked == false)
            {
                FilterString = "";
            }
            //arraydate
            if (arraydate.Checked)
            {
                if (dtToDate.Value < dtFromDate.Value)
                    MessageBox.Show("起始日期不可大於結束日期");
                FilterString = "and hh.arraydate >= '" + From + "' and hh.arraydate  <= '" + To + "'" + FilterString;
            }
            return FilterString;
        }
        //自費加選 公費 判斷
        private string getPayKindString()
        {
            string str = "(hb.PayKind = '00' or hb.PayKind = '01')";

            if (ReportSelect.Text == "健康檢查紀錄表A4")
            {
                if (check_all.Checked)
                {
                    str = "(hb.PayKind = '00' or hb.PayKind = '01')";
                }
                else if (check_ganeral.Checked)
                {
                    str = " hb.PayKind = '00' ";
                }
                else if (check_plus_election.Checked)
                {
                    str = " hb.PayKind = '01' ";
                }
            }
            if (ReportSelect.Text == "健康檢查紀錄表A3")
            {
                str = "(hb.PayKind = '00' or (hb.PayKind = '01' and i.category <> 'B004M053') or (i.is_pe = '1' and i.category = 'B004M057'))";
            }

            return str;
        }
        //是否加入受檢日期判斷
        private void arraydate_CheckedChanged(object sender, EventArgs e)
        {
            if (arraydate.Checked == true)
            {
                btnLoad.Enabled = true;
                dtFromDate.Enabled = true;
                dtToDate.Enabled = true;
            }
            else
            {
                btnLoad.Enabled = false;
                dtFromDate.Enabled = false;
                dtToDate.Enabled = false;
            }
        }
        //單筆or批次選擇
        private void export_batch_CheckedChanged(object sender, EventArgs e)
        {
            if (export_batch.Checked)
            {
                checkBox_sid.Enabled = false;
                checkBox_name.Enabled = false;
                checkBox_deptName.Enabled = false;
                checkBox_empNo.Enabled = false;
                checkBox_pid.Enabled = false;
            }
            else
            {
                checkBox_sid.Enabled = true;
                checkBox_name.Enabled = true;
                checkBox_deptName.Enabled = true;
                checkBox_empNo.Enabled = true;
                checkBox_pid.Enabled = true;
            }
        }
        //關閉視窗
        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            DialogResult dr = MessageBox.Show(this, "確定退出？", "退出視窗通知", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr != DialogResult.Yes)
            {
                //3.Cancel 取得或設定數值，表示是否應該取消事件。
                e.Cancel = true;

            }
            else
            {
                this.Owner.Show();
            }
        }



        //private void projectNoText_TextChanged(object sender, EventArgs e)
        //{
        //    //AutoCompleteStringCollection acc = new AutoCompleteStringCollection();

        //    //projectNoText.AutoCompleteSource = AutoCompleteSource.CustomSource;
        //    //projectNoText.AutoCompleteMode = AutoCompleteMode.Suggest;
        //    //projectNoText.AutoCompleteCustomSource = acc;
        //}
        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            projectNoLoad.Enabled = true;
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            projectNoText.Text = listBox1.Text;
        }
        private void listBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            companyLoad.Enabled = true;
        }
        //點選公司清單 列出專案
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            companyText.Text = listBox2.Text;
            listBox3.Items.Clear();
            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select p.`Name`,p.`No`,c.`Name` as cName from project p INNER JOIN customer c on p.CustomerNo = c.`No` where c.`Name` like '%" + listBox2.Text+"%' ;", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataReader companyRead = command.ExecuteReader();
                    while (companyRead.Read())
                    {
                        listBox3.Items.Add(companyRead.GetString(0)+":"+ companyRead.GetString(1));
                    }
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料量過大，請重新查詢\n" + ex.Data);

            }
        }
        //專案名稱選擇
        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            projectNameText.Text = listBox3.Text;
            projectNameLoad.Enabled = true;
            
        }

        

        //根據專案名稱顯示部門(至下拉式清單)
        private void projectNameText_TextChanged(object sender, EventArgs e)
        {
            deptNameSelect.Items.Clear();
            deptNameSelect.Text = "";
            listBox3.Enabled = false;
            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("SELECT IFNULL(hh.dept_name,'') as dept_name from history_head hh INNER JOIN project p on hh.ProjectNo = p.`No` WHERE p.No = '" + projectNameText.Text.Split(':').GetValue(1) + "' GROUP BY hh.dept_name;", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataReader deptRead = command.ExecuteReader();
                    while (deptRead.Read())
                    {
                        deptNameSelect.Items.Add(deptRead.GetString(0));
                    }
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料為空值或是資料量過大，請重新查詢\n" + ex.Data);

            }
            listBox3.Enabled = true;
            deptNameLord.Enabled = true;
        }

        //輸出iTextSharp
        private void dirNameLord_Click(object sender, EventArgs e)
        {
            string dirPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), dirNameText.Text);
            string str = "";
            if (Directory.Exists(dirPath))
            {
                str = "The directory " + dirPath + " already exists.已經有此檔案，請重新命名";
                MessageBox.Show(str);
            }
            else
            {
                str = "The directory " + dirPath + " was created.已新建檔案";
                MessageBox.Show(str);
                Directory.CreateDirectory(dirPath);
                System.Diagnostics.Process.Start(dirPath);
                string dirName = "";
                dirName += (checkBox_sid.Checked) ? "true," : ",";
                dirName += (checkBox_name.Checked) ? "true," : ",";
                dirName += (checkBox_deptName.Checked) ? "true," : ",";
                dirName += (checkBox_empNo.Checked) ? "true," : ",";
                dirName += (checkBox_pid.Checked) ? "true," : ",";
                switch (ReportSelect.Text)
                {

                    case "勞工一般體格及健康檢查紀錄":
                        if (export_single.Checked)
                        {
                            iTS_Report_General_health_check General = new iTS_Report_General_health_check();
                            General.ExportPdf(dt1, dirPath, dirName);
                        }
                        else
                        {
                            iTS_Report_General_health_check_batch General = new iTS_Report_General_health_check_batch();
                            General.ExportPdf(dt1, dirPath, dirName);
                        }


                        break;
                    case "從事供膳作業員工健康檢查紀錄表":
                        if (export_single.Checked)
                        {
                            iTS_Report_FeedStaff_health_check feedstaff = new iTS_Report_FeedStaff_health_check();
                            feedstaff.ExportPdf(dt1, dirPath, dirName);
                        }
                        else
                        {
                            iTS_Report_FeedStaff_health_check_batch feedstaff = new iTS_Report_FeedStaff_health_check_batch();
                            feedstaff.ExportPdf(dt1, dirPath, dirName);
                        }
                            
                        break;
                    case "健康檢查紀錄表A4":
                        if (export_single.Checked)
                        {
                            iTS_Report_health_check_A4 health_check_A4 = new iTS_Report_health_check_A4();
                            health_check_A4.ExportPdf(dt1, dirPath, dirName);
                        }
                        else
                        {
                            iTS_Report_health_check_A4_batch health_check_A4 = new iTS_Report_health_check_A4_batch();
                            health_check_A4.ExportPdf(dt1, dirPath, dirName);
                        }

                        
                        break;
                    case "健康檢查紀錄表A3":
                        if (export_single.Checked)
                        {
                            iTS_Report_health_check_A3 health_check_A3 = new iTS_Report_health_check_A3();
                            health_check_A3.ExportPdf(dt1, dirPath, dirName);
                        }
                        else
                        {
                            iTS_Report_health_check_A3_batch health_check_A3 = new iTS_Report_health_check_A3_batch();
                            health_check_A3.ExportPdf(dt1, dirPath, dirName);
                        }
                        
                        break;
                }
            }
        }
        

        
    }
}
