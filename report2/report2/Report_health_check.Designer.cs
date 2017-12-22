using System.ComponentModel;

namespace report2
{
    partial class Report_health_check
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.DataTable1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DS = new report2.DS();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtToDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.dtFromDate = new System.Windows.Forms.DateTimePicker();
            this.btnLoad = new System.Windows.Forms.Button();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.label4 = new System.Windows.Forms.Label();
            this.sceneNoText = new System.Windows.Forms.TextBox();
            this.sceneNoLoad = new System.Windows.Forms.Button();
            this.nameLoad = new System.Windows.Forms.Button();
            this.nameText = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.idLoad = new System.Windows.Forms.Button();
            this.idText = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.DeptEmp_no = new System.Windows.Forms.RadioButton();
            this.Emp_no = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.isemployee0 = new System.Windows.Forms.CheckBox();
            this.isemployee1 = new System.Windows.Forms.CheckBox();
            this.arraydate = new System.Windows.Forms.CheckBox();
            this.ReportSelect = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.dirNameText = new System.Windows.Forms.TextBox();
            this.dirNameLord = new System.Windows.Forms.Button();
            this.name_sort = new System.Windows.Forms.RadioButton();
            this.deptName = new System.Windows.Forms.RadioButton();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.company_search = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.companyLoad = new System.Windows.Forms.Button();
            this.companyText = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.projectNameLoad = new System.Windows.Forms.Button();
            this.projectNameText = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.listBox3 = new System.Windows.Forms.ListBox();
            this.專案查詢 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.project_search = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.projectNoLoad = new System.Windows.Forms.Button();
            this.projectNoText = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.公司專案查詢 = new System.Windows.Forms.GroupBox();
            this.deptNameLord = new System.Windows.Forms.Button();
            this.deptNameSelect = new System.Windows.Forms.ComboBox();
            this.checkBox_sid = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.export_batch = new System.Windows.Forms.RadioButton();
            this.export_single = new System.Windows.Forms.RadioButton();
            this.checkBox_pid = new System.Windows.Forms.CheckBox();
            this.checkBox_empNo = new System.Windows.Forms.CheckBox();
            this.checkBox_deptName = new System.Windows.Forms.CheckBox();
            this.checkBox_name = new System.Windows.Forms.CheckBox();
            this.A4_Item_add = new System.Windows.Forms.GroupBox();
            this.check_all = new System.Windows.Forms.RadioButton();
            this.check_plus_election = new System.Windows.Forms.RadioButton();
            this.check_ganeral = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.DataTable1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DS)).BeginInit();
            this.專案查詢.SuspendLayout();
            this.公司專案查詢.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.A4_Item_add.SuspendLayout();
            this.SuspendLayout();
            // 
            // DataTable1BindingSource
            // 
            this.DataTable1BindingSource.DataMember = "DataTable1";
            this.DataTable1BindingSource.DataSource = this.DS;
            // 
            // DS
            // 
            this.DS.DataSetName = "DS";
            this.DS.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(12, 53);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 16);
            this.label3.TabIndex = 23;
            this.label3.Text = "報到日期";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(196, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 16);
            this.label2.TabIndex = 22;
            this.label2.Text = "to date";
            // 
            // dtToDate
            // 
            this.dtToDate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.dtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtToDate.Location = new System.Drawing.Point(252, 69);
            this.dtToDate.Name = "dtToDate";
            this.dtToDate.Size = new System.Drawing.Size(100, 27);
            this.dtToDate.TabIndex = 21;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(12, 76);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 16);
            this.label1.TabIndex = 20;
            this.label1.Text = "from date";
            // 
            // dtFromDate
            // 
            this.dtFromDate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.dtFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtFromDate.Location = new System.Drawing.Point(90, 69);
            this.dtFromDate.Name = "dtFromDate";
            this.dtFromDate.Size = new System.Drawing.Size(100, 27);
            this.dtFromDate.TabIndex = 19;
            this.dtFromDate.Value = new System.DateTime(2017, 10, 13, 10, 42, 59, 0);
            // 
            // btnLoad
            // 
            this.btnLoad.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnLoad.Location = new System.Drawing.Point(175, 99);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(176, 31);
            this.btnLoad.TabIndex = 18;
            this.btnLoad.Text = "只查詢日期";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // reportViewer1
            // 
            reportDataSource1.Name = "DataSet1";
            reportDataSource1.Value = this.DataTable1BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.EnableExternalImages = true;
            this.reportViewer1.LocalReport.EnableHyperlinks = true;
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "report2.Report.Report_General_health_check.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(5, 442);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1347, 491);
            this.reportViewer1.TabIndex = 24;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(12, 137);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 16);
            this.label4.TabIndex = 25;
            this.label4.Text = "場次查詢";
            // 
            // sceneNoText
            // 
            this.sceneNoText.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.sceneNoText.Location = new System.Drawing.Point(15, 159);
            this.sceneNoText.Name = "sceneNoText";
            this.sceneNoText.Size = new System.Drawing.Size(155, 27);
            this.sceneNoText.TabIndex = 26;
            // 
            // sceneNoLoad
            // 
            this.sceneNoLoad.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.sceneNoLoad.Location = new System.Drawing.Point(176, 159);
            this.sceneNoLoad.Name = "sceneNoLoad";
            this.sceneNoLoad.Size = new System.Drawing.Size(81, 27);
            this.sceneNoLoad.TabIndex = 27;
            this.sceneNoLoad.Text = "&場次查詢";
            this.sceneNoLoad.UseVisualStyleBackColor = true;
            this.sceneNoLoad.Click += new System.EventHandler(this.sceneNoLoad_Click);
            // 
            // nameLoad
            // 
            this.nameLoad.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.nameLoad.Location = new System.Drawing.Point(176, 214);
            this.nameLoad.Name = "nameLoad";
            this.nameLoad.Size = new System.Drawing.Size(81, 27);
            this.nameLoad.TabIndex = 30;
            this.nameLoad.Text = "&姓名查詢";
            this.nameLoad.UseVisualStyleBackColor = true;
            this.nameLoad.Click += new System.EventHandler(this.nameLoad_Click);
            // 
            // nameText
            // 
            this.nameText.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.nameText.Location = new System.Drawing.Point(15, 214);
            this.nameText.Name = "nameText";
            this.nameText.Size = new System.Drawing.Size(155, 27);
            this.nameText.TabIndex = 29;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(12, 195);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 16);
            this.label5.TabIndex = 28;
            this.label5.Text = "姓名查詢";
            // 
            // idLoad
            // 
            this.idLoad.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.idLoad.Location = new System.Drawing.Point(176, 270);
            this.idLoad.Name = "idLoad";
            this.idLoad.Size = new System.Drawing.Size(102, 27);
            this.idLoad.TabIndex = 33;
            this.idLoad.Text = "&身分證查詢";
            this.idLoad.UseVisualStyleBackColor = true;
            this.idLoad.Click += new System.EventHandler(this.idLoad_Click);
            // 
            // idText
            // 
            this.idText.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.idText.Location = new System.Drawing.Point(15, 270);
            this.idText.Name = "idText";
            this.idText.Size = new System.Drawing.Size(155, 27);
            this.idText.TabIndex = 32;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(12, 250);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(88, 16);
            this.label6.TabIndex = 31;
            this.label6.Text = "身分證查詢";
            // 
            // DeptEmp_no
            // 
            this.DeptEmp_no.AutoSize = true;
            this.DeptEmp_no.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.DeptEmp_no.Location = new System.Drawing.Point(133, 337);
            this.DeptEmp_no.Margin = new System.Windows.Forms.Padding(2);
            this.DeptEmp_no.Name = "DeptEmp_no";
            this.DeptEmp_no.Size = new System.Drawing.Size(122, 20);
            this.DeptEmp_no.TabIndex = 38;
            this.DeptEmp_no.TabStop = true;
            this.DeptEmp_no.Text = "部門工號排序";
            this.DeptEmp_no.UseVisualStyleBackColor = true;
            // 
            // Emp_no
            // 
            this.Emp_no.AutoSize = true;
            this.Emp_no.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Emp_no.Location = new System.Drawing.Point(133, 361);
            this.Emp_no.Margin = new System.Windows.Forms.Padding(2);
            this.Emp_no.Name = "Emp_no";
            this.Emp_no.Size = new System.Drawing.Size(90, 20);
            this.Emp_no.TabIndex = 39;
            this.Emp_no.TabStop = true;
            this.Emp_no.Text = "工號排序";
            this.Emp_no.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label8.Location = new System.Drawing.Point(130, 318);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(40, 16);
            this.label8.TabIndex = 40;
            this.label8.Text = "排序";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label9.Location = new System.Drawing.Point(11, 318);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(104, 16);
            this.label9.TabIndex = 41;
            this.label9.Text = "員工眷屬篩選";
            // 
            // isemployee0
            // 
            this.isemployee0.AutoSize = true;
            this.isemployee0.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.isemployee0.Location = new System.Drawing.Point(15, 338);
            this.isemployee0.Margin = new System.Windows.Forms.Padding(2);
            this.isemployee0.Name = "isemployee0";
            this.isemployee0.Size = new System.Drawing.Size(59, 20);
            this.isemployee0.TabIndex = 42;
            this.isemployee0.Text = "員工";
            this.isemployee0.UseVisualStyleBackColor = true;
            // 
            // isemployee1
            // 
            this.isemployee1.AutoSize = true;
            this.isemployee1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.isemployee1.Location = new System.Drawing.Point(15, 362);
            this.isemployee1.Margin = new System.Windows.Forms.Padding(2);
            this.isemployee1.Name = "isemployee1";
            this.isemployee1.Size = new System.Drawing.Size(59, 20);
            this.isemployee1.TabIndex = 43;
            this.isemployee1.Text = "眷屬";
            this.isemployee1.UseVisualStyleBackColor = true;
            // 
            // arraydate
            // 
            this.arraydate.AutoSize = true;
            this.arraydate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.arraydate.Location = new System.Drawing.Point(14, 105);
            this.arraydate.Margin = new System.Windows.Forms.Padding(2);
            this.arraydate.Name = "arraydate";
            this.arraydate.Size = new System.Drawing.Size(155, 20);
            this.arraydate.TabIndex = 45;
            this.arraydate.Text = "加上報到日期查詢";
            this.arraydate.UseVisualStyleBackColor = true;
            this.arraydate.CheckedChanged += new System.EventHandler(this.arraydate_CheckedChanged);
            // 
            // ReportSelect
            // 
            this.ReportSelect.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.ReportSelect.FormattingEnabled = true;
            this.ReportSelect.Items.AddRange(new object[] {
            "勞工一般體格及健康檢查紀錄",
            "從事供膳作業員工健康檢查紀錄表",
            "健康檢查紀錄表A3",
            "健康檢查紀錄表A4"});
            this.ReportSelect.Location = new System.Drawing.Point(90, 22);
            this.ReportSelect.Name = "ReportSelect";
            this.ReportSelect.Size = new System.Drawing.Size(246, 24);
            this.ReportSelect.TabIndex = 46;
            this.ReportSelect.Text = "勞工一般體格及健康檢查紀錄";
            this.ReportSelect.SelectedIndexChanged += new System.EventHandler(this.ReportSelect_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label10.Location = new System.Drawing.Point(12, 25);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(72, 16);
            this.label10.TabIndex = 47;
            this.label10.Text = "選擇報表";
            // 
            // dirNameText
            // 
            this.dirNameText.Location = new System.Drawing.Point(6, 176);
            this.dirNameText.Name = "dirNameText";
            this.dirNameText.Size = new System.Drawing.Size(263, 27);
            this.dirNameText.TabIndex = 53;
            // 
            // dirNameLord
            // 
            this.dirNameLord.Location = new System.Drawing.Point(6, 209);
            this.dirNameLord.Name = "dirNameLord";
            this.dirNameLord.Size = new System.Drawing.Size(263, 39);
            this.dirNameLord.TabIndex = 54;
            this.dirNameLord.Text = "匯出檔案到桌面";
            this.dirNameLord.UseVisualStyleBackColor = true;
            this.dirNameLord.Click += new System.EventHandler(this.dirNameLord_Click);
            // 
            // name_sort
            // 
            this.name_sort.AutoSize = true;
            this.name_sort.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.name_sort.Location = new System.Drawing.Point(133, 385);
            this.name_sort.Margin = new System.Windows.Forms.Padding(2);
            this.name_sort.Name = "name_sort";
            this.name_sort.Size = new System.Drawing.Size(90, 20);
            this.name_sort.TabIndex = 55;
            this.name_sort.TabStop = true;
            this.name_sort.Text = "姓名排序";
            this.name_sort.UseVisualStyleBackColor = true;
            // 
            // deptName
            // 
            this.deptName.AutoSize = true;
            this.deptName.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.deptName.Location = new System.Drawing.Point(133, 409);
            this.deptName.Margin = new System.Windows.Forms.Padding(2);
            this.deptName.Name = "deptName";
            this.deptName.Size = new System.Drawing.Size(122, 20);
            this.deptName.TabIndex = 56;
            this.deptName.TabStop = true;
            this.deptName.Text = "部門姓名排序";
            this.deptName.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(3, 157);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(124, 16);
            this.label12.TabIndex = 58;
            this.label12.Text = "填寫資料夾名稱:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 65);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(152, 16);
            this.label13.TabIndex = 70;
            this.label13.Text = "公司關鍵字查詢列表";
            // 
            // company_search
            // 
            this.company_search.Location = new System.Drawing.Point(258, 17);
            this.company_search.Margin = new System.Windows.Forms.Padding(2);
            this.company_search.Name = "company_search";
            this.company_search.Size = new System.Drawing.Size(102, 26);
            this.company_search.TabIndex = 69;
            this.company_search.Text = "關鍵字查詢";
            this.company_search.UseVisualStyleBackColor = true;
            this.company_search.Click += new System.EventHandler(this.company_search_Click);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 16;
            this.listBox2.Location = new System.Drawing.Point(8, 85);
            this.listBox2.Margin = new System.Windows.Forms.Padding(2);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(352, 68);
            this.listBox2.TabIndex = 68;
            this.listBox2.SelectedIndexChanged += new System.EventHandler(this.listBox2_SelectedIndexChanged);
            this.listBox2.SelectedValueChanged += new System.EventHandler(this.listBox2_SelectedValueChanged);
            // 
            // companyLoad
            // 
            this.companyLoad.Location = new System.Drawing.Point(258, 48);
            this.companyLoad.Name = "companyLoad";
            this.companyLoad.Size = new System.Drawing.Size(102, 30);
            this.companyLoad.TabIndex = 67;
            this.companyLoad.Text = "公司查詢";
            this.companyLoad.UseVisualStyleBackColor = true;
            this.companyLoad.Click += new System.EventHandler(this.companyLoad_Click);
            // 
            // companyText
            // 
            this.companyText.Location = new System.Drawing.Point(116, 17);
            this.companyText.Name = "companyText";
            this.companyText.Size = new System.Drawing.Size(137, 27);
            this.companyText.TabIndex = 66;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 27);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(104, 16);
            this.label14.TabIndex = 65;
            this.label14.Text = "公司名稱查詢";
            // 
            // projectNameLoad
            // 
            this.projectNameLoad.Location = new System.Drawing.Point(258, 260);
            this.projectNameLoad.Margin = new System.Windows.Forms.Padding(2);
            this.projectNameLoad.Name = "projectNameLoad";
            this.projectNameLoad.Size = new System.Drawing.Size(102, 28);
            this.projectNameLoad.TabIndex = 73;
            this.projectNameLoad.Text = "專案查詢";
            this.projectNameLoad.UseVisualStyleBackColor = true;
            this.projectNameLoad.Click += new System.EventHandler(this.projectNameLoad_Click);
            // 
            // projectNameText
            // 
            this.projectNameText.Enabled = false;
            this.projectNameText.Location = new System.Drawing.Point(9, 261);
            this.projectNameText.Name = "projectNameText";
            this.projectNameText.Size = new System.Drawing.Size(244, 27);
            this.projectNameText.TabIndex = 72;
            this.projectNameText.TextChanged += new System.EventHandler(this.projectNameText_TextChanged);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(6, 155);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(136, 16);
            this.label15.TabIndex = 71;
            this.label15.Text = "公司專案查詢列表";
            // 
            // listBox3
            // 
            this.listBox3.FormattingEnabled = true;
            this.listBox3.ItemHeight = 16;
            this.listBox3.Location = new System.Drawing.Point(9, 180);
            this.listBox3.Margin = new System.Windows.Forms.Padding(2);
            this.listBox3.Name = "listBox3";
            this.listBox3.Size = new System.Drawing.Size(351, 68);
            this.listBox3.TabIndex = 74;
            this.listBox3.SelectedIndexChanged += new System.EventHandler(this.listBox3_SelectedIndexChanged);
            // 
            // 專案查詢
            // 
            this.專案查詢.Controls.Add(this.label11);
            this.專案查詢.Controls.Add(this.project_search);
            this.專案查詢.Controls.Add(this.listBox1);
            this.專案查詢.Controls.Add(this.projectNoLoad);
            this.專案查詢.Controls.Add(this.projectNoText);
            this.專案查詢.Controls.Add(this.label7);
            this.專案查詢.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.專案查詢.Location = new System.Drawing.Point(358, 17);
            this.專案查詢.Name = "專案查詢";
            this.專案查詢.Size = new System.Drawing.Size(315, 258);
            this.專案查詢.TabIndex = 75;
            this.專案查詢.TabStop = false;
            this.專案查詢.Text = "專案查詢";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(6, 84);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(152, 16);
            this.label11.TabIndex = 56;
            this.label11.Text = "專案關鍵字查詢列表";
            // 
            // project_search
            // 
            this.project_search.Location = new System.Drawing.Point(206, 46);
            this.project_search.Margin = new System.Windows.Forms.Padding(2);
            this.project_search.Name = "project_search";
            this.project_search.Size = new System.Drawing.Size(102, 28);
            this.project_search.TabIndex = 55;
            this.project_search.Text = "關鍵字查詢";
            this.project_search.UseVisualStyleBackColor = true;
            this.project_search.Click += new System.EventHandler(this.project_search_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Location = new System.Drawing.Point(9, 111);
            this.listBox1.Margin = new System.Windows.Forms.Padding(2);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(301, 132);
            this.listBox1.TabIndex = 54;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            this.listBox1.SelectedValueChanged += new System.EventHandler(this.listBox1_SelectedValueChanged);
            // 
            // projectNoLoad
            // 
            this.projectNoLoad.Location = new System.Drawing.Point(207, 79);
            this.projectNoLoad.Name = "projectNoLoad";
            this.projectNoLoad.Size = new System.Drawing.Size(102, 27);
            this.projectNoLoad.TabIndex = 53;
            this.projectNoLoad.Text = "&專案查詢";
            this.projectNoLoad.UseVisualStyleBackColor = true;
            this.projectNoLoad.Click += new System.EventHandler(this.projectNoLoad_Click);
            // 
            // projectNoText
            // 
            this.projectNoText.Location = new System.Drawing.Point(9, 46);
            this.projectNoText.Name = "projectNoText";
            this.projectNoText.Size = new System.Drawing.Size(192, 27);
            this.projectNoText.TabIndex = 52;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(108, 16);
            this.label7.TabIndex = 51;
            this.label7.Text = "專案名稱查詢:";
            // 
            // 公司專案查詢
            // 
            this.公司專案查詢.Controls.Add(this.deptNameLord);
            this.公司專案查詢.Controls.Add(this.deptNameSelect);
            this.公司專案查詢.Controls.Add(this.label14);
            this.公司專案查詢.Controls.Add(this.companyText);
            this.公司專案查詢.Controls.Add(this.listBox3);
            this.公司專案查詢.Controls.Add(this.companyLoad);
            this.公司專案查詢.Controls.Add(this.projectNameLoad);
            this.公司專案查詢.Controls.Add(this.listBox2);
            this.公司專案查詢.Controls.Add(this.projectNameText);
            this.公司專案查詢.Controls.Add(this.company_search);
            this.公司專案查詢.Controls.Add(this.label15);
            this.公司專案查詢.Controls.Add(this.label13);
            this.公司專案查詢.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.公司專案查詢.Location = new System.Drawing.Point(679, 17);
            this.公司專案查詢.Name = "公司專案查詢";
            this.公司專案查詢.Size = new System.Drawing.Size(365, 350);
            this.公司專案查詢.TabIndex = 76;
            this.公司專案查詢.TabStop = false;
            this.公司專案查詢.Text = "公司專案查詢";
            // 
            // deptNameLord
            // 
            this.deptNameLord.Location = new System.Drawing.Point(257, 299);
            this.deptNameLord.Name = "deptNameLord";
            this.deptNameLord.Size = new System.Drawing.Size(102, 26);
            this.deptNameLord.TabIndex = 76;
            this.deptNameLord.Text = "部門查詢";
            this.deptNameLord.UseVisualStyleBackColor = true;
            this.deptNameLord.Click += new System.EventHandler(this.deptNameLord_Click);
            // 
            // deptNameSelect
            // 
            this.deptNameSelect.FormattingEnabled = true;
            this.deptNameSelect.Location = new System.Drawing.Point(9, 301);
            this.deptNameSelect.Name = "deptNameSelect";
            this.deptNameSelect.Size = new System.Drawing.Size(242, 24);
            this.deptNameSelect.TabIndex = 75;
            // 
            // checkBox_sid
            // 
            this.checkBox_sid.AutoSize = true;
            this.checkBox_sid.Location = new System.Drawing.Point(14, 88);
            this.checkBox_sid.Name = "checkBox_sid";
            this.checkBox_sid.Size = new System.Drawing.Size(103, 20);
            this.checkBox_sid.TabIndex = 77;
            this.checkBox_sid.Text = "1.體檢編號";
            this.checkBox_sid.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.export_batch);
            this.groupBox1.Controls.Add(this.export_single);
            this.groupBox1.Controls.Add(this.checkBox_pid);
            this.groupBox1.Controls.Add(this.checkBox_empNo);
            this.groupBox1.Controls.Add(this.checkBox_deptName);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.checkBox_name);
            this.groupBox1.Controls.Add(this.checkBox_sid);
            this.groupBox1.Controls.Add(this.dirNameLord);
            this.groupBox1.Controls.Add(this.dirNameText);
            this.groupBox1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.groupBox1.Location = new System.Drawing.Point(1050, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(279, 258);
            this.groupBox1.TabIndex = 78;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "匯出資料夾";
            // 
            // export_batch
            // 
            this.export_batch.AutoSize = true;
            this.export_batch.Location = new System.Drawing.Point(14, 50);
            this.export_batch.Name = "export_batch";
            this.export_batch.Size = new System.Drawing.Size(90, 20);
            this.export_batch.TabIndex = 83;
            this.export_batch.Text = "批次輸出";
            this.export_batch.UseVisualStyleBackColor = true;
            this.export_batch.CheckedChanged += new System.EventHandler(this.export_batch_CheckedChanged);
            // 
            // export_single
            // 
            this.export_single.AutoSize = true;
            this.export_single.Checked = true;
            this.export_single.Location = new System.Drawing.Point(14, 27);
            this.export_single.Name = "export_single";
            this.export_single.Size = new System.Drawing.Size(90, 20);
            this.export_single.TabIndex = 82;
            this.export_single.TabStop = true;
            this.export_single.Text = "單筆輸出";
            this.export_single.UseVisualStyleBackColor = true;
            // 
            // checkBox_pid
            // 
            this.checkBox_pid.AutoSize = true;
            this.checkBox_pid.Location = new System.Drawing.Point(14, 132);
            this.checkBox_pid.Name = "checkBox_pid";
            this.checkBox_pid.Size = new System.Drawing.Size(119, 20);
            this.checkBox_pid.TabIndex = 81;
            this.checkBox_pid.Text = "5.身分證字號";
            this.checkBox_pid.UseVisualStyleBackColor = true;
            // 
            // checkBox_empNo
            // 
            this.checkBox_empNo.AutoSize = true;
            this.checkBox_empNo.Location = new System.Drawing.Point(132, 110);
            this.checkBox_empNo.Name = "checkBox_empNo";
            this.checkBox_empNo.Size = new System.Drawing.Size(71, 20);
            this.checkBox_empNo.TabIndex = 80;
            this.checkBox_empNo.Text = "4.工號";
            this.checkBox_empNo.UseVisualStyleBackColor = true;
            // 
            // checkBox_deptName
            // 
            this.checkBox_deptName.AutoSize = true;
            this.checkBox_deptName.Location = new System.Drawing.Point(14, 110);
            this.checkBox_deptName.Name = "checkBox_deptName";
            this.checkBox_deptName.Size = new System.Drawing.Size(103, 20);
            this.checkBox_deptName.TabIndex = 79;
            this.checkBox_deptName.Text = "3.部門名稱";
            this.checkBox_deptName.UseVisualStyleBackColor = true;
            // 
            // checkBox_name
            // 
            this.checkBox_name.AutoSize = true;
            this.checkBox_name.Location = new System.Drawing.Point(132, 88);
            this.checkBox_name.Name = "checkBox_name";
            this.checkBox_name.Size = new System.Drawing.Size(71, 20);
            this.checkBox_name.TabIndex = 78;
            this.checkBox_name.Text = "2.姓名";
            this.checkBox_name.UseVisualStyleBackColor = true;
            // 
            // A4_Item_add
            // 
            this.A4_Item_add.Controls.Add(this.check_all);
            this.A4_Item_add.Controls.Add(this.check_plus_election);
            this.A4_Item_add.Controls.Add(this.check_ganeral);
            this.A4_Item_add.Location = new System.Drawing.Point(358, 282);
            this.A4_Item_add.Name = "A4_Item_add";
            this.A4_Item_add.Size = new System.Drawing.Size(98, 99);
            this.A4_Item_add.TabIndex = 79;
            this.A4_Item_add.TabStop = false;
            this.A4_Item_add.Text = "A4項目選擇";
            this.A4_Item_add.Visible = false;
            // 
            // check_all
            // 
            this.check_all.AutoSize = true;
            this.check_all.Checked = true;
            this.check_all.Location = new System.Drawing.Point(7, 67);
            this.check_all.Name = "check_all";
            this.check_all.Size = new System.Drawing.Size(77, 16);
            this.check_all.TabIndex = 2;
            this.check_all.TabStop = true;
            this.check_all.Text = "公費+自費";
            this.check_all.UseVisualStyleBackColor = true;
            // 
            // check_plus_election
            // 
            this.check_plus_election.AutoSize = true;
            this.check_plus_election.Location = new System.Drawing.Point(7, 45);
            this.check_plus_election.Name = "check_plus_election";
            this.check_plus_election.Size = new System.Drawing.Size(71, 16);
            this.check_plus_election.TabIndex = 1;
            this.check_plus_election.TabStop = true;
            this.check_plus_election.Text = "自費加選";
            this.check_plus_election.UseVisualStyleBackColor = true;
            // 
            // check_ganeral
            // 
            this.check_ganeral.AutoSize = true;
            this.check_ganeral.Location = new System.Drawing.Point(7, 22);
            this.check_ganeral.Name = "check_ganeral";
            this.check_ganeral.Size = new System.Drawing.Size(47, 16);
            this.check_ganeral.TabIndex = 0;
            this.check_ganeral.TabStop = true;
            this.check_ganeral.Text = "公費";
            this.check_ganeral.UseVisualStyleBackColor = true;
            // 
            // Report_health_check
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1364, 950);
            this.Controls.Add(this.A4_Item_add);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.公司專案查詢);
            this.Controls.Add(this.專案查詢);
            this.Controls.Add(this.deptName);
            this.Controls.Add(this.name_sort);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.ReportSelect);
            this.Controls.Add(this.arraydate);
            this.Controls.Add(this.isemployee1);
            this.Controls.Add(this.isemployee0);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.Emp_no);
            this.Controls.Add(this.DeptEmp_no);
            this.Controls.Add(this.idLoad);
            this.Controls.Add(this.idText);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.nameLoad);
            this.Controls.Add(this.nameText);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.sceneNoLoad);
            this.Controls.Add(this.sceneNoText);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtToDate);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtFromDate);
            this.Controls.Add(this.btnLoad);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Report_health_check";
            this.Text = "勞工一般體格及健康檢查紀錄";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.Form1_Closing);
            ((System.ComponentModel.ISupportInitialize)(this.DataTable1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DS)).EndInit();
            this.專案查詢.ResumeLayout(false);
            this.專案查詢.PerformLayout();
            this.公司專案查詢.ResumeLayout(false);
            this.公司專案查詢.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.A4_Item_add.ResumeLayout(false);
            this.A4_Item_add.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtToDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtFromDate;
        private System.Windows.Forms.Button btnLoad;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource DataTable1BindingSource;
        private DS DS;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox sceneNoText;
        private System.Windows.Forms.Button sceneNoLoad;
        private System.Windows.Forms.Button nameLoad;
        private System.Windows.Forms.TextBox nameText;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button idLoad;
        private System.Windows.Forms.TextBox idText;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.RadioButton DeptEmp_no;
        private System.Windows.Forms.RadioButton Emp_no;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox isemployee0;
        private System.Windows.Forms.CheckBox isemployee1;
        private System.Windows.Forms.CheckBox arraydate;
        private System.Windows.Forms.ComboBox ReportSelect;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox dirNameText;
        private System.Windows.Forms.Button dirNameLord;
        private System.Windows.Forms.RadioButton name_sort;
        private System.Windows.Forms.RadioButton deptName;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button company_search;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Button companyLoad;
        private System.Windows.Forms.TextBox companyText;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button projectNameLoad;
        private System.Windows.Forms.TextBox projectNameText;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ListBox listBox3;
        private System.Windows.Forms.GroupBox 專案查詢;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button project_search;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button projectNoLoad;
        private System.Windows.Forms.TextBox projectNoText;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox 公司專案查詢;
        private System.Windows.Forms.Button deptNameLord;
        private System.Windows.Forms.ComboBox deptNameSelect;
        private System.Windows.Forms.CheckBox checkBox_sid;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBox_pid;
        private System.Windows.Forms.CheckBox checkBox_empNo;
        private System.Windows.Forms.CheckBox checkBox_deptName;
        private System.Windows.Forms.CheckBox checkBox_name;
        private System.Windows.Forms.RadioButton export_batch;
        private System.Windows.Forms.RadioButton export_single;
        private System.Windows.Forms.GroupBox A4_Item_add;
        private System.Windows.Forms.RadioButton check_all;
        private System.Windows.Forms.RadioButton check_plus_election;
        private System.Windows.Forms.RadioButton check_ganeral;
    }
}

