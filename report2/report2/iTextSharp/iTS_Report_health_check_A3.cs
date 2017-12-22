using iTextSharp.text;
using iTextSharp.text.pdf;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace report2
{
    class iTS_Report_health_check_A3
    {

        BaseFont bfChinese = BaseFont.CreateFont(@"C:\Windows\Fonts\mingliu.ttc,1", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        Font chFont;
        Font chFont_UnderLine;
        Font chFont_blue;
        Font chFont_red, chFont_red_small,chFont_red_ssmall;
        Font chFont_header;
        Font chFont_small,chFont_ssmall, chFont_small_1, chFont_small_info, chFont_small_info_bold, chFont_small_bold;
        BaseColor Dblue = new BaseColor(89, 198, 209), Lblue = new BaseColor(199, 234, 251), blue = new BaseColor(30, 144, 255), Orange = new BaseColor(245, 130, 31), Dgreen = new BaseColor(224, 255, 193), Lgreen = new BaseColor(182, 220, 160);
        string cName, dept_name, name, pid, sex, birtyday, emp_no, sid, arraydate, category, category_tc, name_tc, qty, isunusual, id, m_normal_d, m_normal_u, unit, is_pe, is_inst, roomname, gtext,customerNo, payKind, main_name_tc, speciel_add;
        string group_id = "", group_category = "", doctor = "",doctor2 = "", qty2 = "", qty3 = "", isunusual2 = "", isunusual3 = "";
        DataTable dt_before1 = null, dt_before2 = null,dt_is_pe;
        string r500="", l500 = "", r1000 = "", l1000 = "", r2000 = "", l2000 = "", r3000 = "", l3000 = "", r4000 = "", l4000 = "", r6000 = "", l6000 = "", r8000 = "", l8000 = "";
        string D = "", F = "";
        public void ExportPdf(DataTable dt, string dirPath,string dirName)
        {
            int first = 0, last = 0;
            string str = "";
            group_id = dt.Rows[0]["id"].ToString();
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                id = dt.Rows[i]["id"].ToString();
                if (group_id != id || i == dt.Rows.Count - 1)
                {
                    group_id = id;

                    first = last;
                    last = i;
                    //最後一筆
                    if (i == dt.Rows.Count - 1)
                        last++;

                    pid = dt.Rows[first]["pid"].ToString();
                    arraydate = dt.Rows[first]["arraydate"].ToString();
                    customerNo = dt.Rows[first]["customerNo"].ToString();
                    DataTable dt1 = before_data(pid, arraydate, customerNo);
                    dt_before1 = null;
                    dt_before2 = null;
                    if (dt1.Rows.Count > 1)
                    {
                        dt_before1 = previous_data(dt1.Rows[1][0].ToString());
                        if (dt1.Rows.Count > 2)
                            dt_before2 = previous_data(dt1.Rows[2][0].ToString());
                    }
                    outputDocument(dt, dirPath, first, last, dirName);

                    //str = str +"\n"+ group_id + " " + id + " " + first + " " + last  ;
                }

            }

            //MessageBox.Show(str);
        }

        private void outputDocument(DataTable dt, string dirPath, int first, int last, string dirName)
        {
            //字型設定
            chFont = new Font(bfChinese, 12, Font.NORMAL, BaseColor.BLACK);
            chFont_UnderLine = new Font(bfChinese, 12, Font.UNDERLINE);
            chFont_blue = new Font(bfChinese, 10, Font.NORMAL, new BaseColor(0, 0, 240));
            chFont_red = new Font(bfChinese, 12, Font.NORMAL, BaseColor.RED);
            chFont_red_small = new Font(bfChinese, 8, Font.NORMAL, BaseColor.RED);
            chFont_red_ssmall = new Font(bfChinese, 10, Font.NORMAL, BaseColor.RED);
            chFont_header = new Font(bfChinese, 18, Font.BOLD , Dblue);
            chFont_small = new Font(bfChinese, 8, Font.NORMAL, BaseColor.BLACK);
            chFont_ssmall = new Font(bfChinese, 5, Font.NORMAL, BaseColor.BLACK);
            chFont_small_1 = new Font(bfChinese, 4, Font.NORMAL, Lgreen);
            chFont_small_info = new Font(bfChinese, 11, Font.NORMAL, BaseColor.BLACK);
            chFont_small_info_bold = new Font(bfChinese, 11, Font.BOLD, BaseColor.BLACK);
            chFont_small_bold = new Font(bfChinese, 9, Font.BOLD, BaseColor.BLACK);

            //-------------------------------------
            //dt.Rows[0][1]

            //名稱判斷
            string[] str = new string[5];
            str[0] = (dirName.Split(',').GetValue(0).ToString() == "true") ? dt.Rows[first]["sid"].ToString() : "";
            str[1] = (dirName.Split(',').GetValue(1).ToString() == "true") ? dt.Rows[first]["name"].ToString() : "";
            str[2] = (dirName.Split(',').GetValue(2).ToString() == "true") ? dt.Rows[first]["dept_name"].ToString() : "";
            str[3] = (dirName.Split(',').GetValue(3).ToString() == "true") ? dt.Rows[first]["emp_no"].ToString() : "";
            str[4] = (dirName.Split(',').GetValue(4).ToString() == "true") ? dt.Rows[first]["pid"].ToString() : "";

            

            string outputFile = Path.Combine(dirPath, str[0] + str[1] + str[2] + str[3] + str[4] + ".pdf");

            //Create a standard .Net FileStream for the file, setting various flags
            using (FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                //Create a new PDF document setting the size to A4
                using (Document doc1 = new Document(PageSize.A3.Rotate(), 5, 5, 5, 2))
                {
                    //Bind the PDF document to the FileStream using an iTextSharp PdfWriter
                    using (PdfWriter w = PdfWriter.GetInstance(doc1, fs))
                    {


                        doc1.Open();
                        //table1
                        PdfPTable table1 = new PdfPTable(new float[] {10,10,12,8 });
                        table1.DefaultCell.Border = 0;
                        table1.TotalWidth = 1150f;
                        table1.LockedWidth = true;
                        table1.HorizontalAlignment = Element.ALIGN_CENTER;

                        //檢查項目 左
                        PdfPTable table_l = new PdfPTable(new float[] { 100, 25, 25, 25 });
                        table_l.DefaultCell.FixedHeight = 15f;
                        table_l.DefaultCell.BorderColor = BaseColor.WHITE;
                        table_l.AddCell(CellFactory_line_title(new Phrase("檢查項目", chFont_small_bold),Dblue,BaseColor.WHITE));
                        table_l.AddCell(CellFactory_line_title(new Phrase("本次", chFont_small_bold), Dblue, Lblue));
                        table_l.AddCell(CellFactory_line_title(new Phrase("前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_l.AddCell(CellFactory_line_title(new Phrase("前前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        //檢查項目 右
                        PdfPTable table_r = new PdfPTable(new float[] { 100, 25, 25, 25 });
                        table_r.DefaultCell.FixedHeight = 15f;
                        table_r.DefaultCell.BorderColor = BaseColor.WHITE;
                        table_r.AddCell(CellFactory_line_title(new Phrase("檢查項目", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_r.AddCell(CellFactory_line_title(new Phrase("本次", chFont_small_bold), Dblue, Lblue));
                        table_r.AddCell(CellFactory_line_title(new Phrase("前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_r.AddCell(CellFactory_line_title(new Phrase("前前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        //儀器檢查
                        PdfPTable table_inst = new PdfPTable(new float[] { 1, 6 });
                        table_inst.DefaultCell.Border = 0;
                        table_inst.DefaultCell.FixedHeight = 15f;
                        table_inst.AddCell(CellFactory(new Phrase("儀器檢查", chFont_small_bold), Lgreen, 2));
                        //理學檢查
                        PdfPTable table_pe = new PdfPTable(new float[] { 2, 4 });
                        table_pe.DefaultCell.Border = 0;
                        table_pe.DefaultCell.FixedHeight = 15f;
                        table_pe.AddCell(CellFactory_line_title(new Phrase("理學檢查", chFont_small_bold), Dblue, BaseColor.WHITE, 2));

                        //檢查項目 左2
                        PdfPTable table_l2 = new PdfPTable(new float[] { 100, 25, 25, 25 });
                        table_l2.DefaultCell.FixedHeight = 15f;
                        table_l2.DefaultCell.BorderColor = BaseColor.WHITE;
                        table_l2.AddCell(CellFactory_line_title(new Phrase("檢查項目", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_l2.AddCell(CellFactory_line_title(new Phrase("本次", chFont_small_bold), Dblue, Lblue));
                        table_l2.AddCell(CellFactory_line_title(new Phrase("前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_l2.AddCell(CellFactory_line_title(new Phrase("前前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        //檢查項目 右2
                        PdfPTable table_r2 = new PdfPTable(new float[] { 100, 25, 25, 25 });
                        table_r2.DefaultCell.FixedHeight = 15f;
                        table_r2.DefaultCell.BorderColor = BaseColor.WHITE;
                        table_r2.AddCell(CellFactory_line_title(new Phrase("檢查項目", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_r2.AddCell(CellFactory_line_title(new Phrase("本次", chFont_small_bold), Dblue, Lblue));
                        table_r2.AddCell(CellFactory_line_title(new Phrase("前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        table_r2.AddCell(CellFactory_line_title(new Phrase("前前次", chFont_small_bold), Dblue, BaseColor.WHITE));
                        //儀器檢查2
                        PdfPTable table_inst2 = new PdfPTable(new float[] { 1, 6 });
                        table_inst2.DefaultCell.Border = 0;
                        table_inst2.DefaultCell.FixedHeight = 15f;
                        table_inst2.AddCell(CellFactory(new Phrase("儀器檢查", chFont_small_bold), Lgreen, 2));
                        //理學檢查2
                        PdfPTable table_pe2 = new PdfPTable(new float[] { 2, 4 });
                        table_pe2.DefaultCell.Border = 0;
                        table_pe2.DefaultCell.FixedHeight = 15f;
                        table_pe2.AddCell(CellFactory_line_title(new Phrase("理學檢查", chFont_small_bold),Dblue,BaseColor.WHITE, 2));

                        //頭
                        PdfPCell cell1 = new PdfPCell(new Phrase("健康檢查報告", chFont_header));
                        cell1.Colspan = 2;
                        cell1.FixedHeight = 25f;
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell1.PaddingTop = 0;
                        cell1.Border = 0;
                        
                        PdfPTable table2 = new PdfPTable(1);
                        PdfPTable table3 = new PdfPTable(1);
                        PdfPTable table4 = new PdfPTable(new float[] {3,2,2});
                        PdfPTable table5 = new PdfPTable(1);



                        int k = 1,j=1;
                        qty2 = ""; isunusual2 = "";
                        qty3 = ""; isunusual3 = "";
                        for (int i = first; i < last; i++)
                        {

                            if (dt_before1 != null)
                            {
                                foreach (DataRow before1 in dt_before1.Rows)
                                {
                                    if (dt.Rows[i]["item_iid"].ToString() == before1[3].ToString())
                                    {
                                        qty2 = before1[4].ToString();
                                        isunusual2 = before1[5].ToString();
                                        break;
                                    }
                                    else
                                    {
                                        qty2 = "";
                                        isunusual2 = "";
                                    }
                                }
                            }

                            if (dt_before2 != null)
                            {
                                foreach (DataRow before2 in dt_before2.Rows)
                                {
                                    if (dt.Rows[i]["item_iid"].ToString() == before2[3].ToString())
                                    {
                                        qty3 = before2[4].ToString();
                                        isunusual3 = before2[5].ToString();
                                        break;
                                    }
                                    else
                                    {
                                        qty3 = "";
                                        isunusual3 = "";
                                    }
                                }
                            }

                            id = dt.Rows[i]["id"].ToString();
                            cName = dt.Rows[i]["cName"].ToString();
                            dept_name = dt.Rows[i]["dept_name"].ToString();
                            name = dt.Rows[i]["name"].ToString();
                            pid = dt.Rows[i]["pid"].ToString();
                            sex = dt.Rows[i]["sex"].ToString();
                            birtyday = Convert.ToDateTime(dt.Rows[i]["birtyday"].ToString()).ToString("yyyy/MM/dd");
                            emp_no = dt.Rows[i]["emp_no"].ToString();
                            sid = dt.Rows[i]["sid"].ToString();
                            arraydate = Convert.ToDateTime(dt.Rows[i]["arraydate"].ToString()).ToString("yyyy/MM/dd");
                            name_tc = dt.Rows[i]["name_tc"].ToString();
                            qty = dt.Rows[i]["qty"].ToString();
                            isunusual = dt.Rows[i]["isunusual"].ToString();
                            category = dt.Rows[i]["category"].ToString();
                            category_tc = dt.Rows[i]["category_tc"].ToString();
                            m_normal_d = dt.Rows[i]["m_normal_d"].ToString();
                            m_normal_u = dt.Rows[i]["m_normal_u"].ToString();
                            unit = dt.Rows[i]["unit"].ToString();
                            is_inst = dt.Rows[i]["is_inst"].ToString();
                            is_pe = dt.Rows[i]["is_pe"].ToString();
                            roomname = dt.Rows[i]["roomname"].ToString();
                            gtext = dt.Rows[i]["gtext"].ToString();
                            payKind = dt.Rows[i]["PayKind"].ToString();
                            main_name_tc = dt.Rows[i]["main_name_tc"].ToString();
                            speciel_add = dt.Rows[i]["speciel_add"].ToString();
                            //Open the document for writing
                            //head===================================================================
                            if (i == first)
                            {
                                //頭

                                //table2
                                //既往病史 自覺症狀
                                is_DF_data(pid, arraydate);
                                table2.DefaultCell.Border = 0;
                                table2.AddCell(CellFactory(new Phrase("既往病史", chFont_small_bold), blue));
                                table2.AddCell(CellFactory(new Phrase(D, chFont_small), blue));
                                table2.AddCell(CellFactory(new Phrase("自覺症狀", chFont_small_bold), blue));
                                table2.AddCell(CellFactory(new Phrase(F, chFont_small), blue));
                                table2.AddCell(" ");

                                //table3
                                table3.DefaultCell.Border = 0;
                                table3.AddCell(CellFactory(new Phrase("主套餐："+ main_name_tc, chFont_small_info_bold), Orange));
                                table3.AddCell(CellFactory(new Phrase("身分證號碼：" + pid, chFont_small_info), Orange));
                                sex = (sex == "M")?"男":"女";
                                table3.AddCell(CellFactory(new Phrase("性別：" + sex + "   生日：" + birtyday, chFont_small_info), Orange));
                                table3.AddCell(CellFactory(new Phrase("體檢日期：" + arraydate, chFont_small_info), Orange));
                                table3.AddCell(CellFactory(new Phrase("特殊及加選項目："+ speciel_add, chFont_ssmall), Orange));
                                table3.AddCell(CellFactory(new Phrase(" ", chFont_small_info), Orange));

                                //table4
                                table4.DefaultCell.Border = 0;
                                table4.DefaultCell.BorderWidthTop = 0;
                                table4.AddCell(" "); table4.AddCell(" "); table4.AddCell(" ");
                                table4.AddCell(" "); table4.AddCell(" "); table4.AddCell(" ");
                                table4.AddCell(CellFactory(new Phrase("                  姓名：" + name, chFont), BaseColor.WHITE, 1));
                                table4.AddCell(CellFactory(new Phrase("        職稱：", chFont), BaseColor.WHITE, 2));
                                table4.AddCell(CellFactory(new Phrase("                  體檢編號：" + sid, chFont), BaseColor.WHITE, 1));
                                table4.AddCell(CellFactory(new Phrase("        工號：" + emp_no, chFont), BaseColor.WHITE, 2));
                                table4.AddCell(CellFactory(new Phrase("                  部門：" + dept_name, chFont), BaseColor.WHITE, 3));
                                table4.AddCell(CellFactory(new Phrase("                  受檢單位：" + cName, chFont), BaseColor.WHITE, 3));



                                //table5
                                table5.DefaultCell.Border = 0;
                                table5.DefaultCell.BorderWidthTop = 0;
                                Chunk table5_c1 = new Chunk(
                                    "                     檢查正常值範圍可能因受檢年度有些微差異，本分\n" +
                                    "                     報告正常值僅供本次結果參考，前次參考值請查閱\n" +
                                    "                     當年度體檢報告。檢驗結果位於正常範圍內，不帶\n" +
                                    "                     表一定沒有得到任何疾病，請持續定期健康檢查，\n" +
                                    "                     如有不適請至門診。此份檢驗報告為台北榮民總醫\n" +
                                    "                     院桃園分院健檢管理中心檢驗室所發出\n", chFont_small);
                                Chunk table5_c2 = new Chunk("                    *為加選項目 #為委託檢驗項目", chFont_red_ssmall);
                                Phrase table5_p1 = new Phrase();
                                table5_p1.Add(table5_c1);
                                table5_p1.Add(table5_c2);
                                table5.AddCell(" ");
                                table5.AddCell(" ");
                                table5.AddCell(table5_p1);


                                table1.AddCell(cell1);
                                table1.AddCell(CellFactory(" ",BaseColor.WHITE, 2));
                                table1.AddCell(table2);
                                table1.AddCell(table3);
                                table1.AddCell(table4);
                                table1.AddCell(table5);

                                //寫死 理學檢查
                                doctor = "";doctor2 = "";
                                dt_is_pe =  is_pe_data(pid, arraydate);
                                foreach (DataRow ispe in dt_is_pe.Rows)
                                {
                                    if (ispe[4].ToString() == "R000938")
                                    {
                                        doctor = ispe[3].ToString();
                                    }
                                    else if(ispe[4].ToString() == "R000910")
                                    {
                                        doctor2 = ispe[3].ToString();
                                    }else
                                    {
                                        table_pe.AddCell(CellFactory_line_content(new Phrase(ispe[2].ToString(), chFont_small),Dblue, BaseColor.WHITE));
                                        table_pe.AddCell(CellFactory_line_qty(ispe[3].ToString(), ispe[5].ToString(), BaseColor.WHITE));
                                        j++;
                                    }
                                }
                            }
                            //head end===================================================================
                            if(roomname == "聽力機")
                            {
                                if (name_tc == "左耳頻率500赫茲")
                                    l500 = qty;
                                if (name_tc == "左耳頻率1000赫茲")
                                    l1000 = qty;
                                if (name_tc == "左耳頻率2000赫茲")
                                    l2000 = qty;
                                if (name_tc == "左耳頻率3000赫茲")
                                    l3000 = qty;
                                if (name_tc == "左耳頻率4000赫茲")
                                    l4000 = qty;
                                if (name_tc == "左耳頻率6000赫茲")
                                    l6000 = qty;
                                if (name_tc == "左耳頻率8000赫茲")
                                    l8000 = qty;
                                if (name_tc == "右耳頻率500赫茲")
                                    r500 = qty;
                                if (name_tc == "右耳頻率1000赫茲")
                                    r1000 = qty;
                                if (name_tc == "右耳頻率2000赫茲")
                                    r2000 = qty;
                                if (name_tc == "右耳頻率3000赫茲")
                                    r3000 = qty;
                                if (name_tc == "右耳頻率4000赫茲")
                                    r4000 = qty;
                                if (name_tc == "右耳頻率6000赫茲")
                                    r6000 = qty;
                                if (name_tc == "右耳頻率8000赫茲")
                                    r8000 = qty;

                            }
                            //儀器判斷 || 特殊理學判斷
                            else if (is_inst != "True" || (is_pe == "True" && category == "B004M057"))
                            {
                                if(is_pe == "True" && category != "B004M057")
                                {
                                    break;
                                }
                                //第一頁 檢查項目 左
                                if (table_l.Rows.Count <=41)
                                {

                                    //檢查項目群組 頭
                                    if (group_category != category)
                                    {
                                        group_category = category;

                                        //檢查項目群組標題
                                        table_l.AddCell(CellFactory_line_content(new Phrase(category_tc, chFont_small_bold), Dblue, BaseColor.WHITE, 4));
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_l.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_l.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_l.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_l.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));


                                    }
                                    else
                                    {
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_l.AddCell(CellFactory_line_content(p1,Dblue, BaseColor.WHITE));
                                        table_l.AddCell(CellFactory_line_qty(qty, isunusual,Lblue));
                                        table_l.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_l.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));
                                    }

                                }
                                //第一頁 檢查項目 右
                                else if( table_r.Rows.Count<=41)
                                {
                                    //檢查項目群組 頭
                                    if (group_category != category)
                                    {
                                        group_category = category;

                                        //檢查項目群組標題
                                        table_r.AddCell(CellFactory_line_content(new Phrase(category_tc, chFont_small_bold), Dblue, BaseColor.WHITE, 4));
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Phrase("＊", chFont_red_small));
                                            p1.Add(new Phrase(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_r.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_r.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_r.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_r.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));

                                    }
                                    else
                                    {
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_r.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_r.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_r.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_r.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));
                                    }
                                }
                                //第二頁 檢查項目 左
                                else if(table_l2.Rows.Count <= 41)
                                {
                                    

                                    //檢查項目群組 頭
                                    if (group_category != category)
                                    {
                                        group_category = category;

                                        
                                        //檢查項目群組標題
                                        table_l2.AddCell(CellFactory_line_content(new Phrase(category_tc, chFont_small_bold), Dblue, BaseColor.WHITE, 4));
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_l2.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_l2.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_l2.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_l2.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));
                                        

                                    }
                                    else
                                    {
                                        
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_l2.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_l2.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_l2.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_l2.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));
                                    }

                                }else
                                {
                                    //檢查項目群組 頭
                                    if (group_category != category)
                                    {
                                        group_category = category;


                                        //檢查項目群組標題
                                        table_r2.AddCell(CellFactory_line_content(new Phrase(category_tc, chFont_small_bold), Dblue, BaseColor.WHITE, 4));
                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_r2.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_r2.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_r2.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_r2.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));
                                        

                                    }
                                    else
                                    {

                                        //檢查項目內容
                                        if (m_normal_d != "")
                                            name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                        Phrase p1 = new Phrase();
                                        if (payKind == "01")
                                        {
                                            p1.Add(new Chunk("＊", chFont_red_small));
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        else
                                        {
                                            p1.Add(new Chunk(name_tc, chFont_small));
                                        }
                                        table_r2.AddCell(CellFactory_line_content(p1, Dblue, BaseColor.WHITE));
                                        table_r2.AddCell(CellFactory_line_qty(qty, isunusual, Lblue));
                                        table_r2.AddCell(CellFactory_line_qty(qty2, isunusual2, BaseColor.WHITE));
                                        table_r2.AddCell(CellFactory_line_qty(qty3, isunusual3, BaseColor.WHITE));
                                    }
                                }


                                
                            }
                            //儀器
                            else
                            {
                                //第一頁
                                if(k <= 7)
                                {
                                    //檢查項目內容
                                    if (m_normal_d != "")
                                        name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                    Phrase name_tc_phrase = new Phrase();
                                    if (payKind == "01")
                                    {
                                        name_tc_phrase.Add(new Chunk("＊", chFont_red_small));
                                        name_tc_phrase.Add(new Chunk(name_tc, chFont_small_bold));
                                    }
                                    else
                                    {
                                        name_tc_phrase.Add(new Chunk(name_tc, chFont_small_bold));
                                    }
                                    table_inst.AddCell(CellFactory_line_color_green(name_tc_phrase, Lgreen, 2));

                                    table_inst.AddCell(CellFactory(" ", Lgreen, 1, 3));
                                    Chunk c1 = new Chunk("本次：", chFont_small);
                                    Chunk c2 = new Chunk(qty, (isunusual == "Y") ? chFont_red_small : chFont_small);
                                    Phrase p1 = new Phrase();
                                    p1.Add(c1);
                                    p1.Add(c2);
                                    table_inst.AddCell(CellFactory(p1, Lgreen));
                                    Chunk c3 = new Chunk("前次：", chFont_small);
                                    Chunk c4 = new Chunk(qty2, (isunusual2 == "Y") ? chFont_red_small : chFont_small);
                                    Phrase p2 = new Phrase();
                                    p2.Add(c3);
                                    p2.Add(c4);
                                    table_inst.AddCell(CellFactory(p2, Lgreen));
                                    Chunk c5 = new Chunk("前前次：", chFont_small);
                                    Chunk c6 = new Chunk(qty3, (isunusual3 == "Y") ? chFont_red_small : chFont_small);
                                    Phrase p3 = new Phrase();
                                    p3.Add(c5);
                                    p3.Add(c6);
                                    table_inst.AddCell(CellFactory(p3, Lgreen));
                                }else
                                //第二頁
                                {
                                    if (m_normal_d != "")
                                        name_tc += "<" + m_normal_d + "~" + m_normal_u + unit + ">";
                                    Phrase name_tc_phrase = new Phrase();
                                    if (payKind == "01")
                                    {
                                        name_tc_phrase.Add(new Chunk("＊", chFont_red_small));
                                        name_tc_phrase.Add(new Chunk(name_tc, chFont_small_bold));
                                    }
                                    else
                                    {
                                        name_tc_phrase.Add(new Chunk(name_tc, chFont_small_bold));
                                    }
                                    table_inst2.AddCell(CellFactory_line_color_green(name_tc_phrase, Lgreen, 2));

                                    table_inst2.AddCell(CellFactory(" ", Lgreen, 1, 3));
                                    Chunk c1 = new Chunk("本次：", chFont_small);
                                    Chunk c2 = new Chunk(qty, (isunusual == "Y") ? chFont_red_small : chFont_small);
                                    Phrase p1 = new Phrase();
                                    p1.Add(c1);
                                    p1.Add(c2);
                                    table_inst2.AddCell(CellFactory(p1, Lgreen));
                                    Chunk c3 = new Chunk("前次：", chFont_small);
                                    Chunk c4 = new Chunk(qty2, (isunusual2 == "Y") ? chFont_red_small : chFont_small);
                                    Phrase p2 = new Phrase();
                                    p2.Add(c3);
                                    p2.Add(c4);
                                    table_inst2.AddCell(CellFactory(p2, Lgreen));
                                    Chunk c5 = new Chunk("前前次：", chFont_small);
                                    Chunk c6 = new Chunk(qty3, (isunusual3 == "Y") ? chFont_red_small : chFont_small);
                                    Phrase p3 = new Phrase();
                                    p3.Add(c5);
                                    p3.Add(c6);
                                    table_inst2.AddCell(CellFactory(p3, Lgreen));
                                }
                                
                                k++;
                            }  
                        }


                        //第一欄
                        //補空格
                        if (table_l.Rows.Count < 41)
                        {
                            int m = table_l.Rows.Count;
                            for (int i = 0; i < 41 - m ; i++)
                            {
                                table_l.AddCell(CellFactory_line_content(" ",Dblue,BaseColor.WHITE, 4));
                            }
                        }
                        if (table_r.Rows.Count < 41)
                        {
                            int m = table_r.Rows.Count;
                            for (int i = 0; i < 41 - m; i++)
                            {
                                table_r.AddCell(CellFactory_line_content(" ", Dblue, BaseColor.WHITE, 4));
                            }
                        }
                        if (table_l2.Rows.Count < 41)
                        {
                            int m = table_l2.Rows.Count;
                            for (int i = 0; i < 41 - m; i++)
                            {
                                table_l2.AddCell(CellFactory_line_content(" ", Dblue, BaseColor.WHITE, 4));
                            }
                        }
                        if (table_r2.Rows.Count < 41)
                        {
                            int m = table_r2.Rows.Count;
                            for (int i = 0; i < 41 - m; i++)
                            {
                                table_r2.AddCell(CellFactory_line_content(" ", Dblue, BaseColor.WHITE, 4));
                            }
                        }
                        
                        table_l.AddCell(""); table_l.AddCell(""); table_l.AddCell(""); table_l.AddCell("");
                        table1.AddCell(table_l);
                        //第二欄
                        table_r.AddCell(""); table_r.AddCell(""); table_r.AddCell(""); table_r.AddCell("");
                        table1.AddCell(table_r);
                        //第三欄 儀器
                        //補空格
                        if (k <= 7)
                        {
                            for(int i = 0; i <= 7 - k; i++)
                            {
                                table_inst.AddCell(CellFactory_line_color_green(new Phrase(" ",chFont_small), Lgreen, 2));
                                table_inst.AddCell(CellFactory(new Phrase(" ",chFont_small), Lgreen, 1, 3));
                                table_inst.AddCell(CellFactory(new Phrase("本次", chFont_small), Lgreen));
                                table_inst.AddCell(CellFactory(new Phrase("前次", chFont_small), Lgreen));
                                table_inst.AddCell(CellFactory(new Phrase("前前次", chFont_small), Lgreen));
                            }
                            for (int i =0; i < 7; i++)
                            {
                                table_inst2.AddCell(CellFactory_line_color_green(new Phrase(" ", chFont_small), Lgreen, 2));
                                table_inst2.AddCell(CellFactory(new Phrase(" ", chFont_small), Lgreen, 1, 3));
                                table_inst2.AddCell(CellFactory(new Phrase("本次", chFont_small), Lgreen));
                                table_inst2.AddCell(CellFactory(new Phrase("前次", chFont_small), Lgreen));
                                table_inst2.AddCell(CellFactory(new Phrase("前前次", chFont_small), Lgreen));
                            }
                        }else if(k>7 && k < 14)
                        {
                            
                            for (int i = 0; i <= 14 - k; i++)
                            {
                                table_inst2.AddCell(CellFactory_line_color_green(new Phrase(" ", chFont_small), Lgreen, 2));
                                table_inst2.AddCell(CellFactory(new Phrase(" ", chFont_small), Lgreen, 1, 3));
                                table_inst2.AddCell(CellFactory(new Phrase("本次", chFont_small), Lgreen));
                                table_inst2.AddCell(CellFactory(new Phrase("前次", chFont_small), Lgreen));
                                table_inst2.AddCell(CellFactory(new Phrase("前前次", chFont_small), Lgreen));
                            }
                        }
                        table_inst.AddCell(CellFactory(new Phrase(" ", chFont_small), Lgreen, 2));
                        //掛號費優免服務 
                        //1
                        table_inst.AddCell(CellFactory(new Phrase("持本報告回本院門診想第一次掛號費優免服務", chFont_small), BaseColor.WHITE, 2));
                        table_inst.AddCell(CellFactory(new Phrase("1、請參照總結與建議欄位中建議回診科別，每科限優免掛號費一次", chFont_small), BaseColor.WHITE, 2));
                        table_inst.AddCell(CellFactory(new Phrase("2、優免期間為體檢日後九十天內", chFont_small), BaseColor.WHITE, 2));
                        //2
                        table_inst2.AddCell(CellFactory(new Phrase("持本報告回本院門診想第一次掛號費優免服務", chFont_small), BaseColor.WHITE, 2));
                        table_inst2.AddCell(CellFactory(new Phrase("1、請參照總結與建議欄位中建議回診科別，每科限優免掛號費一次", chFont_small), BaseColor.WHITE, 2));
                        table_inst2.AddCell(CellFactory(new Phrase("2、優免期間為體檢日後九十天內", chFont_small), BaseColor.WHITE, 2));
                        PdfPTable table_service = new PdfPTable(2);
                        table_service.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        table_service.PaddingTop = 1f;
                        table_service.DefaultCell.FixedHeight = 15f;
                        table_service.AddCell(new Phrase("科別",chFont_small));
                        table_service.AddCell(new Phrase("日期", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        table_service.AddCell(new Phrase(" ", chFont_small));
                        PdfPCell c = new PdfPCell(table_service);
                        c.Colspan = 2;
                        table_inst.AddCell(c);
                        table_inst.AddCell(CellFactory(new Phrase("", chFont_small), BaseColor.BLACK, 2));
                        table_inst2.AddCell(c);
                        table_inst2.AddCell(CellFactory(new Phrase("", chFont_small), BaseColor.BLACK, 2));


                        table1.AddCell(table_inst);

                        //第四欄 理學檢查
                        //補空格


                        for (int i = 0; i < 8 - j; i++)
                        {
                            table_pe.AddCell(CellFactory_line_content("",Dblue,BaseColor.WHITE, 2));
                        }
                        for(int i = 0; i < 8; i++)
                        {
                            table_pe2.AddCell(CellFactory_line_content("", Dblue, BaseColor.WHITE, 2));
                        }
                        table_pe.AddCell(""); table_pe.AddCell("");
                        table_pe2.AddCell(""); table_pe2.AddCell("");
                        //聽力檢查
                        PdfPTable listen = new PdfPTable(new float[] { 11, 10, 10, 10, 10, 10, 10, 10 });
                        listen.DefaultCell.BorderWidth =0;
                        PdfPCell cell_l1 = new PdfPCell();
                        listen.AddCell(CellFactory_line(new Phrase("聽力檢查", chFont_small),8));
                        listen.AddCell(CellFactory_line(new Phrase(" ", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("500Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("1000Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("2000Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("3000Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("4000Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("6000Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("7000Hz", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("右耳dB", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r500, chFont_small),1,1,false));
                        listen.AddCell(CellFactory_line(new Phrase(r1000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r2000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r3000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r4000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r6000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r8000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase("左耳dB", chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r500, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r1000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r2000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r3000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r4000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r6000, chFont_small), 1, 1, false));
                        listen.AddCell(CellFactory_line(new Phrase(r8000, chFont_small), 1, 1, false));
                        
                        PdfPCell listen_cell = new PdfPCell(listen);
                        listen_cell.Colspan = 2;
                        listen_cell.BorderColor = Dblue;
                        table_pe.AddCell(listen_cell);
                        table_pe.AddCell(" ");table_pe.AddCell(" ");
                        table_pe2.AddCell(listen_cell);
                        table_pe2.AddCell(" "); table_pe2.AddCell(" ");
                        //總結與建議

                        string[] gtext_array = gtext.Split('\n');
                        Phrase ggtext = new Phrase();
                        ggtext.Add(new Chunk("總結與建議\n\n", chFont_small));
                        foreach (string g_str in gtext_array)
                        {
                            if (g_str != "")
                            {
                                ggtext.Add(new Chunk("●", chFont_small_1));
                                ggtext.Add(new Chunk(g_str, chFont_small));
                                ggtext.Add(new Chunk("\n", chFont_small));
                            }
                                
                        }
                        PdfPCell suggest = new PdfPCell(ggtext);
                        suggest.Colspan = 2;
                        suggest.BorderColor = Orange;
                        suggest.NoWrap = false;
                        suggest.FixedHeight = 400f;
                        table_pe.AddCell(suggest);

                        PdfPCell suggest2 = new PdfPCell(new Phrase("總結與建議\n\n", chFont_small));
                        suggest2.Colspan = 2;
                        suggest2.BorderColor = Orange;
                        suggest2.NoWrap = false;
                        suggest2.FixedHeight = 400f;
                        table_pe2.AddCell(suggest2);


                        PdfPTable doctor_table = new PdfPTable(new float[] { 12,18,12,18});
                        doctor_table.DefaultCell.BorderWidth = 0f; 
                        doctor_table.DefaultCell.BorderColor = BaseColor.WHITE;
                        if (doctor == "")
                        {
                            doctor_table.AddCell(new Phrase("檢查醫師:", chFont_small));
                            doctor_table.AddCell("");
                            
                        }
                        else
                        {
                            Image doctor_img = Image.GetInstance("http://210.61.246.116/drimg/" + doctor + ".jpg");

                            doctor_img.WidthPercentage = 100;

                            PdfPCell imgCell = new PdfPCell();
                            imgCell.Border = 0;
                            imgCell.BorderColor = BaseColor.WHITE;
                            imgCell.BorderWidth = 0;
                            imgCell.AddElement(doctor_img);

                            

                            doctor_table.AddCell(new Phrase("檢查醫師:",chFont_small));
                            doctor_table.AddCell(imgCell);
                        }

                        if(doctor2 == "")
                        {
                            doctor_table.AddCell(new Phrase("總評醫師:", chFont_small));
                            doctor_table.AddCell("");
                        }
                        else
                        {
                            Image doctor_img2 = Image.GetInstance("http://210.61.246.116/drimg/" + doctor2 + ".jpg");

                            doctor_img2.WidthPercentage = 100;

                            PdfPCell imgCell2 = new PdfPCell();
                            imgCell2.Border = 0;
                            imgCell2.BorderColor = BaseColor.WHITE;
                            imgCell2.BorderWidth = 0;
                            imgCell2.AddElement(doctor_img2);

                            doctor_table.AddCell(new Phrase("總評醫師:", chFont_small));
                            doctor_table.AddCell(imgCell2);
                        }

                        PdfPCell doctor_cell = new PdfPCell(doctor_table);
                        doctor_cell.Border = 0;
                        doctor_cell.Colspan = 2;
                        table_pe.AddCell(doctor_cell);
                        table_pe2.AddCell(doctor_cell);
                        //理學檢查加入欄位
                        table1.AddCell(table_pe);
                        doc1.Add(table1);
                        doc1.NewPage();


                        //頭2
                        if ((last - first) >= 82)
                        {

                            PdfPTable table_2 = new PdfPTable(new float[] { 10, 10, 12, 8 });
                            table_2.DefaultCell.Border = 0;
                            table_2.TotalWidth = 1150f;
                            table_2.LockedWidth = true;
                            table_2.HorizontalAlignment = Element.ALIGN_CENTER;

                            table_2.AddCell(cell1);
                            table_2.AddCell(CellFactory(" ",BaseColor.WHITE, 2));
                            table_2.AddCell(table2);
                            table_2.AddCell(table3);
                            table_2.AddCell(table4);
                            table_2.AddCell(table5);
                            //第二頁內容
                            table_2.AddCell(table_l2);
                            table_2.AddCell(table_r2);
                            table_2.AddCell(table_inst2);
                            table_2.AddCell(table_pe2);

                            doc1.Add(table_2);
                        }
                        



                        doc1.Close();
                    }
                }
            }
        }





        private PdfPCell CellFactory_line_title(Phrase phrase,BaseColor Lcolor, BaseColor Bcolor, int colspan = 1, int rowspan = 1, bool LC = true)
        {
            PdfPCell cell = new PdfPCell(phrase);
            cell.HorizontalAlignment = (LC) ? Element.ALIGN_LEFT : Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 0;
            cell.BorderWidthTop = 1.5f;
            cell.BorderWidthBottom = 1.5f;
            cell.MinimumHeight = 15f;
            cell.BorderColor = Lcolor;//Dblue;
            cell.BackgroundColor = Bcolor;//Lblue

            return cell;
        }
        private PdfPCell CellFactory_line_content(Phrase phrase, BaseColor Lcolor, BaseColor Bcolor, int colspan = 1, int rowspan = 1, bool LC = true)
        {
            PdfPCell cell = new PdfPCell(phrase);
            cell.HorizontalAlignment = (LC) ? Element.ALIGN_LEFT : Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 0;
            cell.BorderWidthBottom = 1f;
            cell.MinimumHeight = 16f;
            cell.BorderColor = Lcolor;// Dblue;
            cell.BackgroundColor = Bcolor;//Lblue
            cell.NoWrap = false;
            return cell;
        }
        private PdfPCell CellFactory_line_content(string str, BaseColor Lcolor, BaseColor Bcolor, int colspan = 1, int rowspan = 1, bool LC = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, chFont_small));
            cell.HorizontalAlignment = (LC) ? Element.ALIGN_LEFT : Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 0;
            cell.BorderWidthBottom = 1f;
            cell.MinimumHeight = 16f;
            cell.BorderColor = Lcolor;// Dblue;
            cell.BackgroundColor = Bcolor;//Lblue
            cell.NoWrap = false;
            return cell;
        }
        private PdfPCell CellFactory_line_qty(string str, string isunusual,BaseColor Bcolor, int colspan = 1, int rowspan = 1, bool LC = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, (isunusual == "Y") ? chFont_red_small : chFont_small));
            cell.HorizontalAlignment = (LC) ? Element.ALIGN_LEFT : Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 0;
            cell.BorderWidthBottom = 1f;
            cell.MinimumHeight = 16f;
            cell.BorderColor = Dblue;
            cell.BackgroundColor = Bcolor;
            cell.NoWrap = false;
            return cell;
        }
        private PdfPCell CellFactory(string str, BaseColor color, int colspan = 1, int rowspan = 1, bool LC = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, chFont));
            cell.HorizontalAlignment = (LC) ? Element.ALIGN_LEFT : Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 1;
            cell.Border = 1;
            cell.MinimumHeight = 15f;
            cell.BorderColor = color;
            return cell;
        }
        
        private PdfPCell CellFactory(Phrase phrase , BaseColor color , int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(phrase);
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 1;
            cell.Border = 1;
            cell.MinimumHeight = 16f;
            cell.BorderColor = color;
            return cell;
        }
        
        
        private PdfPCell CellFactory_line(Phrase phrase, int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(phrase);
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidthLeft = 1f;
            cell.BorderWidthRight = 0f;
            cell.BorderWidthTop = 1f;
            cell.BorderWidthBottom = 0f;
            cell.MinimumHeight = 15f;
            cell.BorderColor = Dblue;
            return cell;
        }
        
        
        
        private PdfPCell CellFactory_line_color_green(Phrase phrase, BaseColor color, int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(phrase);
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 1;
            cell.BorderWidthBottom = 1;
            cell.Border = 1;
            cell.MinimumHeight = 15f;
            cell.BackgroundColor = Dgreen;
            cell.BorderColor = Lgreen;
            return cell;
        }
        private DataTable before_data(string pid, string arraydate, string customerNo)
        {
            DataTable dt1 = new DataTable();

            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select id,arraydate,pid,CustomerNo from history_head where pid = '" + pid + "' and arraydate BETWEEN '2000-01-01' and DATE('" + arraydate + "') and CustomerNo = '" + customerNo + "'  ORDER BY arraydate DESC;", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataAdapter myDA = new MySqlDataAdapter(command);
                    myDA.Fill(dt1);
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料量過大，請重新查詢\n" + ex.Data);

            }


            return dt1;
        }

        private DataTable previous_data(string id)
        {
            DataTable dt1 = new DataTable();

            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select hh.id,hh.arraydate,hb.name_tc,hb.item_iid,hb.qty,hb.isunusual from history_head hh INNER JOIN  (history_body hb,item i) on hh.id = hb.head_id and hb.item_no = i.no  WHERE id = '" + id + "' and (hb.PayKind = '00' or hb.PayKind = '01' or (hb.PayKind = '02' and i.is_pe = '1' and i.category = 'B004M057'));", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataAdapter myDA = new MySqlDataAdapter(command);
                    myDA.Fill(dt1);
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料量過大，請重新查詢\n" + ex.Data);

            }


            return dt1;
        }
        private DataTable is_pe_data(string pid,string arraydate)
        {
            DataTable dt1 = new DataTable();

            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select hh.pid,hh.arraydate,hb.name_tc,hb.qty,hb.item_iid,hb.isunusual" +
                                                            " from history_body hb" +
                                                            " INNER JOIN history_head hh on hb.head_id = hh.id" +
                                                            " where" +
                                                            " (hb.item_iid = 'R000938' OR" +
                                                            " hb.item_iid = 'R000910' OR" +
                                                            " hb.item_iid = 'R000583' OR" +
                                                            " hb.item_iid = 'R000142' OR" +
                                                            " hb.item_iid = 'R000143' OR" +
                                                            " hb.item_iid = 'R000144' OR" +
                                                            " hb.item_iid = 'R000145' OR" +
                                                            " hb.item_iid = 'R000146' OR" +
                                                            " hb.item_iid = 'R000147' OR" +
                                                            " hb.item_iid = 'R000148') and hh.arraydate = DATE('"+arraydate+"') and hh.pid = '"+pid+"'; ", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataAdapter myDA = new MySqlDataAdapter(command);
                    myDA.Fill(dt1);
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("資料量過大，請重新查詢\n" + ex.Data);

            }


            return dt1;
        }
        private void is_DF_data(string pid, string arraydate)
        {
            DataTable dt1 = new DataTable();

            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("select class,hb.name_tc,hb.qty"+
                                                            " from history_body hb" +
                                                            " INNER JOIN(history_head hh, item i) on hh.id = hb.head_id and hb.item_no = i.`no`" +
                                                            " where hh.arraydate = DATE('"+arraydate+"') and hh.pid = '"+pid+ "' and(i.class = 'D既往病史' or i.class = 'F自覺症狀') and (i.category = 'B004M032' or i.category = 'B004M034');", cnn);
                    command.CommandTimeout = 0;
                    MySqlDataAdapter myDA = new MySqlDataAdapter(command);
                    myDA.Fill(dt1);
                    cnn.Close();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("既往病史 自覺症狀 資料量過大，請重新查詢\n" + ex.Data);

            }
            D = "";F = "";
            foreach(DataRow dtr in dt1.Rows)
            {
                if(dtr[0].ToString() == "D既往病史")
                {
                    if(dtr[2].ToString() == "true" || dtr[1].ToString() == "癌症名稱" || dtr[1].ToString() == "骨折部位" || dtr[1].ToString() == "其他慢性病名稱")
                    {
                        if(dtr[1].ToString() == "癌症名稱" || dtr[1].ToString() == "骨折部位" || dtr[1].ToString() == "其他慢性病名稱")
                        {
                            if(dtr[2].ToString() != "")
                                D += " " + dtr[2].ToString();
                        }
                        else
                        {
                            D += " " + dtr[1].ToString();
                        }
                    }
                }
                else if(dtr[0].ToString() == "F自覺症狀")
                {
                    if (dtr[2].ToString() == "true" || dtr[1].ToString() == "其他症狀____")
                    {
                        if (dtr[1].ToString() == "其他症狀____")
                        {
                            if(dtr[2].ToString() != "")
                                F += " " + dtr[2].ToString();
                        }
                        else
                        {
                            F += " " + dtr[1].ToString();
                        }
                    }
                }
            }

        }

    }
}

