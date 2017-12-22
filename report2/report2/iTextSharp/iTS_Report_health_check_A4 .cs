using iTextSharp.text;
using iTextSharp.text.pdf;
using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace report2
{
    class iTS_Report_health_check_A4
    {

        BaseFont bfChinese = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        BaseFont bfChinese2 = BaseFont.CreateFont(@"C:\Windows\Fonts\msjh.ttc,0", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

        Font chFont;
        Font chFont_UnderLine;
        Font chFont_blue;
        Font chFont_red;
        Font chFont_header;
        Font chFont_small1, chFont_small2, chFont_small3;

        string cName, dept_name, name, pid, sex, birtyday, emp_no, sid, arraydate, category, category_tc, name_tc, qty, isunusual,id, m_normal_d, m_normal_u, unit,is_pe,is_inst,gtext,PayKind,customerNo;
        string group_id = "",group_category = "",doctor = "",qty2 = "",qty3 = "", isunusual2 = "", isunusual3 = "";
        DataTable dt_before1 = null, dt_before2 = null, dt_is_pe = null;
        string D = "", F = "";
        public void ExportPdf(DataTable dt, string dirPath,string dirName)
        {
            

            int first = 0, last = 0;
            string str = "";
            group_id = dt.Rows[0]["id"].ToString();
            for(int i = 0; i < dt.Rows.Count; i++)
            {
                
                id = dt.Rows[i]["id"].ToString();
                if(group_id != id  || i == dt.Rows.Count-1)
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


                    outputDocument(dt,dt_before1,dt_before2, dirPath, first,last,dirName);

                    //str = str +"\n"+ group_id + " " + id + " " + first + " " + last  ;
                }
                
            }

            //MessageBox.Show(str);
        }

        private void outputDocument(DataTable dt,DataTable dt_1,DataTable dt_2,string dirPath,int first, int last,string dirName)
        {
            //字型設定
            chFont = new Font(bfChinese, 12);
            chFont_UnderLine = new Font(bfChinese, 12, Font.UNDERLINE);
            chFont_blue = new Font(bfChinese, 10, Font.NORMAL, new BaseColor(0, 0, 240));
            chFont_red = new Font(bfChinese, 8, Font.NORMAL, BaseColor.RED);
            chFont_header = new Font(bfChinese, 18, Font.BOLD, BaseColor.BLACK);
            chFont_small1 = new Font(bfChinese2, 10, Font.NORMAL, BaseColor.BLACK);
            chFont_small2 = new Font(bfChinese2, 8, Font.NORMAL, BaseColor.BLACK);
            chFont_small3 = new Font(bfChinese2, 10, Font.BOLD, BaseColor.WHITE);
            //-------------------------------------
            //dt.Rows[0][1]
            //名稱判斷
            string[] str = new string[5];
            str[0] = (dirName.Split(',').GetValue(0).ToString() == "true") ? dt.Rows[first]["sid"].ToString()+" " : "";
            str[1] = (dirName.Split(',').GetValue(1).ToString() == "true") ? dt.Rows[first]["name"].ToString()+" " : "";
            str[2] = (dirName.Split(',').GetValue(2).ToString() == "true") ? dt.Rows[first]["dept_name"].ToString()+" " : "";
            str[3] = (dirName.Split(',').GetValue(3).ToString() == "true") ? dt.Rows[first]["emp_no"].ToString()+" " : "";
            str[4] = (dirName.Split(',').GetValue(4).ToString() == "true") ? dt.Rows[first]["pid"].ToString()+" " : "";

             

            string outputFile = Path.Combine(dirPath, str[0] + str[1] + str[2] + str[3] + str[4] + ".pdf");
            //Create a standard .Net FileStream for the file, setting various flags
            using (FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                
                //Create a new PDF document setting the size to A4
                using (Document doc1 = new Document(PageSize.A4, 5, 5, 3, 3))
                {
                    //Bind the PDF document to the FileStream using an iTextSharp PdfWriter
                    using (PdfWriter w = PdfWriter.GetInstance(doc1, fs))
                    {


                        //Open the document for writing
                        doc1.Open();

                        //body================
                        PdfPTable body = new PdfPTable(2);
                        body.TotalWidth = 570f;
                        body.LockedWidth = true;
                        body.DefaultCell.BorderWidth = 0f;
                        body.PaddingTop = 0f;

                        PdfPTable head = new PdfPTable(3);
                        PdfPTable body_l = new PdfPTable(new float[] {125, 25, 25, 25 });
                        body_l.DefaultCell.Border = 0;

                        PdfPTable body_r = new PdfPTable(new float[] {125, 25, 25, 25 });
                        PdfPTable body_l2 = new PdfPTable(new float[] { 125, 25, 25, 25 });

                        
                        int l = 1;
                        qty2 = ""; isunusual2 = "";
                        qty3 = ""; isunusual3 = "";
                        for (int i = first; i < last; i++)
                        {
                            if(dt_before1 != null)
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
                            birtyday =Convert.ToDateTime(dt.Rows[i]["birtyday"].ToString()).ToString("yyyy/MM/dd");
                            emp_no = dt.Rows[i]["emp_no"].ToString();
                            sid = dt.Rows[i]["sid"].ToString();
                            arraydate = dt.Rows[i]["arraydate"].ToString();
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
                            PayKind = dt.Rows[i]["PayKind"].ToString();
                            if (i == first)//員工資訊
                            {
                                
                                head.TotalWidth = 570f;
                                head.LockedWidth = true;

                                PdfPCell header = new PdfPCell();
                                header.HorizontalAlignment = Element.ALIGN_CENTER;
                                header.AddElement(new Chunk("健康檢查紀錄表", chFont_header));
                                header.Colspan = 3;
                                header.Border = 0;

                                head.AddCell(header);

                                //既往病史 自覺症狀
                                is_DF_data(pid, arraydate);
                                head.AddCell(CellFactory("公司名稱:"+cName));
                                head.AddCell(CellFactory("部門名稱:"+dept_name));
                                head.AddCell(CellFactory("頁數:"+"1/2"));
                                head.AddCell(CellFactory("姓名:"+name));
                                head.AddCell(CellFactory("性別:"+((sex=="M")?"男":"女")));
                                head.AddCell(CellFactory("身分證字號:"+pid));
                                head.AddCell(CellFactory("體檢日期:"+arraydate));
                                head.AddCell(CellFactory("受檢序號:"+sid));
                                head.AddCell(CellFactory("出生年月日:" + birtyday));
                                string str1 = (dt_before1 != null) ?Convert.ToDateTime(dt_before1.Rows[0][1]).ToString("yyyy/MM/dd") : "";
                                head.AddCell(CellFactory("前次體檢日期:"+str1));
                                head.AddCell(CellFactory("工號:" + emp_no));
                                head.AddCell(CellFactory("既往病史:"+D));
                                string str2 = (dt_before2 != null) ? Convert.ToDateTime(dt_before2.Rows[0][1]).ToString("yyyy/MM/dd") : "";
                                head.AddCell(CellFactory("前前次體檢日期:"+str2));
                                head.AddCell(CellFactory(" "));
                                head.AddCell(CellFactory("自覺症狀:"+F));


                                //寫死 理學檢查
                                if(PayKind != "01")
                                {
                                    l++;
                                    body_l.AddCell(CellFactory_bcolor("理學檢查"));
                                    body_l.AddCell(CellFactory_bcolor("本次"));
                                    body_l.AddCell(CellFactory_bcolor("前次"));
                                    body_l.AddCell(CellFactory_bcolor("前前次"));
                                }
                                doctor = "";
                                dt_is_pe = is_pe_data(pid, arraydate);
                                foreach (DataRow ispe in dt_is_pe.Rows)
                                {
                                    if (ispe[4].ToString() == "R000910" || ispe[4].ToString() == "R000938")
                                    {
                                        if (ispe[4].ToString() == "R000910")
                                            doctor = ispe[3].ToString();
                                    }
                                    else
                                    {
                                        if (PayKind != "01")
                                        {
                                            l++;
                                            body_l.AddCell(CellFactory_small(ispe[2].ToString()));
                                            body_l.AddCell(CellFactory_small(ispe[3].ToString(), ispe[5].ToString()));
                                            body_l.AddCell(CellFactory_small(" "));
                                            body_l.AddCell(CellFactory_small(" "));

                                        }

                                    }
                                }



                            }

                            if ((is_pe == "True" && (category == "B004M036" || category == "B004M056")) || name_tc == "總評醫師")
                            {
                                
                            }
                            //
                            //左 超過44列換第二欄
                            else if (body_l.Rows.Count < 44)
                            {
                                
                                //category群組
                                if (group_category != category)
                                {

                                    group_category = category;

                                    
                                    //儀器判斷 群組名稱
                                    if (is_inst != "True")
                                    {
                                        l++;
                                        if (PayKind == "00")//加選項目判斷 00 非加選 01 加選
                                        {
                                            body_l.AddCell(CellFactory_bcolor(category_tc));
                                            body_l.AddCell(CellFactory_bcolor("本次"));
                                            body_l.AddCell(CellFactory_bcolor("前次"));
                                            body_l.AddCell(CellFactory_bcolor("前前次"));
                                        }
                                        else
                                        {
                                            PdfPCell payKindCell = new PdfPCell(new Phrase("加選項目-" + category_tc, chFont_small3));
                                            payKindCell.FixedHeight = 15f;
                                            payKindCell.BackgroundColor = new BaseColor(165, 42, 42);
                                            body_l.AddCell(payKindCell);
                                            body_l.AddCell(CellFactory_wcolor("本次"));
                                            body_l.AddCell(CellFactory_wcolor("前次"));
                                            body_l.AddCell(CellFactory_wcolor("前前次"));
                                        }
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        body_l.AddCell(CellFactory_small(name_tc));
                                        body_l.AddCell(CellFactory_small(qty, isunusual));
                                        body_l.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_l.AddCell(CellFactory_small(qty3, isunusual3));

                                    }
                                    else
                                    {
                                        l += 3;
                                        PdfPCell ctitle = new PdfPCell(new Phrase(category_tc, chFont_small3));
                                        ctitle.Colspan = 4;
                                        ctitle.FixedHeight = 15f;
                                        ctitle.BackgroundColor = new BaseColor(165, 42, 42);
                                        body_l.AddCell(ctitle);


                                        PdfPTable body_inst = new PdfPTable(new float[] { 2, 2, 6 });
                                        body_inst.WidthPercentage = 100f;
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        PdfPCell cell = new PdfPCell(new Phrase(name_tc, chFont_small2));
                                        cell.Rowspan = 3;
                                        cell.FixedHeight = 15f;
                                        body_inst.AddCell(cell);
                                        // 本次 前次 前前次
                                        body_inst.AddCell(CellFactory_small("本次"));
                                        body_inst.AddCell(CellFactory_small(qty, isunusual));
                                        body_inst.AddCell(CellFactory_small("前次"));
                                        body_inst.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_inst.AddCell(CellFactory_small("前前次"));
                                        body_inst.AddCell(CellFactory_small(qty3, isunusual3));

                                        PdfPCell bigcell = new PdfPCell(body_inst);
                                        bigcell.Colspan = 4;

                                        body_l.AddCell(bigcell);
                                    }

                                }
                                else
                                {
                                    //儀器判斷 非群組
                                    if (is_inst != "True")
                                    {
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        body_l.AddCell(CellFactory_small(name_tc));
                                        body_l.AddCell(CellFactory_small(qty, isunusual));
                                        body_l.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_l.AddCell(CellFactory_small(qty3, isunusual3));
                                    }
                                    else
                                    {
                                        l += 2;
                                        PdfPTable body_inst = new PdfPTable(new float[] { 2, 2, 6 });
                                        body_inst.WidthPercentage = 100f;
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        PdfPCell cell = new PdfPCell(new Phrase(name_tc, chFont_small2));
                                        cell.Rowspan = 3;
                                        cell.FixedHeight = 15f;
                                        body_inst.AddCell(cell);
                                        // 本次 前次 前前次
                                        body_inst.AddCell(CellFactory_small("本次"));
                                        body_inst.AddCell(CellFactory_small(qty, isunusual));
                                        body_inst.AddCell(CellFactory_small("前次"));
                                        body_inst.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_inst.AddCell(CellFactory_small("前前次"));
                                        body_inst.AddCell(CellFactory_small(qty3, isunusual3));

                                        PdfPCell bigcell = new PdfPCell(body_inst);
                                        bigcell.Colspan = 4;
                                        body_l.AddCell(bigcell);
                                    }

                                }

                            }//右
                            else if (body_r.Rows.Count <44)
                            {
                                
                                //category群組判斷
                                if (group_category != category)
                                {
                                    group_category = category;
                                    //儀器判斷 群組名稱
                                    if (is_inst != "True")
                                    {
                                        l++;
                                        if (PayKind == "00")//加選項目
                                        {
                                            body_r.AddCell(CellFactory_bcolor(category_tc));
                                            body_r.AddCell(CellFactory_bcolor("本次"));
                                            body_r.AddCell(CellFactory_bcolor("前次"));
                                            body_r.AddCell(CellFactory_bcolor("前前次"));
                                        }
                                        else
                                        {
                                            PdfPCell payKindCell = new PdfPCell(new Phrase("加選項目-" + category_tc, chFont_small3));
                                            payKindCell.FixedHeight = 15f;
                                            payKindCell.BackgroundColor = new BaseColor(165, 42, 42);
                                            body_r.AddCell(payKindCell);
                                            body_r.AddCell(CellFactory_wcolor("本次"));
                                            body_r.AddCell(CellFactory_wcolor("前次"));
                                            body_r.AddCell(CellFactory_wcolor("前前次"));
                                        }
                                        //數值基準值
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        body_r.AddCell(CellFactory_small(name_tc));
                                        body_r.AddCell(CellFactory_small(qty, isunusual));
                                        body_r.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_r.AddCell(CellFactory_small(qty3, isunusual3));
                                    }
                                    else
                                    {
                                        l += 3;
                                        PdfPCell ctitle = new PdfPCell(new Phrase(category_tc, chFont_small3));
                                        ctitle.Colspan = 4;
                                        ctitle.FixedHeight = 15f;
                                        ctitle.BackgroundColor = new BaseColor(165, 42, 42);
                                        body_r.AddCell(ctitle);



                                        PdfPTable body_inst = new PdfPTable(new float[] { 2, 2, 6 });
                                        body_inst.WidthPercentage = 100f;
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        PdfPCell cell = new PdfPCell(new Phrase(name_tc, chFont_small2));
                                        cell.Rowspan = 3;
                                        cell.FixedHeight = 15f;
                                        body_inst.AddCell(cell);
                                        // 本次 前次 前前次
                                        body_inst.AddCell(CellFactory_small("本次"));
                                        body_inst.AddCell(CellFactory_small(qty, isunusual));
                                        body_inst.AddCell(CellFactory_small("前次"));
                                        body_inst.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_inst.AddCell(CellFactory_small("前前次"));
                                        body_inst.AddCell(CellFactory_small(qty3, isunusual3));

                                        PdfPCell bigcell = new PdfPCell(body_inst);
                                        bigcell.Colspan = 4;
                                        body_r.AddCell(bigcell);
                                    }

                                }
                                else
                                {
                                    //儀器判斷 非群組名稱
                                    if (is_inst != "True")
                                    {
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        body_r.AddCell(CellFactory_small(name_tc));
                                        body_r.AddCell(CellFactory_small(qty, isunusual));
                                        body_r.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_r.AddCell(CellFactory_small(qty3, isunusual3));
                                    }
                                    else
                                    {
                                        l += 2;
                                        PdfPTable body_inst = new PdfPTable(new float[] { 2, 2, 6 });
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        PdfPCell cell = new PdfPCell(new Phrase(name_tc, chFont_small2));
                                        cell.Rowspan = 3;
                                        cell.FixedHeight = 15f;
                                        body_inst.AddCell(cell);
                                        // 本次 前次 前前次
                                        body_inst.AddCell(CellFactory_small("本次"));
                                        body_inst.AddCell(CellFactory_small(qty, isunusual));
                                        body_inst.AddCell(CellFactory_small("前次"));
                                        body_inst.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_inst.AddCell(CellFactory_small("前前次"));
                                        body_inst.AddCell(CellFactory_small(qty3, isunusual3));


                                        PdfPCell bigcell = new PdfPCell(body_inst);
                                        bigcell.Colspan = 4;
                                        body_r.AddCell(bigcell);

                                    }

                                }

                            }
                            else
                            {
                                if (group_category != category)
                                {

                                    group_category = category;

                                    //儀器判斷 群組名稱
                                    if (is_inst != "True")
                                    {
                                        l++;
                                        if (PayKind == "00")//加選項目
                                        {
                                            body_l2.AddCell(CellFactory_bcolor(category_tc));
                                            body_l2.AddCell(CellFactory_bcolor("本次"));
                                            body_l2.AddCell(CellFactory_bcolor("前次"));
                                            body_l2.AddCell(CellFactory_bcolor("前前次"));
                                        }
                                        else
                                        {
                                            PdfPCell payKindCell = new PdfPCell(new Phrase("加選項目-" + category_tc, chFont_small3));
                                            payKindCell.FixedHeight = 15f;
                                            payKindCell.BackgroundColor = new BaseColor(165, 42, 42);
                                            body_l2.AddCell(payKindCell);
                                            body_l2.AddCell(CellFactory_wcolor("本次"));
                                            body_l2.AddCell(CellFactory_wcolor("前次"));
                                            body_l2.AddCell(CellFactory_wcolor("前前次"));
                                        }

                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        body_l2.AddCell(CellFactory_small(name_tc));
                                        body_l2.AddCell(CellFactory_small(qty, isunusual));
                                        body_l2.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_l2.AddCell(CellFactory_small(qty3, isunusual3));
                                    }
                                    else
                                    {
                                        l += 3;
                                        PdfPCell ctitle = new PdfPCell(new Phrase(category_tc, chFont_small3));
                                        ctitle.Colspan = 4;
                                        ctitle.FixedHeight = 15f;
                                        ctitle.BackgroundColor = new BaseColor(165, 42, 42);
                                        body_l2.AddCell(ctitle);


                                        PdfPTable body_inst = new PdfPTable(new float[] { 2, 2, 6 });
                                        body_inst.WidthPercentage = 100f;
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        PdfPCell cell = new PdfPCell(new Phrase(name_tc, chFont_small2));
                                        cell.Rowspan = 3;
                                        body_inst.AddCell(cell);
                                        // 本次 前次 前前次
                                        body_inst.AddCell(CellFactory_small("本次"));
                                        body_inst.AddCell(CellFactory_small(qty, isunusual));
                                        body_inst.AddCell(CellFactory_small("前次"));
                                        body_inst.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_inst.AddCell(CellFactory_small("前前次"));
                                        body_inst.AddCell(CellFactory_small(qty3, isunusual3));

                                        PdfPCell bigcell = new PdfPCell(body_inst);
                                        bigcell.Colspan = 4;

                                        body_l2.AddCell(bigcell);
                                    }

                                }
                                else
                                {
                                    //儀器判斷 非群組名稱
                                    if (is_inst != "True")
                                    {
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        body_l2.AddCell(CellFactory_small(name_tc));
                                        body_l2.AddCell(CellFactory_small(qty, isunusual));
                                        body_l2.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_l2.AddCell(CellFactory_small(qty3, isunusual3));
                                    }
                                    else
                                    {
                                        l += 2;
                                        PdfPTable body_inst = new PdfPTable(new float[] { 2, 2, 6 });
                                        body_inst.WidthPercentage = 100f;
                                        if (m_normal_d != "")
                                            name_tc = name_tc + "<" + m_normal_d + "-" + m_normal_u + unit + ">";
                                        PdfPCell cell = new PdfPCell(new Phrase(name_tc, chFont_small2));
                                        cell.Rowspan = 3;
                                        body_inst.AddCell(cell);
                                        // 本次 前次 前前次
                                        body_inst.AddCell(CellFactory_small("本次"));
                                        body_inst.AddCell(CellFactory_small(qty, isunusual));
                                        body_inst.AddCell(CellFactory_small("前次"));
                                        body_inst.AddCell(CellFactory_small(qty2, isunusual2));
                                        body_inst.AddCell(CellFactory_small("前前次"));
                                        body_inst.AddCell(CellFactory_small(qty3, isunusual3));

                                        PdfPCell bigcell = new PdfPCell(body_inst);
                                        bigcell.Colspan = 4;
                                        body_l2.AddCell(bigcell);
                                    }

                                }
                            }






                            l++;
                            
                        }
                        body_l.AddCell(CellFactory(" ", 4));
                        body_r.AddCell(CellFactory(" ", 4));
                        //head--------------
                        PdfPCell headCell = new PdfPCell(head);
                        headCell.Colspan = 2;
                        headCell.Border = 0;
                        body.AddCell(headCell);
                        //head--------------

                        body.AddCell(body_l);
                        body.AddCell(body_r);
                        //紅字註釋--------------
                        PdfPCell redCell = new PdfPCell(new Phrase("結果值內空值表示未檢 紅色表示應多注意 參考值內正常範圍",chFont_red));
                        redCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        redCell.Border = 0;
                        redCell.Colspan = 2;
                        body.AddCell(redCell);
                        //--------------
                        if (l > 90)
                        {
                            body.AddCell(body_l2);
                            body.AddCell(CellFactory(" "));
                        }

                        //第二頁===============
                        //body
                        PdfPTable body2 = new PdfPTable(2);
                        body2.TotalWidth = 570f;
                        body2.LockedWidth = true;
                        body2.DefaultCell.BorderWidth = 0f;
                        body2.PaddingTop = 0f;
                        //header2
                        PdfPCell header2 = new PdfPCell();
                        header2.HorizontalAlignment = Element.ALIGN_CENTER;
                        header2.AddElement(new Chunk("健康檢查紀錄表\n", chFont_header));
                        Chunk c1 = new Chunk("頁數：2/2    ", chFont);
                        Chunk c2 = new Chunk("受檢序號：" + sid + "        ", chFont);
                        Chunk c3 = new Chunk("姓 名：" + name + "        ", chFont);
                        Chunk c4 = new Chunk("出生年月日：" + birtyday, chFont);
                        Phrase p1 = new Phrase();
                        p1.Add(c1);
                        p1.Add(c2);
                        p1.Add(c3);
                        p1.Add(c4);
                        header2.AddElement(p1);
                        header2.Colspan = 2;
                        header2.Border = 0;
                        body2.AddCell(header2);
                        //content
                        PdfPTable body_page2 = new PdfPTable(new float[] {2,5,2,4});
                        body_page2.DefaultCell.Border = 1;
                        body_page2.AddCell(new PdfPCell(new Phrase("檢查醫療機構名稱\n電話及地址",chFont)));
                        body_page2.AddCell(new PdfPCell(new Phrase("臺北榮民總院桃園分院\n(03)3384889分機1731\n桃園市成功路三段100號", chFont)));
                        body_page2.AddCell(new PdfPCell(new Phrase("總評醫師姓名\n及證書字號", chFont)));
                        if(doctor == "")
                        {
                            body_page2.AddCell(new PdfPCell(new Phrase(" ", chFont)));
                        }
                        else
                        {
                            Image doctor_img = Image.GetInstance("http://210.61.246.116/drimg/" + doctor + ".jpg");
                            doctor_img.WidthPercentage = 40;
                            PdfPCell imgCell = new PdfPCell();
                            imgCell.Border = 0;
                            imgCell.AddElement(doctor_img);
                            body_page2.AddCell(imgCell);
                        }
                        

                        PdfPCell body_page2_cell = new PdfPCell(body_page2);
                        body_page2_cell.Colspan = 2;
                        body2.AddCell(body_page2_cell);


                        //gtext醫師評語
                        gtext = dt.Rows[0]["gtext"].ToString();
                        string[] gtext_array = gtext.Split('\n');
                        gtext = "";
                        foreach(string g_str in gtext_array)
                        {
                            if(g_str != "")
                                gtext += "●" + g_str + "\n";
                        }

                        PdfPCell gtext_cell = new PdfPCell(new Phrase("醫師評語\n\n"+gtext,chFont));
                        gtext_cell.Border = 0;
                        gtext_cell.Colspan = 2;
                        body2.AddCell(gtext_cell);

                        //
                        doc1.Add(body);
                        doc1.NewPage();
                        doc1.Add(body2);
                        //--------------


                        

                        doc1.Close();
                    }
                }
            }
        }

        private PdfPCell CellFactory(string str, int colspan = 1, int rowspan = 1,bool LR = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, chFont));
            cell.HorizontalAlignment =(LR)? Element.ALIGN_LEFT:Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.Border = 0;
            cell.BorderWidth = 1f;
            cell.Padding = 2f;
            
            return cell;
        }
        private PdfPCell CellFactory_bcolor(string str, int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, chFont_small1));
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BackgroundColor = BaseColor.ORANGE;
            cell.BorderWidth = 0.5f;
            cell.BorderWidthLeft = 0.5f;
            cell.BorderWidthRight = 0.5f;
            cell.BorderWidthBottom = 0.5f;
            cell.FixedHeight = 15f;
            return cell;
        }
        private PdfPCell CellFactory_wcolor(string str, int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, chFont_small1));
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 0.5f;
            cell.BorderWidthLeft = 0.5f;
            cell.BorderWidthRight = 0.5f;
            cell.BorderWidthBottom = 0.5f;
            cell.FixedHeight = 15f;
            return cell;
        }
        private PdfPCell CellFactory_small(string str,string isunusual ="", int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str,(isunusual == "Y")?chFont_red: chFont_small2));
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.BorderWidth = 0.5f;
            cell.BorderWidthLeft = 0.5f;
            cell.BorderWidthRight = 0.5f;
            cell.BorderWidthBottom = 0.5f;
            cell.FixedHeight = 15f;
            return cell;
        }
        private PdfPCell CellFactory(Phrase phrase, int colspan = 1, int rowspan = 1, bool LR = true)
        {
            PdfPCell cell = new PdfPCell(new Phrase(phrase));
            cell.HorizontalAlignment = (LR) ? Element.ALIGN_LEFT : Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_CENTER;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.Border = 0;
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
                    MySqlCommand command = new MySqlCommand("select hh.id,hh.arraydate,hb.name_tc,hb.item_iid,hb.qty,hb.isunusual from history_head hh INNER JOIN history_body hb on hh.id = hb.head_id WHERE id = '" + id+"' and (hb.PayKind = '00' or hb.PayKind = '01');", cnn);
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
        private DataTable is_pe_data(string pid, string arraydate)
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
                                                            " hb.item_iid = 'R000148') and hh.arraydate = DATE('" + arraydate + "') and hh.pid = '" + pid + "'; ", cnn);
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
                    MySqlCommand command = new MySqlCommand("select class,hb.name_tc,hb.qty" +
                                                            " from history_body hb" +
                                                            " INNER JOIN(history_head hh, item i) on hh.id = hb.head_id and hb.item_no = i.`no`" +
                                                            " where hh.arraydate = DATE('" + arraydate + "') and hh.pid = '" + pid + "' and(i.class = 'D既往病史' or i.class = 'F自覺症狀') and (i.category = 'B004M032' or i.category = 'B004M034');", cnn);
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
            D = ""; F = "";
            foreach (DataRow dtr in dt1.Rows)
            {
                if (dtr[0].ToString() == "D既往病史")
                {
                    if (dtr[2].ToString() == "true" || dtr[1].ToString() == "癌症名稱" || dtr[1].ToString() == "骨折部位" || dtr[1].ToString() == "其他慢性病名稱")
                    {
                        if (dtr[1].ToString() == "癌症名稱" || dtr[1].ToString() == "骨折部位" || dtr[1].ToString() == "其他慢性病名稱")
                        {
                            if (dtr[2].ToString() != "")
                                D += " " + dtr[2].ToString();
                        }
                        else
                        {
                            D += " " + dtr[1].ToString();
                        }
                    }
                }
                else if (dtr[0].ToString() == "F自覺症狀")
                {
                    if (dtr[2].ToString() == "true" || dtr[1].ToString() == "其他症狀____")
                    {
                        if (dtr[1].ToString() == "其他症狀____")
                        {
                            if (dtr[2].ToString() != "")
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
