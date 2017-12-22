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
    class iTS_Report_FeedStaff_health_check_batch
    {

        BaseFont bfChinese = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        Font chFont;
        Font chFont_UnderLine;
        Font chFont_blue;
        Font chFont_red;
        Font chFont_header;
        Font chFont_small;
        DataTable dt_test_item = null;
        int count = 0;
        public void ExportPdf(DataTable dt, string dirPath,string dirName)
        {
            string[] str = new string[5];
            //字型設定
            chFont = new Font(bfChinese, 12);
            chFont_UnderLine = new Font(bfChinese, 12, Font.UNDERLINE);
            chFont_blue = new Font(bfChinese, 10, Font.NORMAL, new BaseColor(0, 0, 240));
            chFont_red = new Font(bfChinese, 12, Font.NORMAL, BaseColor.RED);
            chFont_header = new Font(bfChinese, 18, Font.BOLD, BaseColor.BLACK);
            chFont_small = new Font(bfChinese, 8, Font.NORMAL, BaseColor.BLACK);

            string outputFile = Path.Combine(dirPath, "batch_file.pdf");
            //Create a standard .Net FileStream for the file, setting various flags
            using (FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                //Create a new PDF document setting the size to A4
                using (Document doc1 = new Document(PageSize.A4, 10, 10, 10, 10))
                {
                    //Bind the PDF document to the FileStream using an iTextSharp PdfWriter
                    using (PdfWriter w = PdfWriter.GetInstance(doc1, fs))
                    {
                        doc1.Open();

                        //dt.Rows[0][1]
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {



                            //data
                            string name = dt.Rows[i]["name"].ToString();
                            string sex = dt.Rows[i]["sex"].ToString();
                            string birth = Convert.ToDateTime(dt.Rows[i]["birtyday"]).ToString("yyyy/MM/dd");
                            string sid = dt.Rows[i]["sid"].ToString();
                            string emp_no = dt.Rows[i]["emp_no"].ToString();
                            string arraydate = dt.Rows[i]["arraydate"].ToString();
                            string cName = dt.Rows[i]["cName"].ToString();
                            string pid = dt.Rows[i]["pid"].ToString();
                            string R000001 = dt.Rows[i]["R000001"].ToString();
                            string R000114 = dt.Rows[i]["R000114"].ToString();
                            string R000084 = dt.Rows[i]["R000084"].ToString();
                            string R000092 = dt.Rows[i]["R000092"].ToString();
                            string R000085 = dt.Rows[i]["R000085"].ToString();
                            string R000109 = dt.Rows[i]["R000109"].ToString();
                            string R000110 = dt.Rows[i]["R000110"].ToString();
                            string R000188 = dt.Rows[i]["R000188"].ToString();
                            string R000188T01 = dt.Rows[i]["R000188T01"].ToString();
                            string A2110101 = dt.Rows[i]["A2110101"].ToString();
                            string A2110102 = dt.Rows[i]["A2110102"].ToString();
                            string R000909 = dt.Rows[i]["R000909"].ToString();
                            string R000919 = dt.Rows[i]["R000919"].ToString();
                            string R000920 = dt.Rows[i]["R000920"].ToString();
                            string R000921 = dt.Rows[i]["R000921"].ToString();

                            string suggest = dt.Rows[i]["suggest"].ToString();
                            string R000910 = dt.Rows[i]["R000910"].ToString();
                            string R000036 = dt.Rows[i]["R000036"].ToString();
                            string R000887 = dt.Rows[i]["R000887"].ToString();
                            string A902 = dt.Rows[i]["A902"].ToString();
                            //Open the document for writing


                            //header
                            Paragraph header = new Paragraph(10f, "從事供膳作業員工健康檢查紀錄表", chFont_header);
                            header.Alignment = Element.ALIGN_CENTER;
                            doc1.Add(header);
                            //customer checkid emp_no
                            Chunk customer = new Chunk("  "+cName + "                                     ", chFont);
                            Chunk checkId = new Chunk("受檢序號:" + sid + "   ", chFont);
                            Chunk Emp_no = new Chunk("工號:" + emp_no, chFont);
                            Phrase p1 = new Phrase();
                            p1.Add(customer);
                            p1.Add(checkId);
                            p1.Add(Emp_no);
                            doc1.Add(p1);
                            //表格
                            PdfPTable table = new PdfPTable(new float[] { 1, 2, 2, 1, 2, 4, 2, 2 });
                            table.HorizontalAlignment = Element.ALIGN_CENTER;
                            table.TotalWidth = 550f;
                            table.LockedWidth = true;
                            table.DefaultCell.FixedHeight = 20f;
                            //1
                            table.AddCell(CellFactory(new Phrase("姓名", chFont))); table.AddCell(CellFactory(name));
                            sex = (sex == "M") ? "男" : "女";
                            table.AddCell(CellFactory(new Phrase("性別", chFont))); table.AddCell(CellFactory(sex));
                            table.AddCell(CellFactory(new Phrase("出生年月日", chFont))); table.AddCell(CellFactory(birth));
                            table.AddCell(CellFactory(new Phrase("受雇年月日", chFont))); table.AddCell(CellFactory(R000001));
                            //2
                            PdfPCell id = CellFactory("身  分  證  字  號", 4, 1);
                            table.AddCell(id);
                            PdfPCell idContent = CellFactory(pid, 2);
                            table.AddCell(idContent);
                            PdfPCell Arraydate = CellFactory("檢查年月日");
                            table.AddCell(Arraydate);
                            PdfPCell arraydateContent = CellFactory(arraydate, 1, 1);
                            table.AddCell(arraydateContent);
                            //3
                            PdfPCell checkTime = CellFactory("健康檢查時期", 4);
                            table.AddCell(checkTime);
                            PdfPCell checkTimeContent = CellFactory((R000036 == "true") ? "新進員工（受僱時）" : "定期檢查", 4);
                            table.AddCell(checkTimeContent);
                            //4
                            PdfPCell rowspan = CellFactory("特\n\n殊\n\n健\n\n康\n\n檢\n\n查", 1, 11);
                            rowspan.VerticalAlignment = Element.ALIGN_MIDDLE;
                            rowspan.HorizontalAlignment = Element.ALIGN_CENTER;
                            table.AddCell(rowspan);
                            //5
                            PdfPCell symptom = CellFactory("自 覺\n症 狀", 1);
                            symptom.FixedHeight = 40f;


                            table.AddCell(symptom);
                            PdfPCell sym = CellFactory(" ", 6);
                            Phrase symPhrase = new Phrase();
                            symPhrase.Add(new Chunk((R000114 == "true") ? "■1.食慾不振" : "□1.食慾不振", chFont));
                            symPhrase.Add(new Chunk((R000084 == "true") ? "■2.咳嗽" : "□2.咳嗽", chFont));
                            symPhrase.Add(new Chunk((R000092 == "true") ? "■3.倦怠" : "□3.倦怠", chFont));
                            symPhrase.Add(new Chunk((R000085 == "true") ? "■4.咳痰\n" : "□4.咳痰\n", chFont));
                            symPhrase.Add(new Chunk((R000109 == "true") ? "■5.焦躁感" : "□5.焦躁感", chFont));
                            symPhrase.Add(new Chunk((R000110 == "true") ? "■6.失眠" : "□6.失眠", chFont));
                            symPhrase.Add(new Chunk((R000188 == "true") ? "■7.其他" : "□7.其他", chFont));
                            symPhrase.Add(new Chunk((R000114 != "true"
                                && R000084 != "true"
                                && R000092 != "true"
                                && R000085 != "true"
                                && R000109 != "true"
                                && R000110 != "true"
                                && R000188 != "true") ? "■8.以上皆無" : "□8.以上皆無", chFont));

                            symPhrase.Add(new Chunk(R000188T01, chFont_UnderLine));

                            sym.AddElement(symPhrase);
                            table.AddCell(sym);
                            //6 General Reasoning
                            PdfPCell gr = CellFactory("一般理學", 1);
                            table.AddCell(gr);
                            PdfPCell height = CellFactory("身高", 3);
                            table.AddCell(height);
                            PdfPCell h = CellFactory(A2110101.Split(':').GetValue(0).ToString(), A2110101.Split(':').GetValue(1).ToString(), 1);
                            table.AddCell(h);
                            PdfPCell weight = CellFactory("體重", 1);
                            table.AddCell(weight);
                            PdfPCell w1 = CellFactory(A2110102.Split(':').GetValue(0).ToString(), A2110102.Split(':').GetValue(1).ToString(), 2);
                            table.AddCell(w1);
                            //doctor rowspan
                            PdfPCell doctor = CellFactory("\n醫\n師\n問\n答", 1, 4);
                            doctor.VerticalAlignment = Element.ALIGN_CENTER;
                            table.AddCell(doctor);

                            //7
                            PdfPCell sick1 = CellFactory("癲病", 3);
                            sick1.FixedHeight = 25f;
                            table.AddCell(sick1);
                            PdfPCell s1 = CellFactory(R000909.Split(':').GetValue(0).ToString(), R000909.Split(':').GetValue(1).ToString(), 3);
                            table.AddCell(s1);
                            //8
                            PdfPCell sick2 = CellFactory("精神病", 3);
                            sick2.FixedHeight = 25f;
                            table.AddCell(sick2);
                            PdfPCell s2 = CellFactory(R000919.Split(':').GetValue(0).ToString(), R000919.Split(':').GetValue(1).ToString(), 3);
                            table.AddCell(s2);
                            //9
                            PdfPCell sick3 = CellFactory("傳染性眼疾", 3);
                            sick3.FixedHeight = 25f;
                            table.AddCell(sick3);
                            PdfPCell s3 = CellFactory(R000920.Split(':').GetValue(0).ToString(), R000920.Split(':').GetValue(1).ToString(), 3);
                            table.AddCell(s3);
                            //9
                            PdfPCell sick4 = CellFactory("傳染性皮膚病", 3);
                            sick4.FixedHeight = 25f;
                            table.AddCell(sick4);
                            PdfPCell s4 = CellFactory(R000921.Split(':').GetValue(0).ToString(), R000921.Split(':').GetValue(1).ToString(), 3);
                            table.AddCell(s4);

                            //check rowspan
                            PdfPCell check = CellFactory("檢\n\n\n驗", 1, 4);
                            table.AddCell(check);

                            count = 0;
                            dt_test_item = test_item_data(pid, arraydate);
                            foreach (DataRow ti_dt in dt_test_item.Rows)
                            {
                                count++;
                                if (ti_dt[0].ToString() == "傷寒凝集試驗Widal test")
                                {
                                    table.AddCell(CellFactory("傷寒", 2));
                                    table.AddCell(CellFactory(ti_dt[3].ToString(), ti_dt[2].ToString(), 1));
                                }
                                else
                                {
                                    table.AddCell(CellFactory(ti_dt[0].ToString(), 2));
                                    table.AddCell(CellFactory(ti_dt[3].ToString(), ti_dt[2].ToString(), 1));
                                }

                            }
                            if (count < 8)
                            {
                                for (int j = 0; j < 8 - count; j++)
                                {
                                    table.AddCell(CellFactory(" ", 2));
                                    table.AddCell(CellFactory(" ", 1));
                                }
                            }
                            //14
                            PdfPCell cell_14 = CellFactory("\n胸部X光\n(肺結核)");
                            cell_14.FixedHeight = 50f;
                            table.AddCell(cell_14);
                            A902 = (A902 == "N") ? "正常" : "異常";
                            PdfPCell cell_14_ct = new PdfPCell(new Phrase(A902, chFont))
                            {
                                Colspan = 6,
                                HorizontalAlignment = Element.ALIGN_CENTER,
                                VerticalAlignment = Element.ALIGN_MIDDLE
                            };
                            table.AddCell(cell_14_ct);
                            //16
                            PdfPCell cell_16 = CellFactory("應 處 理 及\n\n\n注 意 事 項", 2);
                            cell_16.FixedHeight = 100f;
                            table.AddCell(cell_16);

                            string[] suggestArray = suggest.Split(',');
                            Phrase suggest_text = new Phrase();
                            foreach (string suggest_str in suggestArray)
                            {
                                if (suggest_str.Split(':').GetValue(0).ToString() != "")
                                {
                                    suggest_text.Add(new Chunk(suggest_str.Split(':').GetValue(0).ToString(), (suggest_str.Split(':').GetValue(1).ToString() == "Y") ? chFont_red : chFont));
                                    suggest_text.Add(new Chunk("\n", chFont));
                                }
                            }

                            table.AddCell(CellFactory(suggest_text, 6));
                            //rowspan
                            table.AddCell(CellFactory("健\n康\n檢\n查", 1, 4));
                            //17
                            table.AddCell(CellFactory("複查年月日", 1));
                            table.AddCell(CellFactory(" ", 6));

                            //table2
                            PdfPTable table2 = new PdfPTable(2);
                            table2.AddCell(CellFactory("檢\n驗\n項\n目", 1, 1));
                            table2.AddCell(CellFactory("由\n醫\n師\n決\n定", 1, 1));
                            //rowspan
                            PdfPCell cell2 = new PdfPCell(table2);
                            cell2.Rowspan = 3;
                            table.AddCell(cell2);

                            //18
                            var setheight1 = CellFactory(" ", 6);
                            setheight1.FixedHeight = 25f;
                            table.AddCell(setheight1);
                            //19
                            var setheight2 = CellFactory(" ", 6);
                            setheight2.FixedHeight = 25f;
                            table.AddCell(setheight2);
                            //20
                            var setheight3 = CellFactory(" ", 6);
                            setheight3.FixedHeight = 25f;
                            table.AddCell(setheight3);
                            //rowspan
                            table.AddCell(CellFactory("健\n康\n管\n理", 1, 5));
                            //21
                            PdfPCell setheight4 = CellFactory("合格\n判定", 1);
                            setheight4.FixedHeight = 50f;
                            table.AddCell(setheight4);
                            if (R000887 != null)
                            {
                                PdfPCell pass = CellFactory(R000887.Split(':').GetValue(0).ToString(), R000887.Split(':').GetValue(1).ToString(), 6);
                                pass.VerticalAlignment = Element.ALIGN_MIDDLE;
                                pass.HorizontalAlignment = Element.ALIGN_CENTER;
                                table.AddCell(pass);
                            }

                            //rowspan
                            table.AddCell(CellFactory("檢查醫療機構名稱，電話及住址", 4, 3));
                            //23
                            table.AddCell(CellFactory("台北榮民總醫院桃園分院", 3));
                            //24
                            table.AddCell(CellFactory("桃園市成功路三段100號", 3));
                            //25
                            table.AddCell(CellFactory("(03)3384889分機", 3));
                            //26
                            var setheight6 = CellFactory("負責醫師姓名及證書字號", 4);
                            setheight6.FixedHeight = 45f;
                            table.AddCell(setheight6);

                            if (R000910 == "")
                            {
                                table.AddCell(CellFactory(" ", 3));
                            }
                            else
                            {
                                Image img = Image.GetInstance("http://210.61.246.116/drimg/" + R000910 + ".jpg");
                                img.WidthPercentage = 30;
                                PdfPCell imgCell = new PdfPCell();
                                imgCell.Colspan = 3;
                                imgCell.AddElement(img);
                                table.AddCell(imgCell);
                            }



                            ////
                            //for(int i = 1; i <= 12; i++)
                            //{
                            //    for (int j = 1; j <= 7; j++)
                            //    {
                            //        table.AddCell(i+"-"+j);
                            //    }
                            //}


                            doc1.Add(table);
                            doc1.NewPage();




                            //Close our document
                            
                        }

                        doc1.Close();
                    }
                }
            }
            
        }

        private PdfPCell CellFactory(string str , int colspan = 1, int rowspan = 1)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, chFont));
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.MinimumHeight = 25f;
            return cell;
        }
        private PdfPCell CellFactory(Phrase phr, int colspan = 1, int rowspan = 1)
        {
            PdfPCell cell = new PdfPCell(phr);
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.MinimumHeight = 25f;
            return cell;
        }
        private PdfPCell CellFactory(string str, string isunusual, int colspan = 1, int rowspan = 1)
        {
            PdfPCell cell = new PdfPCell(new Phrase(str, (isunusual == "Y") ? chFont_red : chFont));
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = colspan;
            cell.Rowspan = rowspan;
            cell.MinimumHeight = 25f;
            return cell;
        }
        private DataTable test_item_data(string pid,string arraydate)
        {
            DataTable dt1 = new DataTable();

            try
            {
                //Data Fill
                string connectionstring = "server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;";
                using (MySqlConnection cnn = new MySqlConnection(connectionstring))
                {
                    cnn.Open();
                    MySqlCommand command = new MySqlCommand("SELECT hb.name_tc,hb.qty,hb.isunusual,hb.report_text,i.category_tc,hh.arraydate "+
                                                            " from history_body hb"+
                                                            " INNER JOIN(history_head hh, item i) on hb.head_id = hh.id and hb.item_no = i.`no` "+
                                                            " where hh.pid = '"+pid+"' and hh.arraydate = DATE('"+arraydate+"') and(i.category_tc = '肝炎標記' or i.category_tc = '傳染病檢驗' or i.category_tc = '糞便檢驗'); ", cnn);
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
    }
}
