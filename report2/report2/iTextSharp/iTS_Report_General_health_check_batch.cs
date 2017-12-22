using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace report2
{
    class iTS_Report_General_health_check_batch
    {

        BaseFont bfChinese = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        Font chFont;
        Font chFont_UnderLine;
        Font chFont_blue;
        Font chFont_red;
        Font chFont_header;
        Font chFont_small;
        public void ExportPdf(DataTable dt, string dirPath,string dirName)
        {
            string[] str = new string[5];            
            //字型設定
            chFont = new Font(bfChinese, 12);
            chFont_UnderLine = new Font(bfChinese, 12, Font.UNDERLINE);
            chFont_blue = new Font(bfChinese, 8, Font.NORMAL, new BaseColor(0, 0, 240));
            chFont_red = new Font(bfChinese, 12, Font.NORMAL, BaseColor.RED);
            chFont_header = new Font(bfChinese, 18, Font.BOLD, BaseColor.BLACK);
            chFont_small = new Font(bfChinese, 8, Font.NORMAL, BaseColor.BLACK);
            //-------------------------------------
            //dt.Rows[0][1]
            string outputFile = Path.Combine(dirPath,"batch_file.pdf");

            //Create a standard .Net FileStream for the file, setting various flags
            using (FileStream fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                //Create a new PDF document setting the size to A4
                using (Document doc1 = new Document(PageSize.A4, 30, 30, 10, 10))
                {
                    //Bind the PDF document to the FileStream using an iTextSharp PdfWriter
                    using (PdfWriter w = PdfWriter.GetInstance(doc1, fs))
                    {

                        doc1.Open();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {


                            //data 基本資料
                            string name = dt.Rows[i]["name"].ToString();
                            string sex = dt.Rows[i]["sex"].ToString();
                            string birth = Convert.ToDateTime(dt.Rows[i]["birtyday"]).ToString("yyyy/MM/dd");
                            string sid = dt.Rows[i]["sid"].ToString();
                            string emp_no = dt.Rows[i]["emp_no"].ToString();
                            string arraydate = Convert.ToDateTime(dt.Rows[i]["arraydate"]).ToString("yyyy/MM/dd");
                            string cName = dt.Rows[i]["cName"].ToString();
                            string pid = dt.Rows[i]["pid"].ToString();
                            string item_isunusual = dt.Rows[i]["item_isunusual"].ToString();
                            string suggest = dt.Rows[i]["suggest"].ToString();
                            string dept_name = dt.Rows[i]["dept_name"].ToString();
                            string R000001 = dt.Rows[i]["R000001"].ToString();
                            //2-8 作業經歷
                            string R000002 = dt.Rows[i]["R000002"].ToString();
                            string R000003 = dt.Rows[i]["R000003"].ToString();
                            string R000004 = dt.Rows[i]["R000004"].ToString();
                            string R000005 = dt.Rows[i]["R000005"].ToString();
                            string R000006 = dt.Rows[i]["R000006"].ToString();
                            string R000007 = dt.Rows[i]["R000007"].ToString();
                            string R000008 = dt.Rows[i]["R000008"].ToString();
                            //65-71
                            string R000065 = dt.Rows[i]["R000065"].ToString();
                            string R000066 = dt.Rows[i]["R000066"].ToString();
                            string R000067 = dt.Rows[i]["R000067"].ToString();
                            string R000068 = dt.Rows[i]["R000068"].ToString();
                            string R000069 = dt.Rows[i]["R000069"].ToString();
                            string R000070 = dt.Rows[i]["R000070"].ToString();
                            string R000071 = dt.Rows[i]["R000071"].ToString();
                            //953-954 
                            string R000953 = dt.Rows[i]["R000953"].ToString();
                            string R000954 = dt.Rows[i]["R000954"].ToString();
                            //R000036,R000038 檢查時期
                            string R000036 = dt.Rows[i]["R000036"].ToString();
                            string R000038 = dt.Rows[i]["R000038"].ToString();
                            //R000040-R000062 既往病史
                            string R000040 = dt.Rows[i]["R000040"].ToString();
                            string R000041 = dt.Rows[i]["R000041"].ToString();
                            string R000042 = dt.Rows[i]["R000042"].ToString();
                            string R000043 = dt.Rows[i]["R000043"].ToString();
                            string R000043T01 = dt.Rows[i]["R000043T01"].ToString();
                            string R000044 = dt.Rows[i]["R000044"].ToString();
                            string R000045 = dt.Rows[i]["R000045"].ToString();
                            string R000046 = dt.Rows[i]["R000046"].ToString();
                            string R000047 = dt.Rows[i]["R000047"].ToString();
                            string R000048 = dt.Rows[i]["R000048"].ToString();
                            string R000049 = dt.Rows[i]["R000049"].ToString();
                            string R000050 = dt.Rows[i]["R000050"].ToString();
                            string R000051 = dt.Rows[i]["R000051"].ToString();
                            string R000052 = dt.Rows[i]["R000052"].ToString();
                            string R000053 = dt.Rows[i]["R000053"].ToString();
                            string R000054 = dt.Rows[i]["R000054"].ToString();
                            string R000055 = dt.Rows[i]["R000055"].ToString();
                            string R000056 = dt.Rows[i]["R000056"].ToString();
                            string R000057 = dt.Rows[i]["R000057"].ToString();
                            string R000058 = dt.Rows[i]["R000058"].ToString();
                            string R000058T01 = dt.Rows[i]["R000058T01"].ToString();
                            string R000060 = dt.Rows[i]["R000060"].ToString();
                            string R000060T01 = dt.Rows[i]["R000060T01"].ToString();
                            string R000061 = dt.Rows[i]["R000061"].ToString();
                            string R000061T01 = dt.Rows[i]["R000061T01"].ToString();
                            string R000062 = dt.Rows[i]["R000062"].ToString();
                            //生活習慣 吸菸
                            string R000072 = dt.Rows[i]["R000072"].ToString();
                            string R000073 = dt.Rows[i]["R000073"].ToString();
                            string R000074 = dt.Rows[i]["R000074"].ToString();
                            string R000074T01 = dt.Rows[i]["R000074T01"].ToString();
                            string R000074T02 = dt.Rows[i]["R000074T02"].ToString();
                            string R000075 = dt.Rows[i]["R000075"].ToString();
                            string R000075T01 = dt.Rows[i]["R000075T01"].ToString();
                            string R000075T02 = dt.Rows[i]["R000075T02"].ToString();
                            //生活習慣 檳榔
                            string R000076 = dt.Rows[i]["R000076"].ToString();
                            string R000077 = dt.Rows[i]["R000077"].ToString();
                            string R000078 = dt.Rows[i]["R000078"].ToString();
                            string R000078T01 = dt.Rows[i]["R000078T01"].ToString();
                            string R000078T02 = dt.Rows[i]["R000078T02"].ToString();
                            string R000079 = dt.Rows[i]["R000079"].ToString();
                            string R000079T01 = dt.Rows[i]["R000079T01"].ToString();
                            string R000079T02 = dt.Rows[i]["R000079T02"].ToString();
                            //生活習慣 酒
                            string R000080 = dt.Rows[i]["R000080"].ToString();
                            string R000081 = dt.Rows[i]["R000081"].ToString();
                            string R000082 = dt.Rows[i]["R000082"].ToString();
                            string R000082T01 = dt.Rows[i]["R000082T01"].ToString();
                            string R000082T02 = dt.Rows[i]["R000082T02"].ToString();
                            string R000082T03 = dt.Rows[i]["R000082T03"].ToString();
                            string R000083 = dt.Rows[i]["R000083"].ToString();
                            string R000083T01 = dt.Rows[i]["R000083T01"].ToString();
                            string R000083T02 = dt.Rows[i]["R000083T02"].ToString();
                            //生活習慣 睡眠
                            string R000955 = dt.Rows[i]["R000955"].ToString();
                            //自覺症狀
                            string R000084 = dt.Rows[i]["R000084"].ToString();
                            string R000085 = dt.Rows[i]["R000085"].ToString();
                            string R000086 = dt.Rows[i]["R000086"].ToString();
                            string R000087 = dt.Rows[i]["R000087"].ToString();
                            string R000088 = dt.Rows[i]["R000088"].ToString();
                            string R000089 = dt.Rows[i]["R000089"].ToString();
                            string R000090 = dt.Rows[i]["R000090"].ToString();
                            string R000091 = dt.Rows[i]["R000091"].ToString();
                            string R000092 = dt.Rows[i]["R000092"].ToString();
                            string R000093 = dt.Rows[i]["R000093"].ToString();
                            string R000094 = dt.Rows[i]["R000094"].ToString();
                            string R000095 = dt.Rows[i]["R000095"].ToString();
                            string R000096 = dt.Rows[i]["R000096"].ToString();
                            string R000097 = dt.Rows[i]["R000097"].ToString();
                            string R000098 = dt.Rows[i]["R000098"].ToString();
                            string R000099 = dt.Rows[i]["R000099"].ToString();
                            string R000100 = dt.Rows[i]["R000100"].ToString();
                            string R000101 = dt.Rows[i]["R000101"].ToString();
                            string R000102 = dt.Rows[i]["R000102"].ToString();
                            string R000103 = dt.Rows[i]["R000103"].ToString();
                            string R000104 = dt.Rows[i]["R000104"].ToString();
                            string R000105 = dt.Rows[i]["R000105"].ToString();
                            string R000106 = dt.Rows[i]["R000106"].ToString();
                            string R000106T01 = dt.Rows[i]["R000106T01"].ToString();
                            string R000107 = dt.Rows[i]["R000107"].ToString();
                            //檢查項目
                            string A2110101 = dt.Rows[i]["A2110101"].ToString();
                            string A2110110 = dt.Rows[i]["A2110110"].ToString();
                            string A2110102 = dt.Rows[i]["A2110102"].ToString();
                            string A2110402 = dt.Rows[i]["A2110402"].ToString();
                            string A2110107 = dt.Rows[i]["A2110107"].ToString();
                            string A2110111 = dt.Rows[i]["A2110111"].ToString();
                            string A2110106 = dt.Rows[i]["A2110106"].ToString();
                            //眼睛
                            string A2110201 = dt.Rows[i]["A2110201"].ToString();
                            string A2110202 = dt.Rows[i]["A2110202"].ToString();
                            string A2110203 = dt.Rows[i]["A2110203"].ToString();
                            string A2110204 = dt.Rows[i]["A2110204"].ToString();

                            string A2110209 = dt.Rows[i]["A2110209"].ToString();
                            //耳朵
                            string A2110316 = dt.Rows[i]["A2110316"].ToString();
                            string A2110317 = dt.Rows[i]["A2110317"].ToString();
                            string R000583 = dt.Rows[i]["R000583"].ToString();
                            string R000142 = dt.Rows[i]["R000142"].ToString();
                            string R000143 = dt.Rows[i]["R000143"].ToString();
                            string R000144 = dt.Rows[i]["R000144"].ToString();
                            string R000145 = dt.Rows[i]["R000145"].ToString();
                            string R000146 = dt.Rows[i]["R000146"].ToString();
                            string R000147 = dt.Rows[i]["R000147"].ToString();
                            string A902 = dt.Rows[i]["A902"].ToString();
                            string M0230002 = dt.Rows[i]["M0230002"].ToString();
                            string M0230003 = dt.Rows[i]["M0230003"].ToString();
                            string M0060003 = dt.Rows[i]["M0060003"].ToString();
                            string M0060001 = dt.Rows[i]["M0060001"].ToString();
                            string M0120001 = dt.Rows[i]["M0120001"].ToString();
                            string M0070001 = dt.Rows[i]["M0070001"].ToString();
                            string M0100002 = dt.Rows[i]["M0100002"].ToString();
                            string M0130001 = dt.Rows[i]["M0130001"].ToString();
                            string M0130002 = dt.Rows[i]["M0130002"].ToString();
                            string M0130003 = dt.Rows[i]["M0130003"].ToString();
                            string A3020502 = dt.Rows[i]["A3020502"].ToString();
                            string M0100003 = dt.Rows[i]["M0100003"].ToString();
                            //電子章
                            string R000938 = dt.Rows[i]["R000938"].ToString();
                            string R000910 = dt.Rows[i]["R000910"].ToString();







                            //Open the document for writing
                            

                            //header
                            Paragraph header = new Paragraph(10f, "勞工一般體格及健康檢查紀錄", chFont_header);
                            header.Alignment = Element.ALIGN_CENTER;
                            header.SpacingAfter = 15f;
                            doc1.Add(header);
                            //table1
                            PdfPTable table1 = new PdfPTable(2);
                            table1.TotalWidth = 500f;
                            table1.LockedWidth = true;


                            table1.AddCell(CellFactory(new Phrase(cName + "\n" + dept_name, chFont_small), 1, 1, true));
                            table1.AddCell(CellFactory(new Phrase("體檢編號:" + sid + "\n" + "員工編號:" + emp_no, chFont_small), 1, 1, false));
                            doc1.Add(table1);

                            //項目一
                            Paragraph item1 = new Paragraph(10f, "一、基本資料", chFont);
                            doc1.Add(item1);
                            PdfPTable table2 = new PdfPTable(new float[] { 4, 3 });
                            table2.TotalWidth = 450f;
                            table2.LockedWidth = true;
                            table2.HorizontalAlignment = Element.ALIGN_CENTER;
                            table2.AddCell(CellFactory("1.姓名:" + name));
                            table2.AddCell(CellFactory((sex == "M") ? "2.性別:■男 □女" : "2.性別:□男 ■女"));
                            table2.AddCell(CellFactory("3.身分證字號:" + pid));
                            table2.AddCell(CellFactory("4.出生日期:" + birth));
                            table2.AddCell(CellFactory("5.受雇日期:" + R000001));
                            table2.AddCell(CellFactory("6.檢查日期:" + arraydate));
                            table2.SpacingAfter = 10f;
                            table2.SpacingBefore = 10f;
                            doc1.Add(table2);

                            //項目二
                            Paragraph item2 = new Paragraph(10f, "二、作業經歷", chFont);
                            item2.SpacingAfter = 15f;
                            doc1.Add(item2);


                            Phrase p1 = new Phrase();
                            p1.Add(new Chunk("1.曾經從事", chFont));
                            p1.Add(new Chunk(" " + R000002 + " ", chFont_UnderLine));
                            p1.Add(new Chunk("，起始日期", chFont));
                            p1.Add(new Chunk(" " + R000003 + "/" + R000004 + " ", chFont_UnderLine));
                            p1.Add(new Chunk("截止日期", chFont));
                            p1.Add(new Chunk(" " + R000005 + "/" + R000006 + " ", chFont_UnderLine));
                            p1.Add(new Chunk("共", chFont));
                            p1.Add(new Chunk(" " + R000007 + "年" + R000008 + "個月" + " ", chFont_UnderLine));
                            Paragraph paragraph = new Paragraph(p1);
                            paragraph.Alignment = Element.ALIGN_JUSTIFIED;
                            paragraph.FirstLineIndent = 20f;
                            paragraph.SetLeading(0.0f, 1.0f);
                            paragraph.SpacingAfter = 15f;
                            doc1.Add(paragraph);


                            Phrase p2 = new Phrase();
                            p2.Add(new Chunk("2.目前從事", chFont));
                            p2.Add(new Chunk(" " + R000065 + " ", chFont_UnderLine));
                            p2.Add(new Chunk("，起始日期", chFont));
                            p2.Add(new Chunk(" " + R000066 + "/" + R000067 + " ", chFont_UnderLine));
                            p2.Add(new Chunk("截止日期", chFont));
                            p2.Add(new Chunk(" " + R000068 + "/" + R000069 + " ", chFont_UnderLine));
                            p2.Add(new Chunk("共", chFont));
                            p2.Add(new Chunk(" " + R000070 + "年" + R000071 + "個月" + " ", chFont_UnderLine));

                            Paragraph paragraph2 = new Paragraph(p2);
                            paragraph2.Alignment = Element.ALIGN_JUSTIFIED;
                            paragraph2.FirstLineIndent = 20f;
                            paragraph2.SetLeading(0.0f, 1.0f);
                            paragraph2.SpacingAfter = 15f;
                            doc1.Add(paragraph2);

                            //Chunk i2_31 = new Chunk("", chFont);

                            Phrase p3 = new Phrase();
                            p3.Add(new Chunk("3.過去一個月，平均每週工作:", chFont));
                            p3.Add(new Chunk(" " + R000953 + " ", chFont_UnderLine));
                            p3.Add(new Chunk("小時;過去六個月，平均每週工作:", chFont));
                            p3.Add(new Chunk(" " + R000954 + " ", chFont_UnderLine));
                            p3.Add(new Chunk("小時;", chFont));

                            Paragraph paragraph3 = new Paragraph(p3);
                            paragraph3.Alignment = Element.ALIGN_JUSTIFIED;
                            paragraph3.FirstLineIndent = 20f;
                            paragraph3.SetLeading(0.0f, 1.0f);
                            paragraph3.SpacingAfter = 25f;
                            doc1.Add(paragraph3);

                            //項目三
                            Phrase p4 = new Phrase();
                            p4.Add(new Chunk("三、檢查時期(原因):", chFont));
                            p4.Add(new Chunk((R000036 == "true") ? "■" : "□", chFont));
                            p4.Add(new Chunk("新進員工(受雇時) ", chFont));
                            p4.Add(new Chunk((R000038 == "true") ? "■" : "□", chFont));
                            p4.Add(new Chunk("定期年度檢查", chFont));
                            Paragraph paragraph4 = new Paragraph(p4);
                            paragraph4.Alignment = Element.ALIGN_JUSTIFIED;
                            paragraph4.SetLeading(0.0f, 0.2f);
                            paragraph4.SpacingAfter = 15f;
                            doc1.Add(paragraph4);
                            //項目四
                            Paragraph item4 = new Paragraph(10f, "四、既往病史", chFont);
                            doc1.Add(item4);
                            Paragraph item4_1 = new Paragraph("你是否曾患有下列慢性疾病:(請在適合的項目打勾)", chFont);
                            item4_1.FirstLineIndent = 20f;
                            doc1.Add(item4_1);
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000040 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("高血壓  ", chFont));
                            doc1.Add(new Chunk((R000041 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("糖尿病  ", chFont));
                            doc1.Add(new Chunk((R000042 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("心臟病  ", chFont));
                            doc1.Add(new Chunk((R000043 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("癌症:", chFont)); doc1.Add(new Chunk(R000043T01, chFont));
                            doc1.Add(new Chunk((R000044 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("白內障  ", chFont));
                            doc1.Add(new Chunk((R000045 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("中風  ", chFont));
                            doc1.Add(new Chunk((R000046 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("癲癇  ", chFont));
                            doc1.Add(new Chunk((R000047 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("氣喘\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000048 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("慢性氣管炎、肺氣腫  ", chFont));
                            doc1.Add(new Chunk((R000049 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("肺結核  ", chFont));
                            doc1.Add(new Chunk((R000050 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("腎臟病  ", chFont));
                            doc1.Add(new Chunk((R000051 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("肝病  ", chFont));
                            doc1.Add(new Chunk((R000052 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("貧血  ", chFont));
                            doc1.Add(new Chunk((R000053 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("中耳炎  ", chFont));
                            doc1.Add(new Chunk((R000054 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("聽力障礙\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000055 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("甲狀腺疾病  ", chFont));
                            doc1.Add(new Chunk((R000056 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("消化性潰瘍、胃炎  ", chFont));
                            doc1.Add(new Chunk((R000057 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("逆流性食道炎  ", chFont));
                            doc1.Add(new Chunk((R000058 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("骨折:  ", chFont)); doc1.Add(new Chunk(R000058T01 + "\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000060 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("手術開刀:  ", chFont)); doc1.Add(new Chunk(R000060T01 + "\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000061 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("其他慢性病:  ", chFont)); doc1.Add(new Chunk(R000061T01 + "\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk(R000062 != "true" ? "□" : "■", chFont)); doc1.Add(new Chunk("以上皆無 \n", chFont));
                            doc1.Add(new Chunk("      "));
                            //項目五
                            Paragraph item5 = new Paragraph(10f, "五、生活習慣", chFont);
                            doc1.Add(item5);
                            doc1.Add(new Paragraph("1.請問您過去一個月是否有吸菸?", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000072 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("從未吸菸  ", chFont));
                            doc1.Add(new Chunk((R000073 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("偶而吸(不是天天)  ", chFont));
                            doc1.Add(new Chunk((R000074 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("(幾乎)每天吸，平均每天吸 ", chFont));
                            doc1.Add(new Chunk(" " + R000074T01 + " ", chFont_UnderLine)); doc1.Add(new Chunk("支，已吸菸", chFont));
                            doc1.Add(new Chunk(" " + R000074T02 + " ", chFont_UnderLine)); doc1.Add(new Chunk("年\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000075 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("已經戒菸，戒了", chFont));
                            doc1.Add(new Chunk(" " + R000075T01 + " ", chFont_UnderLine));
                            doc1.Add(new Chunk("年", chFont));
                            doc1.Add(new Chunk(" " + R000075T02 + " ", chFont_UnderLine));
                            doc1.Add(new Chunk("個月", chFont));

                            doc1.Add(new Paragraph("2.請問您最近六個月內是否有嚼食檳榔?", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000076 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("從未嚼食檳榔  ", chFont));
                            doc1.Add(new Chunk((R000077 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("偶而嚼(不是天天)  ", chFont));
                            doc1.Add(new Chunk((R000078 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("(幾乎)每天嚼，平均每天嚼 ", chFont));
                            doc1.Add(new Chunk(" " + R000078T01 + " ", chFont_UnderLine)); doc1.Add(new Chunk("顆，已嚼", chFont));
                            doc1.Add(new Chunk(" " + R000078T02 + " ", chFont_UnderLine)); doc1.Add(new Chunk("年\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000079 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("已經戒食，戒了", chFont));
                            doc1.Add(new Chunk(" " + R000079T01 + " ", chFont_UnderLine));
                            doc1.Add(new Chunk("年", chFont));
                            doc1.Add(new Chunk(" " + R000079T02 + " ", chFont_UnderLine));
                            doc1.Add(new Chunk("個月", chFont));

                            doc1.Add(new Paragraph("3.請問過去一個月是否有喝酒?", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000080 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("從未喝酒 ", chFont));
                            doc1.Add(new Chunk((R000081 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("偶而喝(不是天天)\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000082 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("(幾乎)每天喝，平均每周喝 ", chFont));
                            doc1.Add(new Chunk(" " + R000082T01 + " ", chFont_UnderLine)); doc1.Add(new Chunk("次，最常喝 ", chFont));
                            doc1.Add(new Chunk(" " + R000082T02 + " ", chFont_UnderLine)); doc1.Add(new Chunk("酒，每次 ", chFont));
                            doc1.Add(new Chunk(" " + R000082T03 + " ", chFont_UnderLine)); doc1.Add(new Chunk("瓶 \n", chFont));

                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000083 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("已經戒酒，戒了", chFont));
                            doc1.Add(new Chunk(" " + R000083T01 + " ", chFont_UnderLine)); doc1.Add(new Chunk("年 ", chFont));
                            doc1.Add(new Chunk(" " + R000083T02 + " ", chFont_UnderLine)); doc1.Add(new Chunk("月 ", chFont));
                            doc1.Add(new Chunk("      "));
                            Chunk c1 = new Chunk("4.您於工作期間，每日睡眠時間為：", chFont);
                            Chunk c2 = new Chunk(" " + R000955 + " ", chFont_UnderLine);
                            Chunk c3 = new Chunk("小時", chFont);
                            Paragraph p = new Paragraph();
                            p.Add(c1);
                            p.Add(c2);
                            p.Add(c3);
                            doc1.Add(p);
                            //項目六
                            doc1.Add(new Paragraph("六、自覺症狀:你最近三個月是否常有下列症狀:(請在適當項目打勾)", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000084 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("咳嗽 ", chFont));
                            doc1.Add(new Chunk((R000085 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("咳痰 ", chFont));
                            doc1.Add(new Chunk((R000086 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("呼吸困難 ", chFont));
                            doc1.Add(new Chunk((R000087 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("胸痛 ", chFont));
                            doc1.Add(new Chunk((R000088 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("心悸 ", chFont));
                            doc1.Add(new Chunk((R000089 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("頭暈 ", chFont));
                            doc1.Add(new Chunk((R000090 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("頭痛 ", chFont));
                            doc1.Add(new Chunk((R000091 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("耳鳴 ", chFont));
                            doc1.Add(new Chunk((R000092 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("倦怠 ", chFont));
                            doc1.Add(new Chunk((R000093 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("噁心 \n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000094 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("腹痛 ", chFont));
                            doc1.Add(new Chunk((R000095 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("便祕 ", chFont));
                            doc1.Add(new Chunk((R000096 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("腹瀉 ", chFont));
                            doc1.Add(new Chunk((R000097 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("血便 ", chFont));
                            doc1.Add(new Chunk((R000098 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("上背痛 ", chFont));
                            doc1.Add(new Chunk((R000099 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("下背痛 ", chFont));
                            doc1.Add(new Chunk((R000100 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("手腳麻痛 ", chFont));
                            doc1.Add(new Chunk((R000101 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("關節疼痛\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000102 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("排尿不適 ", chFont));
                            doc1.Add(new Chunk((R000103 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("多尿、頻尿 ", chFont));
                            doc1.Add(new Chunk((R000104 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("手腳肌肉無力 ", chFont));
                            doc1.Add(new Chunk((R000105 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("體重減輕3公斤以上\n", chFont));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000106 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("其他症狀 ", chFont)); doc1.Add(new Chunk(" " + R000106T01 + " \n", chFont_UnderLine));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk((R000107 != "true") ? "□" : "■", chFont)); doc1.Add(new Chunk("以上皆無 ", chFont));

                            PdfPTable lable = new PdfPTable(1);
                            lable.TotalWidth = 45f;
                            lable.LockedWidth = true;
                            lable.HorizontalAlignment = Element.ALIGN_LEFT;
                            PdfPCell cell1 = new PdfPCell(new Phrase("說明", chFont));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            lable.AddCell(cell1);
                            doc1.Add(lable);

                            Paragraph description = new Paragraph(10f, "一、請受檢員工於勞工健檢前，填妥基本資料、作業經歷、檢查時期、既往病史、生活習慣及自覺症狀六大項，" +
                                "再交由醫護人員做確認，以有效篩檢出疾病；若事業單位以提供受檢員工基本資料及作業經歷電子檔給指定醫療機構，可不必請受檢員工重複填寫。", chFont_small);
                            doc1.Add(description);
                            Paragraph description2 = new Paragraph(10f, "二、自覺症狀乙項，請受檢者依自身實際症狀勾選", chFont_small);
                            doc1.Add(description2);

                            //===換頁===
                            doc1.NewPage();
                            //===換頁===

                            Paragraph lable2 = new Paragraph("========以下由醫護人員填寫========", chFont);
                            lable2.Alignment = Element.ALIGN_CENTER;
                            doc1.Add(lable2);

                            //項目七
                            doc1.Add(new Paragraph("七、檢驗項目", chFont));
                            doc1.Add(new Chunk("1.身高：", chFont));
                            if (A2110101 != "0")
                            { doc1.Add(new Chunk(A2110101.Split(':').GetValue(0).ToString(), (A2110101.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red)); doc1.Add(new Chunk("公分　", chFont)); }

                            doc1.Add(new Chunk("腰圍：", chFont));
                            if (A2110110 != "0")
                            { doc1.Add(new Chunk(A2110110.Split(':').GetValue(0).ToString(), (A2110110.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red)); doc1.Add(new Chunk("公分　", chFont)); }

                            doc1.Add(new Chunk("男0-89.9公分，女0-79公分\n", chFont_blue));

                            doc1.Add(new Chunk("2.體重：", chFont));
                            if (A2110102 != "0")
                                doc1.Add(new Chunk(A2110102.Split(':').GetValue(0).ToString(), (A2110102.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("公斤　", chFont));
                            doc1.Add(new Chunk("身體質量指數BMI：", chFont));
                            if (A2110402 != "0")
                                doc1.Add(new Chunk(A2110402.Split(':').GetValue(0).ToString(), (A2110402.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("18.5-23.99\n", chFont_blue));

                            doc1.Add(new Chunk("3.血壓：", chFont));
                            if (A2110107 != "0")
                                doc1.Add(new Chunk(A2110107.Split(':').GetValue(0).ToString(), (A2110107.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("mmHg/", chFont));
                            if (A2110111 != "0")
                                doc1.Add(new Chunk(A2110111.Split(':').GetValue(0).ToString(), (A2110111.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("mmHg", chFont));

                            doc1.Add(new Chunk(" 90-139/60-89mmHg ", chFont_blue));
                            doc1.Add(new Chunk("脈搏", chFont));
                            if (A2110106 != "0")
                                doc1.Add(new Chunk(A2110106.Split(':').GetValue(0).ToString(), (A2110106.Split(':').GetValue(0).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("次/分", chFont));
                            doc1.Add(new Chunk(" 60-100次/分\n", chFont_blue));

                            doc1.Add(new Chunk("4.視力：", chFont));

                            doc1.Add(new Chunk((A2110201.Split(':').GetValue(1).ToString() == "") ? A2110203.Split(':').GetValue(0).ToString() : A2110201.Split(':').GetValue(0).ToString(), chFont));
                            if (A2110203 != "0")
                                doc1.Add(new Chunk((A2110201.Split(':').GetValue(1).ToString() == "") ? A2110203.Split(':').GetValue(1).ToString() : A2110201.Split(':').GetValue(1).ToString()
                                    , (((A2110201.Split(':').GetValue(1).ToString() == "") ? A2110203.Split(':').GetValue(2).ToString() : A2110201.Split(':').GetValue(2).ToString()) != "Y") ? chFont : chFont_red));

                            doc1.Add(new Chunk((A2110202.Split(':').GetValue(1).ToString() == "") ? A2110204.Split(':').GetValue(0).ToString() : A2110202.Split(':').GetValue(0).ToString(), chFont));
                            if (A2110204 != "0")
                                doc1.Add(new Chunk((A2110202.Split(':').GetValue(1).ToString() == "") ? A2110204.Split(':').GetValue(1).ToString() : A2110202.Split(':').GetValue(1).ToString()
                                    , (((A2110202.Split(':').GetValue(1).ToString() == "") ? A2110204.Split(':').GetValue(2).ToString() : A2110202.Split(':').GetValue(2).ToString()) != "Y") ? chFont : chFont_red));

                            doc1.Add(new Chunk(" 0.6-1.5          ", chFont_blue));
                            doc1.Add(new Chunk("辨色能力測試：", chFont));
                            if (A2110209 != "0")
                                doc1.Add(new Chunk(A2110209.Split(':').GetValue(0).ToString() + "\n", (A2110209.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));

                            doc1.Add(new Chunk("5.聽力檢查：", chFont));
                            doc1.Add(new Chunk("左耳:", chFont));
                            if (A2110316 != "0")
                                doc1.Add(new Chunk(A2110316.Split(':').GetValue(0).ToString(), (A2110316.Split(':').GetValue(0).ToString() == "Y") ? chFont_red : chFont));
                            doc1.Add(new Chunk("                ", chFont));
                            doc1.Add(new Chunk("右耳:", chFont));
                            if (A2110317 != "0")
                                doc1.Add(new Chunk(A2110317.Split(':').GetValue(0).ToString(), (A2110317.Split(':').GetValue(0).ToString() == "Y") ? chFont_red : chFont));
                            doc1.Add(new Chunk("\n"));

                            doc1.Add(new Chunk("6.各系統或部位理學檢查：", chFont));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(1)頭頸部(結膜、淋巴腺、甲狀腺)：", chFont));
                            if (R000583 != "")
                                doc1.Add(new Chunk((R000583.Split(':').GetValue(1).ToString() == "Y") ? R000583.Split(':').GetValue(0).ToString() : "無異常", (R000583.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(2)呼吸系統：", chFont));
                            if (R000142 != "")
                                doc1.Add(new Chunk((R000142.Split(':').GetValue(1).ToString() == "Y") ? R000142.Split(':').GetValue(0).ToString() : "無異常", (R000142.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(3)心臟血管系統(心律、心雜音)：", chFont));
                            if (R000143 != "")
                                doc1.Add(new Chunk((R000143.Split(':').GetValue(1).ToString() == "Y") ? R000143.Split(':').GetValue(0).ToString() : "無異常", (R000143.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(4)消化系統(黃膽、肝臟、腹部)：", chFont));
                            if (R000144 != "")
                                doc1.Add(new Chunk((R000144.Split(':').GetValue(1).ToString() == "Y") ? R000144.Split(':').GetValue(0).ToString() : "無異常", (R000144.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(5)神經系統(感覺)：", chFont));
                            if (R000145 != "")
                                doc1.Add(new Chunk((R000145.Split(':').GetValue(1).ToString() == "Y") ? R000145.Split(':').GetValue(0).ToString() : "無異常", (R000145.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(6)肌肉骨骼(四肢)：", chFont));
                            if (R000146 != "")
                                doc1.Add(new Chunk((R000146.Split(':').GetValue(1).ToString() == "Y") ? R000146.Split(':').GetValue(0).ToString() : "無異常", (R000146.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));
                            doc1.Add(new Chunk("      "));
                            doc1.Add(new Chunk("(7)皮膚：", chFont));
                            if (R000147 != "")
                                doc1.Add(new Chunk((R000147.Split(':').GetValue(1).ToString() == "Y") ? R000147.Split(':').GetValue(0).ToString() : "無異常", (R000147.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            doc1.Add(new Chunk("\n"));

                            //table
                            PdfPTable table3 = new PdfPTable(new float[] { 5, 4 });
                            table3.TotalWidth = 540f;
                            table3.LockedWidth = true;
                            table3.HorizontalAlignment = Element.ALIGN_LEFT;

                            Phrase phrase_x = new Phrase();
                            phrase_x.Add(new Chunk("7.胸部X光：", chFont));
                            phrase_x.Add(new Chunk(A902.Split(':').GetValue(0).ToString(), (A902.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            table3.AddCell(CellFactory(phrase_x, 2));

                            Phrase phrase_u1 = new Phrase();
                            phrase_u1.Add(new Chunk("8.尿液檢查：尿蛋白： ", chFont));
                            phrase_u1.Add(new Chunk(M0230002.Split(':').GetValue(0).ToString(), (M0230002.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            table3.AddCell(CellFactory(phrase_u1));
                            Phrase phrase_u2 = new Phrase();
                            phrase_u2.Add(new Chunk("尿潛血： ", chFont));
                            phrase_u2.Add(new Chunk(M0230003.Split(':').GetValue(0).ToString(), (M0230003.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            table3.AddCell(CellFactory(phrase_u2));

                            Phrase phrase_b1 = new Phrase();
                            phrase_b1.Add(new Chunk("9.血液檢查：血色素： ", chFont));
                            if (M0060003 != "")
                                phrase_b1.Add(new Chunk(M0060003.Split(':').GetValue(0).ToString(), (M0060003.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            if (M0060003 != "")
                                phrase_b1.Add(new Chunk(" " + M0060003.Split(':').GetValue(2).ToString() + "-" + M0060003.Split(':').GetValue(3).ToString() + " g/dl", chFont_blue));
                            table3.AddCell(CellFactory(phrase_b1));
                            Phrase phrase_b2 = new Phrase();
                            phrase_b2.Add(new Chunk("白血球： ", chFont));
                            if (M0060001 != "0")
                                phrase_b2.Add(new Chunk(M0060001.Split(':').GetValue(0).ToString(), (M0060001.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            phrase_b2.Add(new Chunk(" /ul", chFont));
                            if (M0060001 != "0")
                                phrase_b2.Add(new Chunk(" " + M0060001.Split(':').GetValue(2).ToString() + "-" + M0060001.Split(':').GetValue(3).ToString() + " /ul", chFont_blue));
                            table3.AddCell(CellFactory(phrase_b2));

                            Phrase phrase1 = new Phrase();
                            phrase1.Add(new Chunk("10.生化血液檢查：血糖： ", chFont));
                            if (M0120001 != "0")
                                phrase1.Add(new Chunk(M0120001.Split(':').GetValue(0).ToString(), (M0120001.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));//
                            phrase1.Add(new Chunk(" mg/dL", chFont));
                            if (M0120001 != "0")
                                phrase1.Add(new Chunk(" " + M0120001.Split(':').GetValue(2).ToString() + "-" + M0120001.Split(':').GetValue(3).ToString() + "mg/dL", chFont_blue));
                            table3.AddCell(CellFactory(phrase1));

                            Phrase phrase2 = new Phrase();
                            phrase2.Add(new Chunk("麩胺酸丙酮酸轉胺基脢(ALT) ", chFont));
                            if (M0070001 != "0")
                                phrase2.Add(new Chunk(M0070001.Split(':').GetValue(0).ToString(), (M0070001.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));//
                            phrase2.Add(new Chunk(" U/L", chFont));
                            if (M0070001 != "0")
                                phrase2.Add(new Chunk(" " + M0070001.Split(':').GetValue(2).ToString() + "-" + M0070001.Split(':').GetValue(3).ToString() + "U/L", chFont_blue));
                            phrase2.Add(new Chunk("\n"));
                            table3.AddCell(CellFactory(phrase2));

                            Phrase phrase3 = new Phrase();
                            phrase3.Add(new Chunk("      "));
                            phrase3.Add(new Chunk("肌酸酐： ", chFont));
                            if (M0100002 != "0")
                                phrase3.Add(new Chunk(M0100002.Split(':').GetValue(0).ToString(), (M0100002.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            phrase3.Add(new Chunk(" mg/dl", chFont));
                            if (M0100002 != "0")
                                phrase3.Add(new Chunk(" " + M0100002.Split(':').GetValue(2).ToString() + "-" + M0100002.Split(':').GetValue(3).ToString() + " mg/dl", chFont_blue));
                            table3.AddCell(CellFactory(phrase3));

                            Phrase phrase4 = new Phrase();
                            phrase4.Add(new Chunk("膽固醇： ", chFont));
                            if (M0130001 != "0")
                                phrase4.Add(new Chunk(M0130001.Split(':').GetValue(0).ToString(), (M0130001.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            phrase4.Add(new Chunk(" mg/dl", chFont));
                            if (M0130001 != "0")
                                phrase4.Add(new Chunk(" " + M0130001.Split(':').GetValue(2).ToString() + "-" + M0130001.Split(':').GetValue(3).ToString() + " mg/dl", chFont_blue));
                            table3.AddCell(CellFactory(phrase4));

                            Phrase phrase5 = new Phrase();
                            phrase5.Add(new Chunk("      "));
                            phrase5.Add(new Chunk("高密度脂蛋白膽固醇： ", chFont));
                            if (M0130002 != "0")
                                phrase5.Add(new Chunk(M0130002.Split(':').GetValue(0).ToString(), (M0130002.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));//
                            phrase5.Add(new Chunk(" mg/dl", chFont));
                            if (M0130002 != "0")
                                phrase5.Add(new Chunk(" " + M0130002.Split(':').GetValue(2).ToString() + " mg /dl以上", chFont_blue));
                            table3.AddCell(CellFactory(phrase5));

                            Phrase phrase6 = new Phrase();
                            phrase6.Add(new Chunk("低密度脂蛋白膽固醇： ", chFont));
                            if (M0130003 != "0")
                                phrase6.Add(new Chunk(M0130003.Split(':').GetValue(0).ToString(), (M0130003.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));//
                            phrase6.Add(new Chunk(" mg/dl", chFont));
                            if (M0130003 != "0")
                                phrase6.Add(new Chunk(" " + M0130003.Split(':').GetValue(2).ToString() + "-" + M0130003.Split(':').GetValue(3).ToString() + " mg/dl", chFont_blue));
                            phrase6.Add(new Chunk("\n"));
                            table3.AddCell(CellFactory(phrase6));

                            Phrase phrase7 = new Phrase();
                            phrase7.Add(new Chunk("      "));
                            phrase7.Add(new Chunk("三酸甘油脂： ", chFont));
                            if (A3020502 != "0")
                                phrase7.Add(new Chunk(A3020502.Split(':').GetValue(0).ToString(), (A3020502.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));//
                            phrase7.Add(new Chunk(" mg/dl", chFont));
                            if (A3020502 != "0")
                                phrase7.Add(new Chunk(" " + A3020502.Split(':').GetValue(2).ToString() + "-" + A3020502.Split(':').GetValue(3).ToString() + " mg/dl", chFont_blue));
                            table3.AddCell(CellFactory(phrase7));

                            Phrase phrase8 = new Phrase();
                            phrase8.Add(new Chunk("腎絲球過濾率eGFR： ", chFont));
                            if (M0100003 != "0")
                                phrase8.Add(new Chunk(M0100003.Split(':').GetValue(0).ToString(), (M0100003.Split(':').GetValue(1).ToString() != "Y") ? chFont : chFont_red));
                            phrase8.Add(new Chunk(" mg/dl", chFont));
                            if (M0100003 != "0")
                                phrase8.Add(new Chunk(">60ml/min/1.7", chFont_blue));
                            table3.AddCell(CellFactory(phrase8));
                            doc1.Add(table3);
                            doc1.Add(new Paragraph("11.其他經中央主管機關規定之檢查", chFont));
                            //項目八
                            doc1.Add(new Paragraph("八.應處理及注意事項(可複選)", chFont_small));

                            //doc1.Add(new Chunk("    ", chFont_small));
                            doc1.Add(new Chunk("1.", chFont_small));
                            doc1.Add(new Chunk((item_isunusual != "正常") ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("檢查結果大致正常，請定期健康檢查。\n", chFont_small));

                            //doc1.Add(new Chunk("    ", chFont_small));
                            doc1.Add(new Chunk("2.", chFont_small));
                            doc1.Add(new Chunk((true) ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("檢查結果部分異常，宜在()內至醫療機構  科，實施健康追蹤檢查。\n", chFont_small));

                            //doc1.Add(new Chunk("    ", chFont_small));
                            doc1.Add(new Chunk("3.", chFont_small));
                            doc1.Add(new Chunk((true) ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("檢查結果異常，建議不適宜從事。", chFont_small)); doc1.Add(new Chunk("  ", chFont_UnderLine));
                            doc1.Add(new Chunk("作業\n", chFont_small));

                            //doc1.Add(new Chunk("    ", chFont_small));
                            doc1.Add(new Chunk("4.", chFont_small));
                            doc1.Add(new Chunk((item_isunusual == "正常") ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("檢查結果異常，建議調整工作(可複選)\n", chFont_small));

                            doc1.Add(new Chunk("        ", chFont_small));
                            doc1.Add(new Chunk((true) ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("縮短工作時間(請說明原因： )\n", chFont_small));

                            doc1.Add(new Chunk("        ", chFont_small));
                            doc1.Add(new Chunk((true) ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("更新工作內容(請說明原因： )\n", chFont_small));

                            doc1.Add(new Chunk("        ", chFont_small));
                            doc1.Add(new Chunk((true) ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("變更作業場所(請說明原因： )\n", chFont_small));

                            doc1.Add(new Chunk("        ", chFont_small));
                            doc1.Add(new Chunk((item_isunusual == "正常") ? "□" : "■", chFont_small));
                            doc1.Add(new Chunk("其他:＊異常建議尋求醫療協助由醫師面談及健康指導，工作時如有發生身體不適請立即休息並建議尋求醫療協助\n", chFont_small));

                            //doc1.Add(new Chunk("    ", chFont_small));
                            doc1.Add(new Chunk("5.", chFont_small));
                            doc1.Add(new Chunk((item_isunusual == "正常") ? "□" : "■", chFont_small));
                            suggest = suggest.Split('\n').GetValue(0).ToString();
                            doc1.Add(new Chunk("其他：" + suggest + "\n", chFont_small));

                            PdfPTable table4 = new PdfPTable(new float[] { 2, 3 });
                            table4.TotalWidth = 500f;
                            table4.LockedWidth = true;
                            Phrase phrase9 = new Phrase("健檢醫師姓名(簽章)及證書字號：", chFont);
                            table4.AddCell(CellFactory(phrase9));

                            if (R000938 != "")
                            {
                                Image img1 = Image.GetInstance("http://210.61.246.116/drimg/" + R000938 + ".jpg");
                                img1.WidthPercentage = 30;
                                PdfPCell imgCell1 = new PdfPCell();
                                imgCell1.Border = 0;
                                imgCell1.AddElement(img1);

                                table4.AddCell(imgCell1);
                            }
                            else
                            {
                                table4.AddCell(CellFactory(" "));
                            }

                            Phrase phrase10 = new Phrase("總評醫師姓名(簽章)及證書字號：", chFont);
                            table4.AddCell(CellFactory(phrase10));

                            if (R000910 != "")
                            {
                                Image img2 = Image.GetInstance("http://210.61.246.116/drimg/" + R000910 + ".jpg");
                                img2.WidthPercentage = 30;
                                PdfPCell imgCell2 = new PdfPCell();
                                imgCell2.Border = 0;
                                imgCell2.AddElement(img2);

                                table4.AddCell(imgCell2);
                            }
                            else
                            {
                                table4.AddCell(CellFactory(" "));
                            }
                            doc1.Add(table4);

                            Paragraph paragraph7 = new Paragraph("健檢機構名稱、電話、地址：", chFont);
                            doc1.Add(paragraph7);

                            Paragraph paragraph8 = new Paragraph("台北榮民總醫院桃園分院;地址:桃園市成功路三段100號;電話:(03)3384889", chFont);
                            paragraph8.Alignment = Element.ALIGN_CENTER;
                            doc1.Add(paragraph8);

                            doc1.NewPage();

                            
                        }
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
    }
}
