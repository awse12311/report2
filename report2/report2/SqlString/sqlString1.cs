using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace report2.SqlString
{
    class sqlString1
    {
        //2 sql 欄位字串
        private string getSqlString(string reportSelect)
        {
            string sql = " select " +
                            " p.id,p.`Name`,p.Pid,p.Sex,DATE_FORMAT(p.Birtyday, '%Y/%m/%d') as Birtyday" +
                            ",hh.id as hid,hh.sid,hh.emp_no,hh.ProjectNo,DATE_FORMAT(hh.arraydate, '%Y/%m/%d') as arraydate,DATE_FORMAT(hh.ddate, '%Y/%m/%d') as ddate,hh.sceneNo,hh.dept_name" +
                            ",c.`No`,c.`Name` as cName" +
                            ",pj.`Name` as pjName"
                            + getItem() + getSuggestString(reportSelect) +
                            ",MAX(CASE WHEN isunusual='Y' then '異常' ELSE '正常' END) as item_isunusual" +
                            " from patient p" +
                            " INNER JOIN history_head hh on p.pid = hh.pid and p.id = hh.uid" +
                            " INNER JOIN customer c on hh.CustomerNo = c.`No`" +
                            " INNER JOIN history_body hb on hh.id = hb.head_id" +
                            " INNER JOIN project pj on  hh.ProjectNo = pj.`No`" +
                            " INNER JOIN history_head_add hha  on hh.id=hha.head_id" +
                            " INNER JOIN projectclass prjc on hha.ProjectClassNo = prjc.no" +
                            " INNER JOIN item i on hb.item_no = i.`no`";
            return sql;
        }
        //1
        public string getSqlString1(string filter, string CheckPrintKind, string data, string getFS, string getSS, string deptName, string reportSelect)
        {
            string sql = "";

            //篩選排序
            switch (filter)
            {
                case "date":
                    sql = getSqlString(reportSelect) +
                            " WHERE prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getFS.Remove(0, 3) + "" +
                            " GROUP BY  hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;
                case "sceneNo":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and hh.sceneNo = '" + data + "' " + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;
                case "name":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and p.`Name` = '" + data + "' " + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;
                case "id":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and p.Pid = '" + data + "' " + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;
                case "projectname":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and pj.`Name` = '" + data + "' " + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;

                case "projectname2":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and pj.`Name` = '" + data + "' " + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;

                case "company":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and c.`Name` = '" + data + "' " + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;

                case "deptName":
                    sql = getSqlString(reportSelect) +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and pj.`Name` = '" + data + "' and hh.dept_name = '" + deptName + "'" + getFS + "" +
                            " GROUP BY hb.head_id" +
                            " ORDER BY " + getSS + ";";
                    return sql;

                default:
                    break;

            }



            return sql;
        }


        private string getSuggestString(string reportSelect)
        {
            string whenSuggest = "(" +
                                "item_iid='2110101' OR " +
                                "item_iid='2110102' OR " +
                                "item_iid='2110106' OR " +
                                "item_iid='2110107' OR " +
                                "item_iid='2110111' OR " +
                                "item_iid='2110110' OR " +
                                "item_iid='2110402' OR " +
                                "item_iid='2110201' OR " +
                                "item_iid='2110202' OR " +
                                "item_iid='2110203' OR " +
                                "item_iid='2110204' OR " +
                                "item_iid='2110209' OR " +
                                "item_iid='M0060001' OR " +
                                "item_iid='M0060003' OR " +
                                "item_iid='M0070001' OR " +
                                "item_iid='M0100002' OR " +
                                "item_iid='M0120001' OR " +
                                "item_iid='3020502' OR " +
                                "item_iid='M0130001' OR " +
                                "item_iid='M0130002' OR " +
                                "item_iid='M0130003' OR " +
                                "item_iid='M0230002' OR " +
                                "item_iid='M0230003' OR " +
                                "item_iid='902' OR " +
                                "item_iid='2110316' OR " +
                                "item_iid='2110317' OR " +
                                "item_iid='R000142' OR " +
                                "item_iid='R000143' OR " +
                                "item_iid='R000144' OR " +
                                "item_iid='R000145' OR " +
                                "item_iid='R000146' OR " +
                                "item_iid='R000147' OR " +
                                "item_iid='R000583' OR " +
                                "item_iid='M0100003' OR " +
                                "item_iid='R000148' " +
                                ")";
            if (reportSelect == "從事供膳作業員工健康檢查紀錄表")
            {
                whenSuggest = "(i.category_tc = '肝炎標記' or i.category_tc = '傳染病檢驗' or i.category_tc = '糞便檢驗')";
            }
            if (reportSelect == "從事供膳作業員工健康檢查紀錄表")
                whenSuggest = ",GROUP_CONCAT(case WHEN hb.suggest<>'' and " + whenSuggest + " THEN CONCAT(hb.suggest,\":\",isunusual,\",\") ELSE '' END  SEPARATOR '') as suggest";
            else
                whenSuggest = ",GROUP_CONCAT(case WHEN hb.suggest<>'' and " + whenSuggest + " THEN CONCAT('*',hb.suggest) ELSE '' END  SEPARATOR '') as suggest";
            return whenSuggest;
        }
        private string getItem()
        {
            string itemText =
                            //作業經歷
                            ",max(case when item_iid = 'R000001' then DATE_FORMAT(qty, '%Y/%m/%d') else null end) as 'R000001'" +
                            ",max(case when item_iid = 'R000002' then qty else null end) as 'R000002'" +
                            ",max(case when item_iid = 'R000003' then qty else null end) as 'R000003'" +
                            ",max(case when item_iid = 'R000004' then qty else null end) as 'R000004'" +
                            ",max(case when item_iid = 'R000005' then qty else null end) as 'R000005'" +
                            ",max(case when item_iid = 'R000006' then qty else null end) as 'R000006'" +
                            ",max(case when item_iid = 'R000007' then qty else null end) as 'R000007'" +
                            ",max(case when item_iid = 'R000008' then qty else null end) as 'R000008'" +
                            ",max(case when item_iid = 'R000065' then qty else null end) as 'R000065'" +
                            ",max(case when item_iid = 'R000066' then qty else null end) as 'R000066'" +
                            ",max(case when item_iid = 'R000067' then qty else null end) as 'R000067'" +
                            ",max(case when item_iid = 'R000068' then qty else null end) as 'R000068'" +
                            ",max(case when item_iid = 'R000069' then qty else null end) as 'R000069'" +
                            ",max(case when item_iid = 'R000070' then qty else null end) as 'R000070'" +
                            ",max(case when item_iid = 'R000071' then qty else null end) as 'R000071'" +
                            ",max(case when item_iid = 'R000036' then qty else null end) as 'R000036'" +
                            ",max(case when item_iid = 'R000038' then qty else null end) as 'R000038'" +
                            //
                            ",max(case when item_iid = 'R000953' then qty else null end) as 'R000953'" +
                            ",max(case when item_iid = 'R000954' then qty else null end) as 'R000954'" +
                            //既往病史
                            ",max(case when item_iid = 'R000040' then qty else null end) as 'R000040'" +
                            ",max(case when item_iid = 'R000041' then qty else null end) as 'R000041'" +
                            ",max(case when item_iid = 'R000042' then qty else null end) as 'R000042'" +
                            ",max(case when item_iid = 'R000043' then qty else null end) as 'R000043'" +
                            ",max(case when item_iid = 'R000043T01' then qty else null end) as 'R000043T01'" +
                            ",max(case when item_iid = 'R000044' then qty else null end) as 'R000044'" +
                            ",max(case when item_iid = 'R000045' then qty else null end) as 'R000045'" +
                            ",max(case when item_iid = 'R000046' then qty else null end) as 'R000046'" +
                            ",max(case when item_iid = 'R000047' then qty else null end) as 'R000047'" +
                            ",max(case when item_iid = 'R000048' then qty else null end) as 'R000048'" +
                            ",max(case when item_iid = 'R000049' then qty else null end) as 'R000049'" +
                            ",max(case when item_iid = 'R000050' then qty else null end) as 'R000050'" +
                            ",max(case when item_iid = 'R000051' then qty else null end) as 'R000051'" +
                            ",max(case when item_iid = 'R000052' then qty else null end) as 'R000052'" +
                            ",max(case when item_iid = 'R000053' then qty else null end) as 'R000053'" +
                            ",max(case when item_iid = 'R000054' then qty else null end) as 'R000054'" +
                            ",max(case when item_iid = 'R000055' then qty else null end) as 'R000055'" +
                            ",max(case when item_iid = 'R000056' then qty else null end) as 'R000056'" +
                            ",max(case when item_iid = 'R000057' then qty else null end) as 'R000057'" +
                            ",max(case when item_iid = 'R000058' then qty else null end) as 'R000058'" +
                            ",max(case when item_iid = 'R000058T01' then qty else null end) as 'R000058T01'" +
                            ",max(case when item_iid = 'R000060' then qty else null end) as 'R000060'" +
                            ",max(case when item_iid = 'R000060T01' then qty else null end) as 'R000060T01'" +
                            ",max(case when item_iid = 'R000061' then qty else null end) as 'R000061'" +
                            ",max(case when item_iid = 'R000061T01' then qty else null end) as 'R000061T01'" +
                            ",max(case when item_iid = 'R000062' then qty else null end) as 'R000062'" +
                            //生活習慣
                            ",max(case when item_iid = 'R000072' then qty else null end) as 'R000072'" +
                            ",max(case when item_iid = 'R000073' then qty else null end) as 'R000073'" +
                            ",max(case when item_iid = 'R000074' then qty else null end) as 'R000074'" +
                            ",max(case when item_iid = 'R000074T01' then qty else null end) as 'R000074T01'" +
                            ",max(case when item_iid = 'R000074T02' then qty else null end) as 'R000074T02'" +
                            ",max(case when item_iid = 'R000075' then qty else null end) as 'R000075'" +
                            ",max(case when item_iid = 'R000075T01' then qty else null end) as 'R000075T01'" +
                            ",max(case when item_iid = 'R000075T02' then qty else null end) as 'R000075T02'" +
                            ",max(case when item_iid = 'R000076' then qty else null end) as 'R000076'" +
                            ",max(case when item_iid = 'R000077' then qty else null end) as 'R000077'" +
                            ",max(case when item_iid = 'R000078' then qty else null end) as 'R000078'" +
                            ",max(case when item_iid = 'R000078T01' then qty else null end) as 'R000078T01'" +
                            ",max(case when item_iid = 'R000078T02' then qty else null end) as 'R000078T02'" +
                            ",max(case when item_iid = 'R000079' then qty else null end) as 'R000079'" +
                            ",max(case when item_iid = 'R000079T01' then qty else null end) as 'R000079T01'" +
                            ",max(case when item_iid = 'R000079T02' then qty else null end) as 'R000079T02'" +
                            ",max(case when item_iid = 'R000080' then qty else null end) as 'R000080'" +
                            ",max(case when item_iid = 'R000081' then qty else null end) as 'R000081'" +
                            ",max(case when item_iid = 'R000082' then qty else null end) as 'R000082'" +
                            ",max(case when item_iid = 'R000082T01' then qty else null end) as 'R000082T01'" +
                            ",max(case when item_iid = 'R000082T02' then qty else null end) as 'R000082T02'" +
                            ",max(case when item_iid = 'R000082T03' then qty else null end) as 'R000082T03'" +
                            ",max(case when item_iid = 'R000082T04' then qty else null end) as 'R000082T04'" +
                            ",max(case when item_iid = 'R000082T05' then qty else null end) as 'R000082T05'" +
                            ",max(case when item_iid = 'R000083' then qty else null end) as 'R000083'" +
                            ",max(case when item_iid = 'R000083T01' then qty else null end) as 'R000083T01'" +
                            ",max(case when item_iid = 'R000083T02' then qty else null end) as 'R000083T02'" +
                            ",max(case when item_iid = 'R000955' then qty else null end) as 'R000955'" +
                            //自覺症狀
                            ",max(case when item_iid = 'R000084' then qty else null end) as 'R000084'" +
                            ",max(case when item_iid = 'R000085' then qty else null end) as 'R000085'" +
                            ",max(case when item_iid = 'R000086' then qty else null end) as 'R000086'" +
                            ",max(case when item_iid = 'R000087' then qty else null end) as 'R000087'" +
                            ",max(case when item_iid = 'R000088' then qty else null end) as 'R000088'" +
                            ",max(case when item_iid = 'R000089' then qty else null end) as 'R000089'" +
                            ",max(case when item_iid = 'R000090' then qty else null end) as 'R000090'" +
                            ",max(case when item_iid = 'R000091' then qty else null end) as 'R000091'" +
                            ",max(case when item_iid = 'R000092' then qty else null end) as 'R000092'" +
                            ",max(case when item_iid = 'R000093' then qty else null end) as 'R000093'" +
                            ",max(case when item_iid = 'R000094' then qty else null end) as 'R000094'" +
                            ",max(case when item_iid = 'R000095' then qty else null end) as 'R000095'" +
                            ",max(case when item_iid = 'R000096' then qty else null end) as 'R000096'" +
                            ",max(case when item_iid = 'R000097' then qty else null end) as 'R000097'" +
                            ",max(case when item_iid = 'R000098' then qty else null end) as 'R000098'" +
                            ",max(case when item_iid = 'R000099' then qty else null end) as 'R000099'" +
                            ",max(case when item_iid = 'R000100' then qty else null end) as 'R000100'" +
                            ",max(case when item_iid = 'R000101' then qty else null end) as 'R000101'" +
                            ",max(case when item_iid = 'R000102' then qty else null end) as 'R000102'" +
                            ",max(case when item_iid = 'R000103' then qty else null end) as 'R000103'" +
                            ",max(case when item_iid = 'R000104' then qty else null end) as 'R000104'" +
                            ",max(case when item_iid = 'R000105' then qty else null end) as 'R000105'" +
                            ",max(case when item_iid = 'R000106' then qty else null end) as 'R000106'" +
                            ",max(case when item_iid = 'R000106T01' then qty else null end) as 'R000106T01'" +
                            ",max(case when item_iid = 'R000107' then qty else null end) as 'R000107'" +
                            ",max(case when item_iid = 'R000109' then qty else null end) as 'R000109'" +
                            ",max(case when item_iid = 'R000110' then qty else null end) as 'R000110'" +
                            ",max(case when item_iid = 'R000114' then qty else null end) as 'R000114'" +
                            ",max(case when item_iid = 'R000188' then qty else null end) as 'R000188'" +
                            ",max(case when item_iid = 'R000188T01' then qty else null end) as 'R000188T01'" +
                            //醫師問答
                            ",max(case when item_iid = 'R000909' then CONCAT(qty,\":\",isunusual) else null end) as 'R000909'" +
                            ",max(case when item_iid = 'R000919' then CONCAT(qty,\":\",isunusual) else null end) as 'R000919'" +
                            ",max(case when item_iid = 'R000920' then CONCAT(qty,\":\",isunusual) else null end) as 'R000920'" +
                            ",max(case when item_iid = 'R000921' then CONCAT(qty,\":\",isunusual) else null end) as 'R000921'" +
                            //是否合格
                            ",max(case when item_iid = 'R000887' then CONCAT(qty,\":\",isunusual) else null end) as 'R000887'" +
                            //檢查醫師 總評醫師
                            ",max(case when item_iid = 'R000938' then qty  else null end) as 'R000938'" +
                            ",max(case when item_iid = 'R000910' then qty  else null end) as 'R000910'" +
                            //檢查項目
                            ",max(case when item_iid = '2110101' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110101'" +
                            ",max(case when item_iid = '2110102' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110102'" +
                            ",max(case when item_iid = '2110106' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110106'" +
                            ",max(case when item_iid = '2110107' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110107'" +
                            ",max(case when item_iid = '2110111' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110111'" +
                            ",max(case when item_iid = '2110110' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110110'" +
                            ",max(case when item_iid = '2110115' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110115'" +

                            //眼睛
                            ",max(case when item_iid = '2110201' then CONCAT(\"左眼裸視\",\":\",qty,\":\",isunusual) else 0 end) as 'A2110201'" +
                            ",max(case when item_iid = '2110202' then CONCAT(\"右眼裸視\",\":\",qty,\":\",isunusual) else 0 end) as 'A2110202'" +
                            ",max(case when item_iid = '2110203' then CONCAT(\"左眼矯正\",\":\",qty,\":\",isunusual) else 0 end) as 'A2110203'" +
                            ",max(case when item_iid = '2110204' then CONCAT(\"右眼矯正\",\":\",qty,\":\",isunusual) else 0 end) as 'A2110204'" +

                            ",max(case when item_iid = '2110209' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110209'" +
                            ",max(case when item_iid = '2110210' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110210'" +
                            ",max(case when item_iid = '2110211' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110211'" +
                            //耳朵
                            ",max(case when item_iid = '2110316' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110316'" +
                            ",max(case when item_iid = '2110317' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110317'" +

                            ",max(case when item_iid = 'M0050001' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0050001'" +
                            ",max(case when item_iid = '2110402' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A2110402'" +
                            //各系統或部位理學檢查
                            ",max(case when item_iid = 'R000583' then CONCAT(qty,\":\",isunusual) else null end) as 'R000583'" +
                            ",max(case when item_iid = 'R000143' then CONCAT(qty,\":\",isunusual) else null end) as 'R000143'" +
                            ",max(case when item_iid = 'R000144' then CONCAT(qty,\":\",isunusual) else null end) as 'R000144'" +
                            ",max(case when item_iid = 'R000145' then CONCAT(qty,\":\",isunusual) else null end) as 'R000145'" +
                            ",max(case when item_iid = 'R000146' then CONCAT(qty,\":\",isunusual) else null end) as 'R000146'" +
                            ",max(case when item_iid = 'R000147' then CONCAT(qty,\":\",isunusual) else null end) as 'R000147'" +
                            ",max(case when item_iid = 'R000142' then CONCAT(qty,\":\",isunusual) else null end) as 'R000142'" +
                            //胸部X光檢查
                            ",max(case when item_iid = '902' then CONCAT(qty,\":\",isunusual) else null end) as 'A902'" +
                            //尿液檢查(尿蛋白,尿潛血)
                            ",max(case when item_iid = 'M0230002' then CONCAT(qty,\":\",isunusual) else '' end) as 'M0230002'" +
                            ",max(case when item_iid = 'M0230003' then CONCAT(qty,\":\",isunusual) else '' end) as 'M0230003'" +
                            //血液檢查(白血球,紅血球,血色素)
                            ",max(case when item_iid = 'M0060001' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0060001'" +
                            ",max(case when item_iid = 'M0060002' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0060002'" +
                            ",MAX(CASE hb.item_iid WHEN 'M0060003' THEN CONCAT( if (ISNULL(hb.report_text)OR hb.report_text = '',(CASE WHEN hb.qty = '' THEN IF(hb.report_text <> '', hb.report_text, '缺Data') WHEN(hb.qty IS NULL) THEN '缺Data' ELSE hb.qty END ),hb.report_text),\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) ELSE NULL END ) as 'M0060003'" +
                            //SGPT 
                            ",max(case when item_iid = 'M0070001' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0070001'" +
                            //生化血液檢查
                            ",max(case when item_iid = 'M0120001' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0120001'" +
                            ",max(case when item_iid = 'M0100002' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0100002'" +
                            ",max(case when item_iid = 'M0130001' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0130001'" +
                            ",max(case when item_iid = 'M0130002' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0130002'" +
                            ",max(case when item_iid = 'M0130003' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0130003'" +
                            ",max(case when item_iid = '3020502' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'A3020502'" +
                            ",max(case when item_iid = 'M0100003' then CONCAT(qty,\":\",isunusual,\":\",hb.m_normal_d,\":\",hb.m_normal_u) else 0 end) as 'M0100003'";

            return itemText;
        }
    }
}
