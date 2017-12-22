using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace report2.SqlString
{
    class sqlString2
    {
        //sql 指令
        public string getSqlString2(string filter, string CheckPrintKind, string data, string getFS, string getSS, string deptName, string reportSelect, string getPKS)
        {
            string sql2 = "SELECT hh.id,c.`Name` as cName,hh.dept_name,hh.name,hh.pid" +
                        ",(SELECT name_tc from history_head_add hha WHERE stu = '0' and hh.id = hha.head_id)main_name_tc" +
                        ",(SELECT GROUP_CONCAT(hha.`code`) from history_head_add hha WHERE stu != '0' and hh.id = hha.head_id GROUP BY hha.head_id)speciel_add" +
                        ",(SELECT sex from patient p where p.pid = hh.pid)sex" +
                        ",(SELECT birtyday from patient p where p.pid = hh.pid)birtyday" +
                        ",hh.gtext,hh.emp_no,hh.sid,DATE_FORMAT(hh.arraydate,'%Y/%m/%d') arraydate,hh.dept_name,hb.item_iid,pj.Name as pName,pj.`No`,c.`no` as customerNo" +
                        ",hb.name_tc" +
                        ",qty,isunusual,hb.m_normal_d,hb.m_normal_u,i.unit,i.category,i.category_tc,i.is_pe,i.is_inst,i.roomname,hb.PayKind,c.`No`" +
                        " from history_head hh INNER JOIN (history_body hb, item i, customer c,project pj,history_head_add hha,projectclass prjc)" +
                        " on hh.id = hb.head_id and hb.item_no = i.no and hh.CustomerNo = c.`No` and hh.projectNo = pj.No and hh.id = hha.head_id and hha.ProjectClassNo = prjc.no";
            //" WHERE hb.PayKind = '00' and hh.sceneNo = '"+sceneNoText.Text+"'" +
            //" ORDER BY hh.sid,i.is_pe,i.is_inst,i.category";
            //DATE_FORMAT(p.Birtyday, '%Y/%m/%d') as Birtyday
            string orderBy = "";

            if(reportSelect == "健康檢查紀錄表A4")
            {
                orderBy = "hh.id,hb.PayKind,i.is_pe desc,i.is_inst,i.category,";
            }else
            {
                orderBy = "hh.id,i.category,hb.PayKind,i.is_pe desc,i.is_inst,";
            }

            //篩選排序\n
            switch (filter)
            {

                case "date":

                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " " + getFS +
                            " ORDER BY "+orderBy + getSS + ";";
                    return sql2;
                case "sceneNo":
                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and hh.sceneNo = '" + data + "' " + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;
                case "name":
                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and hh.name = '" + data + "' " + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;
                case "id":
                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and hh.Pid = '" + data + "' " + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;
                case "projectname":
                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and pj.`Name` = '" + data + "' " + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;

                case "projectname2":
                    sql2 = sql2 +
                             " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and pj.`Name` = '" + data.Split(':').GetValue(0) + "' " + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;

                case "company":
                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and c.`Name` = '" + data + "' " + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;

                case "deptName":
                    sql2 = sql2 +
                            " where prjc.CheckPrintKind like '%" + CheckPrintKind + "%' and" + getPKS + " and pj.`Name` = '" + data.Split(':').GetValue(0) + "' and hh.dept_name = '" + deptName + "'" + getFS + "" +
                            " ORDER BY " + orderBy + getSS + ";";
                    return sql2;

                default:
                    break;

            }

            return sql2;
        }


    }
}
