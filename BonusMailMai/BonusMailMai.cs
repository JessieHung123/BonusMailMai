using Fpg.Workflow.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

namespace BonusMailMai
{
    public class BonusMailMai : Fpg.Workflow.Engine.ReflectionFactory.ScheduleApp
    {
        //public class BonusMail : Fpg.Workflow.Engine.ReflectionFactory.ScheduleApp
        //{
            private DbHelper gobjDbHelper = new DbHelper();
            private StringBuilder gSqlCommand = new StringBuilder();

            public override void Run(params string[] p_aInitParams)
            {
                RunSchedule();
            }

            private void RunSchedule()
            {
                DataTable oCheckMail = null;
                try
                {
                    string DTContentMonth = DateTime.Now.AddMonths(-2).ToString("yyyyMM");
                    gobjDbHelper.OpenDbResource("BonusODBCMai");
                    gSqlCommand.Remove(0, gSqlCommand.Length);
                    gSqlCommand.AppendFormat(string.Format("select YM,EMPNO,NM,BPT,GRPBNS,SCORE,EVABNS,QLAMT,INSVENVPTBNS,CLMAMT" +
                                                           ",OTHBNS,MSGETAMT,MUSXTDAMT,TOTAMT" +
                                                           ",RTRIM(XMLAGG(XMLELEMENT(e, RSCOMT || ' ' ) ORDER BY RSCOMT).EXTRACT('//text()'), ' ') AS RSCOMT " +
                                                           "from(select bonus.*, adbonus.rscomt from V0ADYF60 bonus " +
                                                           "left join V0ADEF31 adbonus " +
                                                           "on bonus.empno = adbonus.empid and bonus.ym = adbonus.ym) " +
                                                           "where YM = '{0}'" +
                                                           "GROUP BY  YM,EMPNO,NM,BPT,GRPBNS,SCORE,EVABNS,QLAMT,INSVENVPTBNS,CLMAMT,OTHBNS,MSGETAMT,MUSXTDAMT,TOTAMT  " +
                                                           "order by empno ", DTContentMonth));
                    oCheckMail = gobjDbHelper.Query(gSqlCommand.ToString(), "oCheckMail");
                    Fpg.Workflow.Data.LogCentral.CurrentLogger.LogInfo("======BonusMail_mai start count==== " + oCheckMail.Rows.Count.ToString());

                    if (oCheckMail.Rows.Count > 0)
                    {
                        MailService oMailService = new MailService();

                        for (int i = 0; i < oCheckMail.Rows.Count; i++)
                        {
                            string sQLAMT = string.IsNullOrEmpty(oCheckMail.Rows[i]["QLAMT"].ToString()) ? "&nbsp;" : oCheckMail.Rows[i]["QLAMT"].ToString();
                            string sINSVENVPTBNS = string.IsNullOrEmpty(oCheckMail.Rows[i]["INSVENVPTBNS"].ToString()) ? "&nbsp;" : oCheckMail.Rows[i]["INSVENVPTBNS"].ToString();
                            string sCLMAMT = string.IsNullOrEmpty(oCheckMail.Rows[i]["CLMAMT"].ToString()) ? "&nbsp;" : oCheckMail.Rows[i]["CLMAMT"].ToString();
                            string sOTHBNS = string.IsNullOrEmpty(oCheckMail.Rows[i]["OTHBNS"].ToString()) ? "&nbsp;" : oCheckMail.Rows[i]["OTHBNS"].ToString();
                            string sMUSXTDAMT = string.IsNullOrEmpty(oCheckMail.Rows[i]["MUSXTDAMT"].ToString()) ? "&nbsp;" : oCheckMail.Rows[i]["MUSXTDAMT"].ToString();
                            string sRSCOMT = string.IsNullOrEmpty(oCheckMail.Rows[i]["RSCOMT"].ToString()) ? "&nbsp;" : oCheckMail.Rows[i]["RSCOMT"].ToString();

                            string[] arrEMailTo = { oCheckMail.Rows[i]["EMPNO"].ToString() };     //多人傳送分號切割
                            string[] arrEMailCc = { };
                            string szSubject = oCheckMail.Rows[i]["YM"].ToString() + "效獎明細通知";
                            string szContext = "<table width='1200px'  border='1' style='text-align: center;' >" +
                                               "<tr><td rowspan='2'>人員編號</td><td rowspan='2'>姓名</td><td rowspan='2'>基點數</td><td colspan='2'>團體評核</td><td colspan='1'>主管評核</td><td rowspan='2'>作業環境</td><td rowspan='2'>工安環保</td><td rowspan='2'>連續虧損<br/>減發金額</td></tr><tr><td>金額</td><td>得分數</td><td>金額</td></tr>" +
                                               "<tr><td>" + oCheckMail.Rows[i]["EMPNO"].ToString() + "</td><td>" + oCheckMail.Rows[i]["NM"].ToString() + "</td><td> " + oCheckMail.Rows[i]["BPT"].ToString() + " </td><td>" + oCheckMail.Rows[i]["GRPBNS"].ToString() +
                                               " </td><td>" + oCheckMail.Rows[i]["SCORE"].ToString() + "</td><td>" + oCheckMail.Rows[i]["EVABNS"].ToString() + "</td><td>" + sQLAMT + "</td><td>" + sINSVENVPTBNS +
                                               "</td><td>" + sCLMAMT + " </td></tr></table><br/>" +
                                               "<table width = '1200px'  border = '1' style='text-align: center;' >" +
                                               "<tr><td>調整金額 </td><td>應發金額 </td><td>夏月獎金 </td><td>獎金合計 </td><td width = '600px'> 調整金額原因說明 </td></tr>" +
                                               "<tr><td>" + sOTHBNS + "</td><td>" + oCheckMail.Rows[i]["MSGETAMT"].ToString() + " </td><td>" + sMUSXTDAMT + " </td><td>" + oCheckMail.Rows[i]["TOTAMT"].ToString() +
                                               " </td><td>" + sRSCOMT + "</td></tr></table>";

                            oMailService.SendMailByIdNo(arrEMailTo, arrEMailCc, szSubject, szContext, MailService.BodyType.HTML, "BonusMail_" + (i + 1).ToString());
                        }
                    }
                    Fpg.Workflow.Data.LogCentral.CurrentLogger.LogInfo("======BonusMail_mai finish==== " + oCheckMail.Rows.Count.ToString());
                }
                catch (Exception ex)
                {
                    Fpg.Workflow.Data.LogCentral.CurrentLogger.LogError("======BonusMail_mai error==== ex.message:" + ex.Message + "\n ex : " + ex);
                }
            //}
        }
    }
}