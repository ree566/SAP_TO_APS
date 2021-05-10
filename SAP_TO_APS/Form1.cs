using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using System.Data.OleDb;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Threading;

namespace SAP_TO_APS
{
    public partial class Form1 : Form
    {
        string ST; string AT; int R; string WE; string str1;

        string MATNR = string.Empty;
        string SERNR = string.Empty;
        public OleDbConnection cn;
        public OleDbConnection cn1;
        public string sql; string sql2; string sql1; string sql3;

        public DataSet myDataSet = new DataSet();
        public DataSet myDataSet1 = new DataSet();
        public DataSet myDataSet2 = new DataSet();
        public string FR;
        public Form1()
        {


            IDestinationConfiguration ID = new SAP_FRC();

            RfcDestinationManager.UnregisterDestinationConfiguration(ID);
            RfcDestinationManager.RegisterDestinationConfiguration(ID);
            RfcDestination dest = RfcDestinationManager.GetDestination("SAPMES");
            //Delete_Date();

            select_prepare_Schedule(dest);
            //  Zemd003();

            //MAIL();

            this.Close();
            Environment.Exit(Environment.ExitCode);
        }



        public void Zemd005(RfcDestination dest, string AUFN, string MATN)
        {
            //if (AUFN != "PSJ2491RA")
            //{
            RfcRepository repository = dest.Repository;

            IRfcFunction rfc = repository.CreateFunction("ZGET_SAP_SODNWO_DATA_CK");

            rfc.SetValue("WONO", AUFN);
            rfc.SetValue("SONO", "");
            rfc.SetValue("PONO", "");
            rfc.SetValue("SPFLG", "");
            rfc.SetValue("SDATE", "");
            rfc.SetValue("EDATE", "");
            rfc.SetValue("PLANT", "");

            try
            {
                rfc.Invoke(dest);
            }
            catch (Exception e)
            {
                RfcSessionManager.EndContext(dest);
                dest = null;
                repository = null;
                System.Diagnostics.Debug.Print(e.ToString());
                return;
            }

            IRfcTable table = rfc.GetTable("ZWODETAIL");  //獲取相應的業務內表
            DataTable dt = new DataTable();  //新建表格
            //  DataRow dr = dt.NewRow();
            string sql1 = "";

            System.Diagnostics.Debug.Print(table.RowCount + "");

            for (int i = 0; i < table.RowCount; i++)
            //for (int i = 0; i < 1; i++)
            {

                table.CurrentIndex = i;  //當前內表的索引行

                string SHKZG = table.GetString("SHKZG"); string AUFNR = table.GetString("AUFNR");
                double BDMNG = table.GetDouble("BDMNG"); double ENMNG = table.GetDouble("ENMNG"); // Int32 modify to double by Apple.Chen 20190312
                string POSNR = table.GetString("POSNR"); string BWART = table.GetString("BWART"); // Add POSNR & BWART by Apple.Chen 20190912
                string MATNR = table.GetString("MATNR"); string LGORT = table.GetString("LGORT"); string MATKL = table.GetString("EKGRP");
                string SCHGT = table.GetString("SCHGT"); string DUMPS = table.GetString("DUMPS"); string ALPGR = table.GetString("ALPGR");
                string RSNUM = table.GetString("RSNUM"); // Add RSNUM by Apple.Chen 20210324
                string STORLOC_BIN = table.GetString("STORLOC_BIN"); string GSTRP = table.GetString("GSTRP");

                if (SCHGT != "X" & DUMPS != "X" & (BDMNG - ENMNG) != 0 & SHKZG == "H")
                {

                    sql1 = "Insert into 備料明細(訂單,需求溯源,物料,SLoc,生產儲位,狀態,需求日期,需求數量,領料數量,採購群組,負責人,工位,上傳日期,儲格,缺料數量,儲格分類,物管,UnitsInStock,POSNR,RSNUM,BWART" +
                            ") values ('" + AUFNR + "','" + MATN + "','" + MATNR.TrimStart('0') + "','" + LGORT + "','" + STORLOC_BIN + "','REL','" + GSTRP + "','" + BDMNG + "','" + ENMNG + "','" + MATKL + "'," +
                            "'','',Convert(varchar(100),GETDATE(), 120),'','0','','','','" + POSNR + "','" + RSNUM + "', '"+ BWART + "');";//add POSNR by Apple.Chen at 20210205 & Add RSNUM & BWART by Apple.Chen at 20210324



                    conn_open2();



                    OleDbCommand objCmd = new OleDbCommand(sql1, cn);

                    objCmd.CommandTimeout = 0;
                    //執行資料庫指令OleDbCommand 
                    objCmd.ExecuteNonQuery();
                    cn.Close();
                }



            }
            dest = null;
            repository = null;
            // }


        }



        public void Zemd003()
        {
            conn_open2();
            sql1 = "SELECT 訂單,COUNT(儲格分類) AS 應領筆數,儲格分類 FROM 備料明細 left join ZWODETAIL on 生產儲位=STORLOC_BIN WHERE Convert(varchar(100),上傳日期,23)=Convert(varchar(100),GETDATE(),23) and Convert(varchar(100),上傳日期, 108)>'15:59:00'" +
                "and (STORLOC_BIN IS null and 生產儲位 not like 'S%' and 生產儲位 not like 'T%' or (訂單 like '%Q' OR 訂單 like 'A%'or 訂單 like '%Z1' or 訂單 like '%ZZ') AND (生產儲位 NOT LIKE '%5F' AND 生產儲位 NOT LIKE'%6F' AND 生產儲位 NOT LIKE '%7F' and 生產儲位 not like 'Z1')) GROUP BY 訂單,儲格分類";
            //sql1 = "INSERT saveloc_State (訂單,儲格,應領筆數,儲格分類)" +
            //"SELECT 訂單,'',COUNT(儲格分類) AS 應領筆數,儲格分類 FROM 備料明細 WHERE (生產儲位 not in(" + STORLOC_BIN_select() + ") or (訂單 like '%Q%' or 訂單 like '%R%')) and Convert(varchar(100),上傳日期,23)=Convert(varchar(100),GETDATE(),23) GROUP BY 訂單,儲格分類";
            OleDbDataAdapter adp1 = new OleDbDataAdapter(sql1, cn);
            DataSet set1 = new DataSet();
            adp1.Fill(set1, "備料明細");
            for (int x = 0; x <= set1.Tables["備料明細"].Rows.Count - 1; x++)
            {
                if (Convert.ToString(set1.Tables["備料明細"].Rows[x].ItemArray[0]) != "")
                {
                    sql3 = sql3 + "Insert into saveloc_State(訂單,應領筆數,儲格分類,prepare_Stime,prepare_Ftime,儲格)values('" + set1.Tables["備料明細"].Rows[x].ItemArray[0].ToString() + "','" + set1.Tables["備料明細"].Rows[x].ItemArray[1].ToString() + "',N'" + set1.Tables["備料明細"].Rows[x].ItemArray[2].ToString() + "','','','');";
                }
            }
            sql3 = sql3 + "update prepare_Schedule set prepare_stateID=3 where prepareID in (select max(prepareID) from prepare_Schedule GROUP BY PO) and prepare_stateID=2;";
            // conn_open1();
            OleDbCommand objCmd = new OleDbCommand(sql3, cn);
            objCmd.CommandTimeout = 0;
            //執行資料庫指令OleDbCommand 
            objCmd.ExecuteNonQuery();
            cn.Close();
        }

        private void conn_open1()
        {
            try
            {
                string str = "Provider=sqloledb;Data Source=PC-P7H55MLE\\MFGSLQ;Initial Catalog=TWM3;User Id=TWM3;Password=TWM3;";
                //string str = "Provider=sqloledb;Data Source=M3-SERVER\\M3SERVER,1433;Initial Catalog=TWM3;User Id=TWM3;Password=TWM3;";
                cn = new OleDbConnection(str);
                cn.Open();

            }
            catch (Exception ex)
            {
                //擷取錯誤並顯示 
                MessageBox.Show("錯誤訊息: " + ex.ToString());
            }
            finally
            {
                //不管有沒有錯誤都會執行的,你可以在這作關閉資料庫Connection的動作
            }
        }
        private void conn_open2()
        {
            try
            {
                string str = "Provider=sqloledb;Data Source=M3-SERVER\\M3SERVER,1433;Initial Catalog=TWM3;User Id=TWM3;Password=TWM3;";
                cn = new OleDbConnection(str);

                cn.Open();

            }
            catch (Exception ex)
            {
                //擷取錯誤並顯示 
                MessageBox.Show("錯誤訊息: " + ex.ToString());
            }
            finally
            {
                //不管有沒有錯誤都會執行的,你可以在這作關閉資料庫Connection的動作
            }
        }
        private void conn_open3()
        {
            try
            {
                string str1 = "Provider=sqloledb;Data Source=M3-SERVER\\M3SERVER,1433;Initial Catalog=TWM3;User Id=TWM3;Password=TWM3;";
                cn1 = new OleDbConnection(str1);

                cn1.Open();

            }
            catch (Exception ex)
            {
                //擷取錯誤並顯示 
                MessageBox.Show("錯誤訊息: " + ex.ToString());
            }
            finally
            {
                //不管有沒有錯誤都會執行的,你可以在這作關閉資料庫Connection的動作
            }
        }
        private void select_prepare_Schedule(RfcDestination dest)
        {
            conn_open2();
            str1 = "SELECT PO,replace(Model_name,' ','') as Model_name FROM prepare_Schedule where prepareID in (select max(prepareID) from prepare_Schedule GROUP BY PO) and prepare_stateID=2";
            OleDbDataAdapter adp1 = new OleDbDataAdapter(str1, cn);
            DataSet set1 = new DataSet();
            adp1.Fill(set1, "prepare_Schedule");

            for (int x = 0; x <= set1.Tables["prepare_Schedule"].Rows.Count - 1; x++)
            {
                if (Convert.ToString(set1.Tables["prepare_Schedule"].Rows[x].ItemArray[0]) != "")
                {

                    System.Diagnostics.Debug.Print("Row " + x + " " + set1.Tables["prepare_Schedule"].Rows[x].ItemArray[0]); //Debug SAP error model
                    
                    Zemd005(dest, set1.Tables["prepare_Schedule"].Rows[x].ItemArray[0].ToString(), set1.Tables["prepare_Schedule"].Rows[x].ItemArray[1].ToString());
                    
                    System.Diagnostics.Debug.Print("---");
                    
                }
            }
            cn.Close();
            Zemd003();
        }
        private void MAIL()
        {
            DataTable dtsch;
            int flag = 0;


            string body = null; dtsch = getdatable1();

            if (dtsch.Rows.Count > 0)
            {

                body += "<table border=1><td colspan=" + dtsch.Columns.Count + " align=center >庫別異常待確認明細</td><tr >";

                body += "<tr bgcolor='#99CCFF'>";
                for (int i1 = 0; i1 <= dtsch.Columns.Count - 1; i1++)
                { body += "<td>" + dtsch.Columns[i1].ToString() + "</td>"; }
                for (int i2 = 0; i2 <= dtsch.Rows.Count - 1; i2++)
                {
                    if (flag == 1)
                    { body += "<tr bgcolor='#FFCC99'>"; flag = 0; }
                    else { body += "<tr>"; flag = 1; }
                    for (int i3 = 0; i3 <= dtsch.Columns.Count - 1; i3++)
                    { { body += "<td>" + dtsch.Rows[i2][i3].ToString() + "</td>"; } }
                }

                body += "</table><br><br>";
                mail_loop1(body);
            } cn.Close();
        }
        private void mail_loop1(string body)
        {
            MailMessage myMail = new MailMessage();
            myMail.From = new MailAddress("M3-SERVER@advantech.com.tw", "【 " + DateTime.Now.ToString("yyyy/MM/dd") + " 】" + "工單備料庫別異常!!");
            //發送者
            //myMail.From = New MailAddress("系統測試發送超過2日加班/出勤異常尚未維護!!") '發送者
            //myMail.To.Addwe);
            //收件者
            myMail.To.Add("Way.Chien@advantech.com.tw,Jane.Lien@advantech.com.tw,Phoebe.Lee@advantech.com.tw");  //收件者
            //myMail.To.Add("Way.Chien@advantech.com.tw");  //收件者
            myMail.SubjectEncoding = Encoding.UTF8;
            //主題編碼格式
            myMail.Subject = "【 " + DateTime.Now.ToString("yyyy/MM/dd") + " 】 工單備料庫別異常!!!!";
            //主題

            myMail.IsBodyHtml = true;
            //HTML語法(true:開啟false:關閉) 	
            myMail.BodyEncoding = Encoding.UTF8;
            //內文編碼格式
            myMail.Body = "Dear PMC" + "<br/>" + "<br/>" + "<br/>" + body;
            //內文
            //myMail.Attachments.Add(New System.Net.Mail.Attachment("C:\Files\FileA.txt"))  '附件

            SmtpClient mySmtp = new SmtpClient();
            //建立SMT連線	
            //mySmtp.Credentials = New System.Net.NetworkCredential("M3-SERVER@advantech.com.tw", "system")  '連線驗證
            //mySmtp.Port = 25
            //'587   SMTP Port 
            mySmtp.Host = "Relay.advantech.com.tw";
            //SMTP主機名 	
            //mySmtp.EnableSsl = False ' true:開啟false:關閉  SSL驗證
            mySmtp.Send(myMail);
            //發送	
            mySmtp = null;


        }
        private DataTable getdatable1()
        {
            conn_open2();
            str1 = "SELECT 訂單,需求溯源,物料,SLoc,狀態,需求數量,採購群組 FROM 備料明細 where SLoc not in ('0015','0008','0012') and Convert(varchar(100),上傳日期, 23)=Convert(varchar(100),GETDATE(), 23) and Convert(varchar(100),上傳日期, 108)>'15:59:00' and 物管='' ORDER BY 採購群組 desc";
            OleDbDataAdapter adp1 = new OleDbDataAdapter(str1, cn); DataSet set1 = new DataSet(); adp1.Fill(set1, "備料明細"); cn.Close();
            return set1.Tables[0];
        }
        public void Delete_Date()
        {
            ST = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyy-MM-dd 16:00:00");
            conn_open2();
            sql1 = "delete 備料明細 where 上傳日期>'" + ST + "'";


            // conn_open1();
            OleDbCommand objCmd = new OleDbCommand(sql1, cn);
            objCmd.CommandTimeout = 0;
            //執行資料庫指令OleDbCommand 
            objCmd.ExecuteNonQuery();
            cn.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
