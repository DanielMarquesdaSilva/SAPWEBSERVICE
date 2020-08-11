using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAPbobsCOM;

namespace SAP.Models
{
    public class B1Connection
    {
        public string CompanyDB { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        
        internal static Company create_CN(Company oCompany, B1Connection model)
        {
            try
            {
                oCompany.CompanyDB = model.CompanyDB;
                oCompany.Server = "danielms";
                oCompany.LicenseServer = "danielms:30000";
                oCompany.UserName = model.UserName;
                oCompany.Password = model.Password;
                oCompany.DbUserName = model.DbUserName;
                oCompany.DbPassword = model.DbPassword;
                oCompany.DbServerType = BoDataServerTypes.dst_MSSQL2016;
                oCompany.UseTrusted = false;
                oCompany.Connect();
                return oCompany;
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw;
            }
        }

        internal static string read_CN(Company oCompany)
        {
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string StatusConnection;

            try
            {
                string sql = string.Format("select    sess.session_id as session_id, " +
                                           "          convert(varchar,sess.login_time) as login_time, " +
                                           "          sess.login_name as login_name, " +
                                           "          sess.status as status, " +
                                           "          convert(varchar,sess.last_request_end_time) as last_request_end_time, " +
                                           "          base.name as name, " +
                                           "          convert(varchar,conn.connect_time) as connect_time, " +
                                           "          convert(varchar,conn.last_read) as last_read, " +
                                           "          convert(varchar,conn.last_write) as last_write " +
                                           "from      sys.dm_exec_sessions sess " +
                                           "          inner join sys.databases base on (sess.database_id = base.database_id) " +
                                           "          inner join sys.dm_exec_connections conn on (conn.session_id = sess.session_id) " +
                                           "          inner join (select    top 1 " +
                                           "                                usr.USERID as sapuser, " +
                                           "                                usr.U_NAME as sapusername, " +
                                           "                                loguser.ProcessID as sapprocessid, " +
                                           "                                loguser.SessionID as sapsessionid, " +
                                           "                                loguser.Source as sapsource " +
                                           "                      from      [dbo].[OUSR] usr " +
                                           "                                inner join[dbo].[USR7] gro on(usr.USERID = gro.UserId) " +
                                           "                                inner join[dbo].[USR5] loguser on(usr.user_code = loguser.UserCode) " +
                                           "                      where     loguser.UserCode = 'manager' " +
                                           "                      and       loguser.Source = 'SBO_DI_API' " +
                                           "                      and       loguser.Date <= getdate() + 2 " +
                                           "                      order by  loguser.Date desc, loguser.Time desc) sap on(sess.session_id = sap.sapsessionid) ");
                rs.DoQuery(sql);
                StatusConnection = "session_id: " + Convert.ToString(rs.Fields.Item("session_id").Value) + " " +
                                   "login_time: " + Convert.ToString(rs.Fields.Item("login_time").Value) + " " +
                                   "login_name: " + Convert.ToString(rs.Fields.Item("login_name").Value) + " " +
                                   "status: " + Convert.ToString(rs.Fields.Item("status").Value) + " " +
                                   "last_request_end_time: " + Convert.ToString(rs.Fields.Item("last_request_end_time").Value) + " " +
                                   "name: " + Convert.ToString(rs.Fields.Item("name").Value) + " " +
                                   "connect_time: " + Convert.ToString(rs.Fields.Item("connect_time").Value) + " " +
                                   "last_read: " + Convert.ToString(rs.Fields.Item("last_read").Value) + " " +
                                   "last_write: " + Convert.ToString(rs.Fields.Item("last_write").Value);
                return StatusConnection;
            }
            catch (Exception)
            {
                int errCode;
                string errMsg;
                oCompany.GetLastError(out errCode, out errMsg);
                throw;
            }
        }
    }
}