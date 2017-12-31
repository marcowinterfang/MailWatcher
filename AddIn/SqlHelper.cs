using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddIn
{
    public static class SqlHelper
    {
        //private static readonly string constr = "server=.;uid=sa;pwd=123456Aa;database=Janssen";// ConfigurationManager.AppSettings["DBConnection"].ToString();
        //private static readonly string constr = "server=10.254.154.161;uid=sa;pwd=Accenture@2;database=Janssen";
        private static readonly string constr = "server=10.254.154.161;uid=sa;pwd=Accenture@2;database=Janssen_SIT";
        //private static readonly string constr = "server=10.225.64.217;uid=sa;pwd=sa!;database=Janssen_TEST";
        public static int SaveMailRecord(string entityId, string mailName, string attachmentName,
            string senderAdress, string mailBody, int originType, string subject, int caseFlag,
            int attatchmentCount,string createTime)
        {
            int origin = 0;
            decimal caseSt = 1.0M;
            int caseType = 0;
            int init_report_id = 0;
            string serialNumber = GetSerialNumber();
            string guide = GetGuidance(subject);
            string infoSource = ConvertDocType(originType);

            if (originType != 0)
            {
                caseSt = 1.1M;
            }
            string sql = @"INSERT INTO JN_CASE_MANAGEMENT([CASE_NO],[ENTITY_ID],[ORIGIN_ID],[MAIL_PATH]
                        ,[FILE_PATH],[MAIL_SENDER],[MAIL_BODY],[CASE_ST],[CASE_TYPE],[ORIGIN_DOC_TYPE],
                        [INIT_REPORT_ID],[LST_UPDATE_DT],[CREATED_DT],[MAIL_TITLE],[CASE_FLAG],[ATTACHMENT_COUNT],[CASE_GUIDANCE_TYPE])
                        VALUES(@Serial_Number,@Entity_Id,@Origin_Id,@Mail_Path,@File_Path,@Sender_Address,@Mail_Body,
                        @Case_St,@Case_Type,@Origin_Doc_Type,@Init_Report_Id,@Last_Update_Dt,
                        @CREATED_Dt,@Title,@Case_Flag,@attatchmentCount,@guide)";

            SqlParameter[] sp = new SqlParameter[] {
                new SqlParameter("@Serial_Number",serialNumber),
                new SqlParameter("@Entity_Id", entityId),
                new SqlParameter("@Origin_Id",origin),
                new SqlParameter("@Case_St",caseSt),
                new SqlParameter("@Case_Type",caseType),
                new SqlParameter("@Last_Update_Dt",DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                new SqlParameter("@CREATED_Dt",createTime),
                new SqlParameter("@Mail_Path",mailName),
                new SqlParameter("@File_Path",attachmentName),
                new SqlParameter("@Sender_Address",senderAdress),
                new SqlParameter("@Mail_Body",mailBody),
                new SqlParameter("@Origin_Doc_Type",originType),
                new SqlParameter("@Init_Report_Id",init_report_id),
                new SqlParameter("@Title",subject),
                new SqlParameter("@Case_Flag",caseFlag),
                new SqlParameter("@attatchmentCount",attatchmentCount),
                new SqlParameter("@guide",guide)
            };

            int result = ExecuteNoneQuery(sql, sp);
            SaveAccount(serialNumber, subject, infoSource);
            sql = "declare @emer int set @emer = " + caseFlag + " if @emer = 3 begin set @emer = 0 end else if @emer = 5 begin set @emer = 6 end ";
            sql += "update[dbo].[JN_CASE_MANAGEMENT] set CASE_FLAG = @emer where CASE_NO = '" + serialNumber + "' ";
            sql += " select top 1 id from JN_CASE_MANAGEMENT order by id desc ";
            LogHelper.Write(LogType.Debug, sql);
            DataTable dt = ExecuteDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                int.TryParse(dt.Rows[0][0].ToString(), out result);
                return result;
            }
            return 0;
        }

        #region private method
        private static string ConvertDocType(int typeId)
        {
            switch (typeId)
            {
                case 0:
                    return "Web連絡票";
                case 1:
                    return "";
                case 2:
                    return "";
                case 3:
                    return "";
                case 4:
                    return "";
                case 5:
                    return "";
                case 6:
                    return "";
                default:
                    break;
            }
            return "";
        }
        private static string GetGuidance(string subject)
        {
            string sql = "select * from JN_CASE_GUIDANCE";
            DataTable dt = ExecuteDataTable(sql);
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string[] keys = dr["Title"].ToString().Split('%');
                    bool compareResult = true;
                    if (keys.Length > 0)
                    {
                        foreach (var key in keys)
                        {
                            if (!subject.Contains(key))
                            {
                                compareResult = false;
                                continue;
                            }
                        }
                    }
                    if (compareResult)
                    {
                        return dr["Category"].ToString();
                    }
                }
            }
            return "";
        }
        #endregion
        public static int SaveAccount(string caseNo, string subject, string origin)
        {
            string sql = "Insert into JN_ACCOUNT(CASE_NO,RECEIPT_DT,ACCOUNT_NO,SUBJECT_NAME,GET_PATH) Values(@CaseNo,@Reciep,@CaseNo,@Subject,@Origin)";
            SqlParameter[] para = new SqlParameter[]
            {
                new SqlParameter("@CaseNo",caseNo),
                new SqlParameter("@Reciep",DateTime.Now.ToString("yyyy-MM-dd")),
                new SqlParameter("@Subject",subject),
                new SqlParameter("@Origin",origin)
            };
            int result = ExecuteNoneQuery(sql, para);
            return result;
        }
        public static string GetCaseGuidanceType(string subject)
        {
            try
            {
                string sql = "select * from JN_dictonary where dic_name = 'casetype'";
                DataTable dt = ExecuteDataTable(sql);
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow r in dt.Rows)
                    {
                        string[] keys = r["DIC_KEY"].ToString().Split('%');
                        bool flag = true;
                        foreach (var key in keys)
                        {
                            if (!subject.Contains(key))
                            {
                                flag = false;
                                break;
                            }
                        }
                        if (flag)
                        {
                            return r["DIC_VALUE"].ToString();
                        }
                    }
                }
                return "";
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static string GetSerialNumber()
        {
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand command = new SqlCommand("SerialNumber", con))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    var result = command.ExecuteScalar();
                    return result.ToString();
                }
            }
        }
        /// <summary>
        ///执行增删改
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static int ExecuteNoneQuery(string sql, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }
        /// <summary>
        /// 如果SqlParameter为空则转换为DBNull
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        static public object SqlNull(object obj)
        {
            if (obj == null)
            {
                return DBNull.Value;
            }
            return obj;
        }
        /// <summary>
        /// 如果是DBNULL则转换成null
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        static public object DBNullToNull(object obj)
        {
            if (obj == DBNull.Value)
            {
                return null;
            }
            return obj;
        }

        /// <summary>
        /// 返回datatable
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataTable(string sql, params SqlParameter[] pms)
        {
            DataTable dt = new DataTable();
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, constr))
            {
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);
                }
                adapter.Fill(dt);
            }
            return dt;
        }

        public static DataTable ExecuteDataTable(string sql)
        {
            DataTable dt = new DataTable();
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, constr))
            {
                adapter.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 返回单个值
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="pms"></param>
        /// <returns></returns>
        public static object ExecuteScalar(string sql, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }
                    con.Open();
                    return cmd.ExecuteScalar();
                }
            }
        }
    }
}
