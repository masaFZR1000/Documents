using System.Data;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EMS_Small_Chb.Form1;
using static EMS_Small_Chb.Json;

namespace EMS_Small_Chb
{
    internal class Oracle
    {
        static public OracleConnection oCon = null;
        static public OracleCommand oCmd = null;
        static public OracleTransaction oTrn = null;
        static public string strConnString = "";
        static public string sql = "";
        public static void db_FieldName(string conn_str)
        {
            DataTable sds = null;
            try
            {
                oCon = new OracleConnection();
                oCon.ConnectionString = conn_str;
                oCon.Open();

                var command = oCon.CreateCommand();
                command.CommandText = "SELECT * FROM VW_KENSSEIHMESIINFO";
                using (var reader = command.ExecuteReader())
                {
                    for (var i = 0; i < reader.FieldCount; i++)
                    {
                        var name = reader.GetName(i);
                        string sqlType = reader.GetDataTypeName(i);
                        Type clrType = reader.GetFieldType(i);

                        Console.WriteLine(name + "-" + sqlType + "-" + clrType);

                    }
                }
                oCon.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oCon != null) { oCon.Dispose(); oCon = null; }
            }
        }
        public static void db_get(string conn_str)
        {
            DataTable sds = null;
            sql = "";
            sql += "select * from VW_KENSSEIHINFO";
            sql += " where 0=0";
            sql += "  and  SEZONUM = '0015030580'";
            GetData(ref sds, conn_str, sql);
            if (sds != null)
            {
                if (sds.Rows.Count != 0)
                {
                    for (int y = 0; y < sds.Rows.Count; ++y)
                    {
                        string x0 = sds.Rows[y]["SEZONUM"].ToString();
                        string x1 = sds.Rows[y]["KATACD"].ToString();
                        string x2 = sds.Rows[y]["ODERNUM"].ToString();
                        string x3 = sds.Rows[y]["SYHASUU"].ToString();
                        string x4 = sds.Rows[y]["KEIYUSRNAM"].ToString();
                        string x5 = sds.Rows[y]["SIYOUSERNAM"].ToString();
                        string x6 = sds.Rows[y]["JTUKKHNNO"].ToString();
                        Console.WriteLine(x0 + "-" + x1 + "-" + x2 + "-" + x3 + "-" + x4 + "-" + x5 + "-" + x6);
                    }
                }
            }
            if (sds != null) { sds.Dispose(); sds = null; }
        }
        public static void option_get(string conn_str)
        {
            DataTable sds = null;
            sql = "";
            sql += "select * from VW_KENSSEIHMESIINFO";
            sql += " where 0=0";
            sql += "  and  SEZONUM = '0015031674'";
            GetData(ref sds, conn_str, sql);
            if (sds != null)
            {
                if (sds.Rows.Count != 0)
                {
                    for (int y = 0; y < sds.Rows.Count; ++y)
                    {
                        string x0 = sds.Rows[y]["SEZONUM"].ToString();
                        string x1 = sds.Rows[y]["OPTGYOUNUM"].ToString();
                        string x2 = sds.Rows[y]["OPTCOD"].ToString();
                        string x3 = sds.Rows[y]["OPTNAM"].ToString();
                        Console.WriteLine(x0 + "-" + x1 + "-" + x2 + "-" + x3);
                    }
                }
            }
            if (sds != null) { sds.Dispose(); sds = null; }
        }
        private static void GetData(ref DataTable datatable, string conn_str, string strSQL)
        {
            OracleCommand cmd = null;
            try
            {
                oCon = new OracleConnection();
                oCon.ConnectionString = conn_str;
                oCon.Open();
                cmd = new OracleCommand();
                cmd.Connection = oCon;
                cmd.CommandText = strSQL;

                OracleDataAdapter daAdapter = new OracleDataAdapter(cmd);
                DataSet dsDataSet = new DataSet();
                daAdapter.Fill(dsDataSet);
                datatable = new DataTable();
                datatable = dsDataSet.Tables[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                //MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (cmd != null) { cmd.Dispose(); cmd = null; }
                if (oCon != null) { oCon.Dispose(); oCon = null; }
            }
        }
        public static Header headerRead(string conn_str, Json.DATEBASE DatabaseInfo, string Key)
        {
            Header headerdata = new Header();
            String UNFLName = DatabaseInfo.MashineDb.UserName_FLName;
            String SNFLName = DatabaseInfo.MashineDb.SerialNo_FLName;
            String MLFLName = DatabaseInfo.MashineDb.Model_FLName;
            String ODFLName = DatabaseInfo.MashineDb.Order_FLName;
            String FRFLName = DatabaseInfo.MashineDb.Freq_FLName;

            String OCFLName = DatabaseInfo.MashineDb.OptionCode_FLName;
            String ONFLName = DatabaseInfo.MashineDb.OptionName_FLName;

            try
            {
                // 製品情報ゲット
                DataTable sds = null;
                sql = "";
                sql += "select * from " + DatabaseInfo.MashineDb.TableName;
                sql += " where 0=0";
                sql += "  and  " + DatabaseInfo.MashineDb.SerialNo_FLName + " = '" + Key + "'";
                GetData(ref sds, conn_str, sql);
                if (sds != null)
                {
                    if (sds.Rows.Count != 0)
                    {
                        for (int y = 0; y < sds.Rows.Count; ++y)
                        {
                            headerdata.UserName = sds.Rows[y][UNFLName].ToString();
                            headerdata.CreateNo = sds.Rows[y][SNFLName].ToString();
                            headerdata.PuroductModel = sds.Rows[y][MLFLName].ToString();
                            headerdata.PuroductOrder = sds.Rows[y][ODFLName].ToString();
                            headerdata.FreqSpec = int.Parse(sds.Rows[y][FRFLName].ToString());
                        }
                    }
                }
                if (sds != null) { sds.Dispose(); sds = null; }

                // オプション情報ゲット
                string ModelNo = "0000";
                sds = null;
                sql = "";
                sql += "select * from " + DatabaseInfo.MashineDb.OptionTableName;
                sql += " where 0=0";
                sql += "  and  " + DatabaseInfo.MashineDb.SerialNo_FLName + " = '" + Key + "'";
                GetData(ref sds, conn_str, sql);
                if (sds != null)
                {
                    if (sds.Rows.Count != 0)
                    {
                        ModelNo = "0000";
                        for (int y = 0; y < sds.Rows.Count; ++y)
                        {
                            Option Op = new Option();
                            Op.OptionCode = sds.Rows[y][OCFLName].ToString();
                            Op.OptionName = sds.Rows[y][ONFLName].ToString();
                            if (ModelNo == "0000") ModelNo = Op.OptionCode.Substring(0, 4);
                            headerdata.Option.Add(Op);
                        }
                    }
                }
                if (sds != null) { sds.Dispose(); sds = null; }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return headerdata;
            ;
        }
        public static bool db_Exist(string conn_str)
        {
            bool Ret = false;
            OracleCommand cmd = null;
            try
            {
                oCon = new OracleConnection();
                oCon.ConnectionString = conn_str;
                oCon.Open();
                Ret = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (cmd != null) { cmd.Dispose(); cmd = null; }
                if (oCon != null) { oCon.Dispose(); oCon = null; }
            }
            return Ret;
        }
    }
}
