using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EMS_Small_Chb
{
    internal class Access
    {

        private OleDbConnection connection = new OleDbConnection();    // データベース接続用オブジェクト
        private OleDbDataAdapter dataAdapter = new OleDbDataAdapter(); // テーブル操作実行用オブジェクト
        private OleDbCommand command = new OleDbCommand();             // クエリ格納用オブジェクト

        private String DbFileName ="";

        public void DbPathSet(string FName) 
        {
            DbFileName = FName;
        }

        public void DbOpen() 
        {
            // データベースをオープン
            connection.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + DbFileName;
            connection.Open();
        }

        public void DbAllSelect(string sql,ref System.Data.DataTable DataTable) 
        {
            // 直前の取得結果をクリア
            DataTable.Clear();

            command.CommandText = sql;
            command.Connection = connection;

            dataAdapter.SelectCommand = command;

            // 実行
            dataAdapter.Fill(DataTable);

            foreach (DataRow dr in DataTable.Rows)
            {
                Console.WriteLine((String)dr.ItemArray[0]);
            }

        }
        public void DbClose() 
        {
            // データベースをクローズ
            connection.Close();
        }

    }
}
