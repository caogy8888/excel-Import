using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LinqToExcel;
using Oracle.DataAccess.Client;

namespace ConsoleApplication1
{
    class Program
    {
        private static string connectionString = "User Id=hsyw;Password=hsyw;Data Source=125.210.208.56/rmss";
        static void Main(string[] args)
        {
            string path = @"C:\Users\caogy\Documents\Tencent Files\123842446\FileRecv\遗留单报表.xls";
            string path1 = @"C:\Users\caogy\Documents\Tencent Files\123842446\FileRecv\遗留单结单报表.xls";
            
            insertData(path, connectionString);
            insertData(path1, connectionString);
            Console.WriteLine("End");
            Console.ReadLine();
        }

        private static void insertData(string filePath, string connectionString) {
            var excel = new ExcelQueryFactory(filePath);
            var sheet = excel.Worksheet(0);
            var rows = from c in sheet
                       select c;

            foreach (var row in rows) {
                string reason = row["遗留原因"];
                string code = row["故障编号"];
                string guid = findeGuid(code);
                insert(guid, reason);
                
            }
        }

        private static void insert(string guid,string reason) {
            string id = Guid.NewGuid().ToString();
            
            using (OracleConnection connection = new OracleConnection(connectionString)) {
                string sql = "insert into \"T_FAU_ZB_YLYY\" values (:id, :ylyy, :t_fau_zb_zbguid)";
                OracleCommand command = new OracleCommand(sql, connection);
                command.Parameters.Add("id", id);
                command.Parameters.Add("ylyy", reason);
                command.Parameters.Add("t_fau_zb_zbguid", guid);
                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            }
            Console.WriteLine(String.Format("id:{0},guid:{1},reason:{2}", id, guid, reason));
        }

        private static string findeGuid(string code) {
            string guid = "";
            using (OracleConnection connection = new OracleConnection(connectionString)) {
                string sql = "select zbguid from t_fau_zb where gzbh=:gzbh";
                OracleCommand command = new OracleCommand(sql, connection);
                command.Parameters.Add(":gzbh", OracleDbType.Varchar2);
                command.Parameters[":gzbh"].Value = code;
                try
                {
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    while (reader.Read()) {
                        guid = reader[0].ToString();
                    }
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            
            };
            Console.WriteLine("zbguid=" + guid);
            return guid;
        }
    }
}
