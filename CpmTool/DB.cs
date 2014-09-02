using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace CpmTool
{

    public struct TAccount
    {
        public int ID;
        public int sitetype;
        public string username;
        public string password;
        public bool important;
        public string sitename;
        public int volume;
        public int revenue;
        public string company;
    }

    public class DB
    {
        private String _mdbPath;
        private OleDbConnection _conn;

        static private DB _db = null;

        //单例
        private DB()
        {
            _mdbPath = Application.StartupPath + "\\Data\\Data.mdb";
            _conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _mdbPath);
            _conn.Open();
        }

        static public DB getInstence()
        {
            if (_db == null)
            {
                _db = new DB();
            }

            return _db;
        }


        public List<string> getCompanys(int sitetype)
        {
            List<string> comlist = new List<string>();
            OleDbCommand sql = _conn.CreateCommand();

            sql.CommandText = "SELECT company FROM Account WHERE sitetype=" + sitetype + " GROUP BY company";

            OleDbDataReader ret = sql.ExecuteReader();

            while (ret.Read())
            {
                comlist.Add(ret["company"].ToString());
            }
            ret.Close();

            return comlist;

        }

        public List<TAccount> getAccounts(int sitetype)
        {
            return this.getAccounts(sitetype, "");
        }

        public List<TAccount> getAccounts(int sitetype, string company)
        {
            List<TAccount> accountlist = new List<TAccount>();
            OleDbCommand sql = _conn.CreateCommand();

            string whereCompany = ( company == "" ? "" : ("AND company ='" + company + "'") );

            sql.CommandText = "SELECT * FROM Account WHERE sitetype=" + sitetype + whereCompany;

            OleDbDataReader ret = sql.ExecuteReader();

            while (ret.Read())
            {
                if ((bool)ret["enable"])
                {
                    TAccount account = new TAccount();
                    account.ID = (int)ret["ID"];
                    account.sitetype = (int)ret["sitetype"];
                    account.username = ret["username"].ToString();
                    account.password = ret["password"].ToString();
                    account.important = (bool)ret["important"];
                    account.sitename = ret["sitename"].ToString();
                    account.revenue = (int)ret["revenue"];
                    account.volume = (int)ret["volume"];
                    account.company = ret["company"].ToString();
                    accountlist.Add(account);

                }
                
            }
            ret.Close();

            return accountlist;
        }

        public TAccount getAccount(int id)
        {
            OleDbCommand sql = _conn.CreateCommand();

            sql.CommandText = "SELECT * FROM Account WHERE ID=" + id;

            OleDbDataReader ret = sql.ExecuteReader();

            TAccount account = new TAccount();
            if (ret.Read())
            {
                account.ID = id;
                account.sitetype = (int)ret["sitetype"];
                account.username = ret["username"].ToString();
                account.password = ret["password"].ToString();
                account.important = (bool)ret["important"];
                account.sitename = ret["sitename"].ToString();
                account.revenue = (int)ret["revenue"];
                account.volume = (int)ret["volume"];
                account.company = ret["company"].ToString();
            }
            ret.Close();

            return account;
        }
    }
}
