using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace AuthorizeTransactions
{
    class DAL
    {
        public static OleDbConnection conn = new OleDbConnection();
        public static bool UpdateTransactions(Dictionary<Transaction, TransactionResponse> transactionsResponsed)
        {

            try
            {
                if (OpenConnection())
                {
                    foreach (KeyValuePair<Transaction, TransactionResponse> tr in transactionsResponsed)
                    {
                        try
                        {
                            OleDbCommand cmd = default(OleDbCommand);

                            string transactionDescription = Program.BuildDescription(tr.Value);

                            string sql = "UPDATE [MAMI Cobraza] SET [MAMI Cobraza].Observation =  " +
                            " IIF(ISNULL([MAMI Cobraza].Observation),'',[MAMI Cobraza].Observation) & '\r\n" + transactionDescription.Replace("'", "") + "' ";
                            if (tr.Value.code == "1")
                            {
                                sql += ", [MAMI Cobraza].Paid = True ";
                                sql += ", [MAMI Cobraza].[Payment Date] = #" + DateTime.Now.ToString("d") + "#";

                            }
                            sql += ", [MAMI Cobraza].ChargeNow = False ";
                            sql += " WHERE ((([MAMI Cobraza].ID)=" + tr.Key.ID + "));";
                            cmd = new OleDbCommand(sql, conn);
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception es)
                        {

                        }
                    }

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception es)
            {
                return false;
            }

        }

        public static bool CloseConnection()
        {

            try
            {

                conn.Close();
                return true;

            }
            catch (Exception es)
            {
                return false;
            }
        }

        public static bool OpenConnection()
        {

            try
            {

                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Inetpub\\wwwroot\\mami-web.com\\lloydsweb\\v2\\lloydsdb\\db_lloyds.accdb;Persist Security Info=True";
                conn.Open();
                return true;

            }
            catch (Exception es)
            {
                return false;
            }

        }

        public static List<Transaction> GetTransactions(bool chargeNow)
        {
            if (DAL.OpenConnection())
            {
                List<Transaction> transactions = new List<Transaction>();


                string sql = "SELECT mmcbr.ID, mmcbr.[Client Name], mmcbr.Policy, mmcbr.Amount, mmcbr.[Name on Credit Card], " +
                "mmcbr.[Credit Card Number], mmcbr.[Credit Card Exp], mmcbr.[Credit Card Security Code], mmcbr.[Credit Card Address], " +
                "mmcbr.[Credit Card City], mmcbr.[Credit Card State], mmcbr.[Credit Card Zip], mmcbr.[Credit Card Country], " +
                "mmcbr.[Observation], mmcbr.[E-Mail], mmcbr.[Payment Nro], mmcbr.[Credit Card Comission], mmcbr.[Agent Comission], " +
                "mmcbr.[Interest Amount], mmcbr.[Agent], " +
                "(SELECT Max([MAMI Cobraza].[Payment Nro]) AS [MaxOfPayment Nro] FROM [MAMI Cobraza] " +
                " WHERE ((([MAMI Cobraza].[Id Financing])=mmcbr.[Id Financing]))) AS MaxPayment" +
                " FROM [MAMI Cobraza] AS mmcbr WHERE (";
                if (chargeNow)
                {
                    sql += "(mmcbr.[ChargeNow] = True) ";
                }
                else
                {
                    sql += "((mmcbr.[Due Date])=#" + DateTime.Now.ToString("MM/dd/yyyy") + "#) ";
                }
                sql += "AND ((mmcbr.Paid)=False) ";
                sql += "AND ((mmcbr.[Payment Type])='CC'));";
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn);
                adapter.Fill(ds);
                dt = ds.Tables[0];

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        Transaction tr = new Transaction();
                        tr.ID = int.Parse(row["ID"].ToString());
                        tr.ClientName = row["Client Name"].ToString();
                        tr.Policy = row["Policy"].ToString();
                        tr.Amount = double.Parse(row["Amount"].ToString());
                        tr.CCName = row["Name on Credit Card"].ToString();
                        tr.CCNumber = row["Credit Card Number"].ToString();
                        tr.CCExpDate = row["Credit Card Exp"].ToString();
                        tr.CCSecCode = row["Credit Card Security Code"].ToString();
                        tr.CCAddress = row["Credit Card Address"].ToString();
                        tr.CCCity = row["Credit Card City"].ToString();
                        tr.CCState = row["Credit Card State"].ToString();
                        tr.CCZip = row["Credit Card Zip"].ToString();
                        tr.CCCountry = row["Credit Card Country"].ToString();
                        tr.Observations = row["Observation"].ToString();
                        tr.Email = row["E-Mail"].ToString();
                        tr.PaymentNumber = row["Payment Nro"].ToString();
                        tr.TotalNumberPayments = row["MaxPayment"].ToString();
                        tr.CCCommission = string.IsNullOrEmpty(row["Credit Card Comission"].ToString()) ? 0 : double.Parse(row["Credit Card Comission"].ToString());
                        tr.AgentName = row["Agent"].ToString();
                        tr.AgentCommission = string.IsNullOrEmpty(row["Agent Comission"].ToString()) ? 0 : double.Parse(row["Agent Comission"].ToString());
                        transactions.Add(tr);
                    }
                    catch (Exception es) { }
                }

                DAL.CloseConnection();
                return transactions;
            }
            else
            {
                return null;
            }
        }
    }
}
