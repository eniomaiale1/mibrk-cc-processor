using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Net;
using System.IO;
using AuthorizeNet;
using Microsoft.Office.Interop.Excel;

namespace AuthorizeTransactions
{
    class Program
    {
        //Finance Account
        //private const string authorizeLogin = "login";
        //private const string authorizeTranKey = "key";

        //Escrow Account
        private const string authorizeLogin = "login";
        private const string authorizeTranKey = "key";

        static void Main(string[] args)
        {
            //System.Diagnostics.Debugger.Break();
            bool bolChargeNow = false;
            bool bolReport = false;
            if (args.Length > 0) {
                if (args[0] == "/chargeNow")
                {
                    bolChargeNow = true;
                }
                else {
                    if (args[0] == "/report") {
                        bolReport = true;
                    }
                }
            }

            if (bolReport) {
                DoReport();
                return;
            }

            Dictionary<Transaction, TransactionResponse> transactionsResponsed = new Dictionary<Transaction, TransactionResponse>();

            List<Transaction> transactions = DAL.GetTransactions(bolChargeNow);
            foreach (Transaction transaction in transactions) {
                TransactionResponse transactionResponse = ProcessTransaction(transaction);
                transactionsResponsed.Add(transaction, transactionResponse);
            }

            DAL.UpdateTransactions(transactionsResponsed);
        }

        public static void DoReport() {
            
            
            DateTime from = DateTime.Parse(DateTime.Now.AddDays(-1).ToString("M/d/yyyy") + " 00:00:00");
            DateTime to = DateTime.Parse(DateTime.Now.AddDays(-1).ToString("M/d/yyyy") + " 23:59:59");
            
            ReportingGateway gate = new ReportingGateway(authorizeLogin, authorizeTranKey, ServiceMode.Live);
            List<AuthorizeNet.Transaction> transactions = gate.GetTransactionList(from, to);

            List<AuthorizeNet.Transaction> transactionsDetail = new List<AuthorizeNet.Transaction>();
            foreach (AuthorizeNet.Transaction item in transactions)
            {
                AuthorizeNet.Transaction tr = new AuthorizeNet.Transaction();
                tr = gate.GetTransactionDetails(item.TransactionID);
                tr.FirstName = item.FirstName;
                if(tr.Status == "settledSuccessfully")
                    transactionsDetail.Add(tr);
            }

            string fileName = "";
            if (ExcelReport.GenerateReport(transactionsDetail, out  fileName)) {
                Email.SendEmail(fileName);
            }

        }

        public static string GetLast(string source, int tail_length)
        {
            if (tail_length >= source.Length)
                return source;
            return source.Substring(source.Length - tail_length);
        }

        public static TransactionResponse ProcessTransaction(Transaction transaction) {

            TransactionResponse trr = null;

            try
            {
                trr = new TransactionResponse();
                string[] expDate = transaction.CCExpDate.Split('/');
                string expDateF = expDate[0] + Program.GetLast(expDate[1],2); 

                // posting to: https://secure.authorize.net/gateway/transact.dll
                //String post_url = "https://test.authorize.net/gateway/transact.dll";
                String post_url = "https://secure.authorize.net/gateway/transact.dll";

                double totalAmount = transaction.Amount + transaction.CCCommission;

                Dictionary<string, string> post_values = new Dictionary<string, string>();
                //the API Login ID and Transaction Key must be replaced with valid values
                post_values.Add("x_login", authorizeLogin);
                post_values.Add("x_tran_key", authorizeTranKey);
                post_values.Add("x_delim_data", "TRUE");
                post_values.Add("x_delim_char", "|");
                post_values.Add("x_relay_response", "FALSE");
                post_values.Add("x_type", "AUTH_CAPTURE");
                post_values.Add("x_method", "CC");
                post_values.Add("x_card_num", transaction.CCNumber);
                post_values.Add("x_exp_date", expDateF);
                post_values.Add("x_amount", totalAmount.ToString());
                post_values.Add("x_description", transaction.Policy + " / " + transaction.Amount.ToString() + " + " + transaction.CCCommission.ToString());
                post_values.Add("x_first_name", transaction.ClientName);
                post_values.Add("x_last_name", transaction.ClientName);
                post_values.Add("x_address", transaction.CCAddress);
                post_values.Add("x_state", transaction.CCState);
                post_values.Add("x_zip", transaction.CCZip);
                post_values.Add("x_country", transaction.CCCountry);
                post_values.Add("x_email", transaction.Email);

                // Additional fields can be added here as outlined in the AIM integration
                // guide at: http://developer.authorize.net

                // This section takes the input fields and converts them to the proper format
                // for an http post.  For example: "x_login=username&x_tran_key=a1B2c3D4"
                String post_string = "";

                foreach (KeyValuePair<string, string> post_value in post_values)
                {
                    post_string += post_value.Key + "=" + HttpUtility.UrlEncode(post_value.Value) + "&";
                }
                post_string = post_string.TrimEnd('&');

                int countItems = 0;
                if (transaction.Amount > 0) { countItems++; }
                if (transaction.CCCommission > 0) { countItems++; }
 
                string[] line_items = new string[countItems];

                if (transaction.Amount > 0) { line_items[0] = "PMT<|>Finance: Payment " + transaction.PaymentNumber + " of " + transaction.TotalNumberPayments + "<|>" + transaction.Policy + "<|>1<|>" + transaction.Amount.ToString() + "<|>N"; }
                if (transaction.CCCommission > 0) { line_items[1] = "CMM<|>Online Payment Convenience Fee<|><|>1<|>" + transaction.CCCommission.ToString() + "<|>N"; }                

                    //{
                    //"item1<|>golf balls<|><|>2<|>18.95<|>Y",
                    //"item2<|>golf bag<|>Wilson golf carry bag, red<|>1<|>39.99<|>Y",
                    //"item3<|>book<|>Golf for Dummies<|>1<|>21.99<|>Y"};
	
                foreach( string value in line_items )
                {
                    post_string += "&x_line_item=" + HttpUtility.UrlEncode(value);
                }
                

                // create an HttpWebRequest object to communicate with Authorize.net
                HttpWebRequest objRequest = (HttpWebRequest)WebRequest.Create(post_url);
                objRequest.Method = "POST";
                objRequest.ContentLength = post_string.Length;
                objRequest.ContentType = "application/x-www-form-urlencoded";

                // post data is sent as a stream
                StreamWriter myWriter = null;
                myWriter = new StreamWriter(objRequest.GetRequestStream());
                myWriter.Write(post_string);
                myWriter.Close();

                // returned values are returned as a stream, then read into a string
                String post_response;
                HttpWebResponse objResponse = (HttpWebResponse)objRequest.GetResponse();
                using (StreamReader responseStream = new StreamReader(objResponse.GetResponseStream()))
                {
                    post_response = responseStream.ReadToEnd();
                    responseStream.Close();
                }

                // the response string is broken into an array
                // The split character specified here must match the delimiting character specified above
                Array response_array = post_response.Split('|');

                // the results are output to the screen in the form of an html numbered list.
                //resultSpan.InnerHtml += "<OL> \n";
                int counter = 0;
                foreach (string value in response_array)
                {
                    counter++;

                    switch (counter)
                    {
                        case 1:
                            trr.code = value;
                            break;
                        case 2:
                            trr.subcode = value;
                            break;
                        case 3:
                            trr.reasonCode = value;
                            break;
                        case 4:
                            trr.reasinText = value;
                            break;
                        case 5:
                            trr.authorizationCode = value;
                            break;
                        case 6:
                            trr.avs = value;
                            break;
                        case 7:
                            trr.transactionID = value;
                            break;
                        case 8:
                            trr.invoiceNumber = value;
                            break;
                        case 9:
                            trr.description = value;
                            break;
                        case 10:
                            trr.amount = value;
                            break;
                        case 11:
                            trr.method = value;
                            break;
                        case 12:
                            trr.type = value;
                            break;
                        case 13:
                            trr.customerID = value;
                            break;
                        case 14:
                            trr.customerName = value;
                            break;
                        case 15:
                            trr.customerLastName = value;
                            break;
                        case 16:
                            trr.company = value;
                            break;
                        case 17:
                            trr.address = value;
                            break;
                        case 18:
                            trr.city = value;
                            break;
                        case 19:
                            trr.state = value;
                            break;
                        case 20:
                            trr.zip = value;
                            break;
                        case 21:
                            trr.country = value;
                            break;
                        case 22:
                            trr.phone = value;
                            break;
                        case 23:
                            trr.fax = value;
                            break;
                        case 24:
                            trr.email = value;
                            break;
                        case 38:
                            trr.md5Hash = value;
                            break;

                    }

                    //resultSpan.InnerHtml += "<LI>" + value + "&nbsp;</LI> \n";
                }
            }
            catch (Exception es) { }
            return trr;
            //resultSpan.InnerHtml += "</OL> \n";

        }

        public static string BuildDescription(TransactionResponse transactionResponse) {
            
            string description = "--------------------------------------------------\r\n";
            description += "Transaction : " + DateTime.Now.ToString("MM/dd/yyyy") + "\r\n";
            description += "--------------------------------------------------\r\n";
            description += "Code: " + transactionResponse.code +"\r\n";
            description += "subcode: " + transactionResponse.subcode +"\r\n";
            description += "reasonCode: " + transactionResponse.reasonCode +"\r\n";
            description += "reasinText: " + transactionResponse.reasinText +"\r\n";
            description += "authorizationCode: " + transactionResponse.authorizationCode +"\r\n";
            description += "avs: " + transactionResponse.avs +"\r\n";
            description += "transactionID: " + transactionResponse.transactionID +"\r\n";
            description += "invoiceNumber: " + transactionResponse.invoiceNumber +"\r\n";
            description += "description: " + transactionResponse.description +"\r\n";
            description += "amount: " + transactionResponse.amount +"\r\n";
            description += "method: " + transactionResponse.method +"\r\n";
            description += "type: " + transactionResponse.type +"\r\n";
            description += "customerID: " + transactionResponse.customerID +"\r\n";
            description += "customerName: " + transactionResponse.customerName +"\r\n";
            description += "customerLastName: " + transactionResponse.customerLastName +"\r\n";
            description += "company: " + transactionResponse.company +"\r\n";
            description += "address: " + transactionResponse.address +"\r\n";
            description += "city: " + transactionResponse.city +"\r\n";
            description += "state: " + transactionResponse.state +"\r\n";
            description += "zip: " + transactionResponse.zip +"\r\n";
            description += "country: " + transactionResponse.country +"\r\n";
            description += "phone: " + transactionResponse.phone +"\r\n";
            description += "fax: " + transactionResponse.fax +"\r\n";
            description += "email: " + transactionResponse.email +"\r\n";
            description += "md5Hash: " + transactionResponse.md5Hash +"\r\n";
            return description;
        }


    }
}
