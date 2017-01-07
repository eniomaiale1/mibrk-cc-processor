using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AuthorizeTransactions
{
    class TransactionResponse
    {
        public string code{get;set;}
        public string subcode{get;set;}
        public string reasonCode{get;set;}
        public string reasinText{get;set;}
        public string authorizationCode{get;set;}
        public string avs{get;set;}
        public string transactionID{get;set;}
        public string invoiceNumber{get;set;}
        public string description{get;set;}
        public string amount{get;set;}
        public string method{get;set;}
        public string type{get;set;}
        public string customerID{get;set;}
        public string customerName{get;set;}
        public string customerLastName{get;set;}
        public string company{get;set;}
        public string address{get;set;}
        public string city{get;set;}
        public string state{get;set;}
        public string zip{get;set;}
        public string country{get;set;}
        public string phone{get;set;}
        public string fax{get;set;}
        public string email{get;set;}
        public string md5Hash{get;set;}

    }
}
