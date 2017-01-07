using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AuthorizeTransactions
{
    class Transaction
    {
        public Int32 ID { get; set; }
        public string ClientName { get; set; }
        public string Policy { get; set; }
        public double Amount { get; set; }
        public string CCName { get; set; }
        public string CCNumber { get; set; }
        public string CCExpDate { get; set; }
        public string CCSecCode { get; set; }
        public string CCAddress { get; set; }
        public string CCCity { get; set; }
        public string CCState { get; set; }
        public string CCZip { get; set; }
        public string CCCountry { get; set; }
        public string Observations { get; set; }
        public string Email { get; set; }
        public string PaymentNumber { get; set; }
        public string TotalNumberPayments { get; set; }
        public double CCCommission { get; set; }
        public string AgentName { get; set; }
        public double AgentCommission { get; set; }
        public double InterestAmount { get; set; }
    }
}
