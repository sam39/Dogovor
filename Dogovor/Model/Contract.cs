using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dogovor.Model
{
    public enum Status { Резидент, Нерезидент};
    public enum Currency { Рубль, Доллар }
    public enum Payment {Фиксированное, По_ставкам}

    public class Contract
    {
        public string Num {get;set;}
        public DateTime Date { get; set;}
        public Status  CustomerStatus { get; set;}
        public Decimal Price { get; set;}
        public Currency Currency { get; set;}
        public bool Report { get; set;}
        public bool Steps { get; set; }
        public bool Eroom { get; set; }
        public bool VozmRash { get; set; }
        public Payment Payment { get; set; }
        public Signatory Signatory { get; set; }
        public Branch Branch { get; set; }
    }
}
