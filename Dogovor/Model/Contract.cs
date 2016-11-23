using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
namespace Dogovor.Model
{
    public enum Status { Резидент, Нерезидент};
    public enum Currency { Рубль, Доллар }
    public enum Payment {Фиксированное, По_ставкам}
    

    public class Contract
    {
        public string TemplatePath {get; set;}
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

        public void save()
        {
            XmlSerializer ser = new XmlSerializer(typeof(Contract));
            TextWriter writer = new StreamWriter("config.xml", false, Encoding.GetEncoding(1251));
            ser.Serialize(writer, this);
            writer.Close();
        }

        public static Contract read()
        {
            Contract s = new Contract();
            if (File.Exists("config.xml"))
            {
                XmlSerializer mySerializer = new XmlSerializer(typeof(Contract));
                FileStream myFileStream = new FileStream("config.xml", FileMode.Open);
                s = (Contract)mySerializer.Deserialize(myFileStream);
                myFileStream.Close();
            }
            return s;
        }
    }
}
