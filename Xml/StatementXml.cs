using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace tomxyz.csob.xml;

[XmlRoot("FINSTA03")]
public class StatementXml
{
    [XmlElement(ElementName = "SHORTNAME")]
    public string Name { get; set; }

    [XmlElement(ElementName = "S25_CISLO_UCTU")]
    public string Account { get; set; }

    [XmlElement(ElementName = "S60_DATUM")]
    public string DateFrom { get; set; }

    [XmlElement(ElementName = "S62_DATUM")]
    public string DateTo { get; set; }

    [XmlElement(ElementName = "S60_CASTKA")]
    public string StartAmount { get; set; }

    [XmlElement(ElementName = "SUMA_KREDIT")]
    public string Plus { get; set; }

    [XmlElement(ElementName = "SUMA_DEBIT")]
    public string Minus { get; set; }

    [XmlElement("FINSTA05")]
    public List<MovementXml> MovementsXml { get; set; }
}


