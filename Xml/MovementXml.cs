using System.Xml.Serialization;

namespace tomxyz.csob.xml;

public class MovementXml
{
    [XmlElement(ElementName = "S61_DATUM")]
    public string DateString { get; set; }

    [XmlElement(ElementName = "S61_CASTKA")]
    public string Amount { get; set; }

    [XmlElement(ElementName = "PART_ACCNO")]
    public string Account { get; set; }

    [XmlElement(ElementName = "S86_SPECSYMPAR")]
    public string SpecificSymbol { get; set; }

    [XmlElement(ElementName = "S86_VARSYMPAR")]
    public string VariableSymbol { get; set; }

    [XmlElement(ElementName = "PART_ID1_1")]
    public string Message1 { get; set; }

    [XmlElement(ElementName = "PART_ID1_2")]
    public string Message2 { get; set; }

    [XmlElement(ElementName = "PART_ID2_1")]
    public string Message3 { get; set; }

    [XmlElement(ElementName = "PART_ID2_2")]
    public string Message4 { get; set; }

    [XmlElement(ElementName = "PART_MSG_1")]
    public string Message5 { get; set; }

    [XmlElement(ElementName = "PART_MSG_2")]
    public string Message6 { get; set; }

    [XmlElement(ElementName = "REMARK")]
    public string Message7 { get; set; }
    
    [XmlElement(ElementName = "PART_ACC_ID")]
    public string Message8 { get; set; }
    
}
