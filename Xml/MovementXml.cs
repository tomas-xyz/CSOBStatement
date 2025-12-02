using System.Xml.Serialization;

namespace tomxyz.csob.xml;

public class MovementXml
{
    [XmlElement(ElementName = "S61_DATUM")]
    public string DateString { get; set; } = string.Empty;

    [XmlElement(ElementName = "S61_CASTKA")]
    public string Amount { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_ACCNO")]
    public string Account { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_BANK_ID")]
    public string BankId { get; set; } = string.Empty;

    [XmlElement(ElementName = "S86_SPECSYMPAR")]
    public string SpecificSymbol { get; set; } = string.Empty;

    [XmlElement(ElementName = "S86_VARSYMPAR")]
    public string VariableSymbol { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_ACC_ID")]
    public string AccountId { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_ID1_1")]
    public string Message1 { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_ID1_2")]
    public string Message2 { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_ID2_1")]
    public string Message3 { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_ID2_2")]
    public string Message4 { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_MSG_1")]
    public string Message5 { get; set; } = string.Empty;

    [XmlElement(ElementName = "PART_MSG_2")]
    public string Message6 { get; set; } = string.Empty;

    [XmlElement(ElementName = "S61_POST_NAR")]
    public string Message7 { get; set; } = string.Empty;

    [XmlElement(ElementName = "REMARK")]
    public string Message8 { get; set; } = string.Empty;

}
