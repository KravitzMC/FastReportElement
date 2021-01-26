using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

public class FastReportElement
{
    private string XMLFile;
    private XmlDocument xml = null;
    private List<string> allTextObjects = new List<string>();
    public static string ReportPath { get; set; }

    public FastReportElement(string xmlFile)
    {
        this.XMLFile = xmlFile;
        this.LoadXML();
    }

    public static void SetTextObject(string objectName, string value)
    {
        XmlDocument xml = new XmlDocument();
        xml.Load(ReportPath);
        XmlElement elm = xml.DocumentElement;

        XDocument xmlDoc = XDocument.Parse(xml.InnerXml);

        var items = from item in xmlDoc.Descendants("TextObject")
                    where item.Attribute("Name").Value == objectName
                    select item;

        foreach (XElement itemElement in items)
            itemElement.SetAttributeValue("Text", value);

        xmlDoc.Save(ReportPath);
    }

    public static void SetCellObject(string objectName, string value)
    {
        XmlDocument xml = new XmlDocument();
        xml.Load(ReportPath);
        XmlElement elm = xml.DocumentElement;

        XDocument xmlDoc = XDocument.Parse(xml.InnerXml);

        var items = from item in xmlDoc.Descendants("TableCell")
                    where item.Attribute("Name").Value == objectName
                    select item;

        foreach (XElement itemElement in items)
            itemElement.SetAttributeValue("Text", value);

        xmlDoc.Save(ReportPath);
    }

    public static void SetBarcodeValue(string objectName, string value)
    {
        XmlDocument xml = new XmlDocument();
        xml.Load(ReportPath);
        XmlElement elm = xml.DocumentElement;

        XDocument xmlDoc = XDocument.Parse(xml.InnerXml);

        var items = from item in xmlDoc.Descendants("BarcodeObject")
                    where item.Attribute("Name").Value == objectName
                    select item;

        foreach (XElement itemElement in items)
            itemElement.SetAttributeValue("Text", value);

        xmlDoc.Save(ReportPath);
    }


    public static void SetPictureObject(string objectName, string base64Image)
    {
        XmlDocument xml = new XmlDocument();
        xml.Load(ReportPath);
        XmlElement elm = xml.DocumentElement;
        XDocument xmlDoc = XDocument.Parse(xml.InnerXml);

        var items = from item in xmlDoc.Descendants("PictureObject")
                    where item.Attribute("Name").Value == objectName
                    select item;

        foreach (XElement itemElement in items)
            itemElement.SetAttributeValue("Image", base64Image);

        xmlDoc.Save(ReportPath);
    }

    public List<string> AllTextObjects
    {
        get
        {
            return allTextObjects;
        }
    }
    public void LoadXML()
    {
        allTextObjects.Clear();
        xml = new XmlDocument();
        xml.Load(XMLFile);

        XmlElement elm = xml.DocumentElement;
        XmlNodeList xnList = xml.SelectNodes("/Report/ReportPage/OverlayBand");
        XmlNodeList name = xml.GetElementsByTagName("TextObject");
        XmlNodeList objReportPage = xml.GetElementsByTagName("DataBand");

        foreach (XmlNode xl in name)
        {
            if (xl.Attributes["Text"] == null)
                AddAttribute(xl, "Text", string.Empty);

            var _text = xl.Attributes["Text"].InnerText;

            allTextObjects.Add(xl.Attributes["Name"].InnerText);
        }

        elm = null;
        xnList = null;
        name = null;
        objReportPage = null;

        GC.Collect();
        GC.WaitForPendingFinalizers();

    }
    protected void AddAttribute(XmlNode node, string attribute, string value)
    {
        if (node.NodeType == XmlNodeType.Element)
        {
            node.Attributes.Append(node.OwnerDocument.CreateAttribute(attribute));
            node.Attributes[attribute].Value = value;
            foreach (XmlNode subnode in node.ChildNodes)
            {
                AddAttribute(subnode, attribute, value);
            }
        }
    }
}
