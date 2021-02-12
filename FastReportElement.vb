'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FileName: FastReportElement.cs
'FileType: Visual C# Source file
'Author : Kirati Petkong
'Copy Rights : (c) 2021 Software & Scale Engineering Co.,Ltd
'Description : Fast Report Element object Redirect fix use for instead from problem offical solution.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml
Imports System.Xml.Linq

Public Class FastReportElement
    Private XMLFile As String
    Private xml As XmlDocument = Nothing
    Private allTextObjectsField As List(Of String) = New List(Of String)()
    Public Shared Property ReportPath As String

    Public Sub New(ByVal xmlFile As String)
        Me.XMLFile = xmlFile
        LoadXML()
    End Sub

    Public Shared Sub SetTextObject(ByVal objectName As String, ByVal value As String)
        Dim xml As XmlDocument = New XmlDocument()
        xml.Load(ReportPath)
        Dim elm As XmlElement = xml.DocumentElement
        Dim xmlDoc As XDocument = XDocument.Parse(xml.InnerXml)
        Dim items = From item In xmlDoc.Descendants("TextObject") Where Equals(item.Attribute("Name").Value, objectName) Select item

        For Each itemElement As XElement In items
            itemElement.SetAttributeValue("Text", value)
        Next

        xmlDoc.Save(ReportPath)
    End Sub

    Public Shared Sub SetCellObject(ByVal objectName As String, ByVal value As String)
        Dim xml As XmlDocument = New XmlDocument()
        xml.Load(ReportPath)
        Dim elm As XmlElement = xml.DocumentElement
        Dim xmlDoc As XDocument = XDocument.Parse(xml.InnerXml)
        Dim items = From item In xmlDoc.Descendants("TableCell") Where Equals(item.Attribute("Name").Value, objectName) Select item

        For Each itemElement As XElement In items
            itemElement.SetAttributeValue("Text", value)
        Next

        xmlDoc.Save(ReportPath)
    End Sub

    Public Shared Sub SetBarcodeValue(ByVal objectName As String, ByVal value As String)
        Dim xml As XmlDocument = New XmlDocument()
        xml.Load(ReportPath)
        Dim elm As XmlElement = xml.DocumentElement
        Dim xmlDoc As XDocument = XDocument.Parse(xml.InnerXml)
        Dim items = From item In xmlDoc.Descendants("BarcodeObject") Where Equals(item.Attribute("Name").Value, objectName) Select item

        For Each itemElement As XElement In items
            itemElement.SetAttributeValue("Text", value)
        Next

        xmlDoc.Save(ReportPath)
    End Sub

    Public Shared Sub SetPictureObject(ByVal objectName As String, ByVal base64Image As String)
        Dim xml As XmlDocument = New XmlDocument()
        xml.Load(ReportPath)
        Dim elm As XmlElement = xml.DocumentElement
        Dim xmlDoc As XDocument = XDocument.Parse(xml.InnerXml)
        Dim items = From item In xmlDoc.Descendants("PictureObject") Where Equals(item.Attribute("Name").Value, objectName) Select item

        For Each itemElement As XElement In items
            itemElement.SetAttributeValue("Image", base64Image)
        Next

        xmlDoc.Save(ReportPath)
    End Sub

    Public ReadOnly Property AllTextObjects As List(Of String)
        Get
            Return allTextObjectsField
        End Get
    End Property

    Public Sub LoadXML()
        allTextObjectsField.Clear()
        xml = New XmlDocument()
        xml.Load(XMLFile)
        Dim elm As XmlElement = xml.DocumentElement
        Dim xnList As XmlNodeList = xml.SelectNodes("/Report/ReportPage/OverlayBand")
        Dim name As XmlNodeList = xml.GetElementsByTagName("TextObject")
        Dim objReportPage As XmlNodeList = xml.GetElementsByTagName("DataBand")

        For Each xl As XmlNode In name
            If xl.Attributes("Text") Is Nothing Then AddAttribute(xl, "Text", String.Empty)
            Dim _text = xl.Attributes("Text").InnerText
            allTextObjectsField.Add(xl.Attributes("Name").InnerText)
        Next

        elm = Nothing
        xnList = Nothing
        name = Nothing
        objReportPage = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Protected Sub AddAttribute(ByVal node As XmlNode, ByVal attribute As String, ByVal value As String)
        If node.NodeType = XmlNodeType.Element Then
            node.Attributes.Append(node.OwnerDocument.CreateAttribute(attribute))
            node.Attributes(attribute).Value = value

            For Each subnode As XmlNode In node.ChildNodes
                AddAttribute(subnode, attribute, value)
            Next
        End If
    End Sub
End Class
