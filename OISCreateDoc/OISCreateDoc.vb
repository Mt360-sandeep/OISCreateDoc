Imports System.Xml
Imports System.IO
Imports EverestAPI

Public Class OISCreateDoc

    Public OrderNo As String = ""
    Public vMsg As List(Of String)
    Public OrderCreated As Boolean = False
    Private sXMLReq As String = ""
    Private sResXML As String = ""

    Public Function CreateDoc(ByRef oisObj As OISModels.OISdata, ByRef evSess As MT360_EvSess.EverestSession) As Boolean
        Try
            Dim eSDoc As New EverestAPI.eoSalesDocument
            Dim iBillIndex As Integer = oisObj.oisCust.FindIndex(Function(addr) addr.ADDRESS_TYPE = "4") 'oisObj.oisCust.LongCount(Function(addr) addr.ADDRESS_TYPE = "4")
            Dim iShipIndex As Integer = oisObj.oisCust.FindIndex(Function(addr) addr.ADDRESS_TYPE = "5") 'oisObj.oisCust.LongCount(Function(addr) addr.ADDRESS_TYPE = "5")

            If oisObj.oisHeader.doc_type = "8" Then
                sXMLReq = getDocXML(oisObj, iBillIndex, iShipIndex)
                sResXML = eSDoc.Create(evSess.sUID, "8", sXMLReq, True)
            End If
            OrderCreated = True
            OrderNo = getXMLNodeVal(sResXML, "DocumentNo")
        Catch ex As Exception
            vMsg.Add(ex.Message)
            Return False
        End Try

        Return False
    End Function

    Public Function getXMLNodeVal(sXMLStr As String, sNode As String) As String
        Using reader As XmlReader = XmlReader.Create(New StringReader(sXMLStr))
            reader.ReadToFollowing(sNode)
            Return reader.ReadElementContentAsString()
        End Using
    End Function
    Private Function getDocXML(ByRef ordData As OISModels.OISdata, ByVal iBInd As Integer, ByVal iSInd As Integer)
        Dim sXMLDoc
        sXMLDoc = "<?xml version='1.0' encoding='utf-8' ?> "
        sXMLDoc = sXMLDoc + "<SalesDocument>"
        sXMLDoc = sXMLDoc + "<CustomerCode>" + ordData.oisCust(iBInd).CUST_CODE + "</CustomerCode>"
        If ordData.oisCust(iSInd).ADDR_CODE <> "" Then sXMLDoc = sXMLDoc + "<ShipToCode>" + ordData.oisCust(iSInd).ADDR_CODE + "</ShipToCode>"
        'sXMLDoc = sXMLDoc + "<SalesRepCode>" + sRep + "</SalesRepCode>"
        If ordData.oisHeader.pay_terms <> "" Then sXMLDoc = sXMLDoc + "<PayTermsCode>" + ordData.oisHeader.pay_terms + "</PayTermsCode>"
        If ordData.oisHeader.JURISDIC <> "" Then sXMLDoc = sXMLDoc + "<JurisdictionCode>" + ordData.oisHeader.JURISDIC + "</JurisdictionCode>"
        If ordData.oisHeader.DELIV_CODE <> "" Then sXMLDoc = sXMLDoc + "<ShipViaCode>" + ordData.oisHeader.DELIV_CODE + "</ShipViaCode>"
        'If sRef <> "" Then sXMLDoc = sXMLDoc + "<Reference>" + sRef + "</Reference>"
        'If sLoc <> "" Then sXMLDoc = sXMLDoc + "<LocationSubLocation>" + sLoc + "</LocationSubLocation>"
        If ordData.oisHeader.doc_alias <> "" Then sXMLDoc = sXMLDoc + "<DocumentAlias>" + ordData.oisHeader.doc_alias + "</DocumentAlias>"
        'sXMLDoc = sXMLDoc + "<CustomFields>"
        'If sCC10 <> "" Then sXMLDoc = sXMLDoc + "<CustomCharacterField Number='10'>" + sCC10 + "</CustomCharacterField>"
        'If sCC18 <> "" Then sXMLDoc = sXMLDoc + "<CustomCharacterField Number='18'>" + sCC18 + "</CustomCharacterField>"
        'If sCC19 <> "" Then sXMLDoc = sXMLDoc + "<CustomCharacterField Number='19'>" + sCC19 + "</CustomCharacterField>"
        'If sCC24 <> "" Then sXMLDoc = sXMLDoc + "<CustomCharacterField Number='24'>" + sCC24 + "</CustomCharacterField>"
        'If dCD1 <> "" Then sXMLDoc = sXMLDoc + "<CustomDateField Number='1'>" + dCD1 + "</CustomDateField>"
        'If dCD2 <> "" Then sXMLDoc = sXMLDoc + "<CustomDateField Number='2'>" + dCD2 + "</CustomDateField>"
        'If dCD3 <> "" Then sXMLDoc = sXMLDoc + "<CustomDateField Number='3'>" + dCD3 + "</CustomDateField>"
        'If dCD7 <> "" Then sXMLDoc = sXMLDoc + "<CustomDateField Number='7'>" + dCD7 + "</CustomDateField>"
        'sXMLDoc = sXMLDoc + "</CustomFields>"
        sXMLDoc = sXMLDoc + "<LineItems>"
        For Each objItms In ordData.oisDetails
            sXMLDoc = sXMLDoc + "<LineItem>"
            If objItms.ITEM_CODE <> "" Then
                sXMLDoc = sXMLDoc + "<ItemCode>" + objItms.ITEM_CODE + "</ItemCode>"
                sXMLDoc = sXMLDoc + "<Description>" + objItms.NOTE + "</Description>"
                sXMLDoc = sXMLDoc + "<Quantity>" + objItms.ITEM_QTY + "</Quantity>"
                sXMLDoc = sXMLDoc + "<ItemPrice>" + objItms.ITEM_PRICE + "</ItemPrice>"
                'If sItemMea <> "" Then sXMLDoc = sXMLDoc + "<UOMCode>" + sItemMea + "</UOMCode>"
                If objItms.ITEM_DISCOUNT <> "" Then sXMLDoc = sXMLDoc + "<DiscountValue>" + objItms.ITEM_DISCOUNT + "</DiscountValue>"
            Else
                sXMLDoc = sXMLDoc + "<Description>" + objItms.NOTE + "</Description>"
            End If
            sXMLDoc = sXMLDoc + "</LineItem>"
        Next
        sXMLDoc = sXMLDoc + "</LineItems>"
        sXMLDoc = sXMLDoc + "</SalesDocument>"
        getDocXML = sXMLDoc
    End Function
End Class
