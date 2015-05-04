VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class should be general and reuseable i.e. it returns data rather than to edit cell values. The editing should be done from where the class is utilized.
' Query-strings named SQL should be dimmed in each querying function. Else it will set the query to Me.SQL, which should be done later.
' Private / Public should be used on class level (class members). Dim should be used on Sub/Function level for local variables.
' This class requires the reference Microsoft ActiveX Data Objects 6.1 Library for the ADODB object.

' Class members. Shared over the class.
Public SQL As String
Public Connection 'As ADODB.Connection
Public RecordSet 'As ADODB.RecordSet

Const DB_PATH = "\\esekina005\groupfbs\EM-SHARED INFO\qolsmat\DMC SmartApp\ECB.mdb" ' LAN path
'Const DB_PATH = "C:\Users\qolsmat\Desktop\ECB.mdb" ' Local path

' Custom "Constructor" - shared stuff for all queries and open DB
Public Sub Init()
    Set Me.Connection = New ADODB.Connection
    Me.Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_PATH & ";Persist Security Info=False"
    Me.Connection.Open
End Sub

' Shared stuff
Public Sub prepare(ByVal SQL As String)
    Me.SQL = SQL
    Set Me.RecordSet = New ADODB.RecordSet
    Me.RecordSet.Open Me.SQL, Me.Connection
End Sub

' Fetch one record as object CECB3
Public Function fetchOneECB3(ByVal SQL As String) As CECB3
    Dim r As CECB3
    Me.prepare SQL
    With Me.RecordSet
        Set r = New CECB3
        r.ID = .fields("ID")
        r.BU = .fields("BU")
        r.segment = .fields("segment")
        r.segmentDesc = .fields("segmentDesc")
        r.productArea = .fields("productArea")
        r.productAreaName = .fields("productAreaName")
        r.pCode = .fields("pCode")
        r.productLine = .fields("productLine")
        r.pCodeOne = .fields("pCodeOne")
        r.pCodeName = .fields("pCodeName")
        r.commodity = .fields("commodity")
        r.pcWarranty = .fields("pcWarranty")
        r.regionOrNot = .fields("regionOrNot")
        r.hmSalesOrNot = .fields("hmSalesOrNot")
        r.crossReference = .fields("crossReference")
        r.paHead = .fields("paHead")
        r.paController = .fields("paController")
        r.econApprover = .fields("econApprover")
        r.comments = .fields("comments")
    End With
    Set fetchOneECB3 = r
End Function

' If query is to be made for many records, this method returns a collection of CECB3 objects that can be looped through.
' Use if statement is "SELECT *" (fetch all columns)
Public Function fetchCollectionECB3(ByVal SQL As String) As Collection
    Dim c As Collection, r As Variant
    Set c = New Collection
    Me.prepare SQL
    With Me.RecordSet
        If .BOF And .EOF Then
            Debug.Print "No records"
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set r = New CECB3
            r.ID = .fields("ID")
            r.BU = .fields("BU")
            r.segment = .fields("segment")
            r.segmentDesc = .fields("segmentDesc")
            r.productArea = .fields("productArea")
            r.productAreaName = .fields("productAreaName")
            r.pCode = .fields("pCode")
            r.productLine = .fields("productLine")
            r.pCodeOne = .fields("pCodeOne")
            r.pCodeName = .fields("pCodeName")
            r.commodity = .fields("commodity")
            r.pcWarranty = .fields("pcWarranty")
            r.regionOrNot = .fields("regionOrNot")
            r.hmSalesOrNot = .fields("hmSalesOrNot")
            r.crossReference = .fields("crossReference")
            r.paHead = .fields("paHead")
            r.paController = .fields("paController")
            r.econApprover = .fields("econApprover")
            r.comments = .fields("comments")
            c.Add r, CStr(r.ID)
            .MoveNext
        Loop
    End With
    Set fetchCollectionECB3 = c
End Function

' Example method to toggle the pCode. If statement to determine if input P-code consists of 3 (P12) or 4 digits.
Public Function togglepCode(ByVal somePCode As String) As String
    Dim rs As CECB3
    If Len(somePCode) = 3 Then
        ' Creating statement manually here.
        SQL = "SELECT * FROM ECB3 WHERE pCodeOne = '" & somePCode & "'"
    Else
        SQL = "SELECT * FROM ECB3 WHERE pCode = '" & somePCode & "'"
    End If
    Set rs = Me.fetchOneECB3(SQL)
    If Len(somePCode) = 3 Then
        togglepCode = rs.pCode
    Else
        togglepCode = rs.pCodeOne
    End If
End Function

Public Function getpCodesPerBU(ByVal BU As String) As Collection
    Dim rs As Collection, SQL As String
    With Me
        SQL = .selectQuery("ECB3")
        SQL = SQL & .where("BU", "BMOD")
        Set rs = .fetchCollectionECB3(SQL)
    End With
    
    Set getpCodesPerBU = rs
End Function

' If fetching ECB4 with SELECT * (all columns)
Public Function fetchCollectionECB4(ByVal SQL As String) As Collection
    Dim c As Collection, r As CECB4
    Set c = New Collection
    Me.prepare SQL
    With Me.RecordSet
        If .BOF And .EOF Then
            Debug.Print "No records"
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set r = New CECB4
            r.ID = .fields("ID")
            r.CRGGlobalId = .fields("CRGGlobalId")
            r.CRGOneId = .fields("CRGOneId")
            r.CRGCBSId = .fields("CRGCBSId")
            r.CRGType = .fields("CRGType")
            r.CRGResp = .fields("CRGResp")
            r.regionCode = .fields("regionCode")
            r.marketCode = .fields("marketCode")
            r.countryCode = .fields("countryCode")
            r.GCUReportingGroupId = .fields("GCUReportingGroupId")
            r.GCUReportingGroupCode = .fields("GCUReportingGroupCode")
            r.GCUReportingGroupName = .fields("GCUReportingGroupName")
            r.parentType = .fields("parentType")
            r.CUReportingGroupId = .fields("CUReportingGroupId")
            r.CUReportingGroupName = .fields("CUReportingGroupName")
            r.CRGCode = .fields("CRGCode")
            r.CRGName = .fields("CRGName")
            r.CUManager = .fields("CUManager")
            r.comments = .fields("comments")
            r.directMarketController = .fields("directMarketController")
            c.Add r, CStr(r.ID)
            .MoveNext
        Loop
    End With
    Set fetchCollectionECB4 = c
End Function

' Could be useful if wanting to fetch several rows from the same column.
Public Function fetchCollection(ByVal SQL As String) As Collection
    Dim c As Collection, field As Variant
    Set c = New Collection
    Me.prepare SQL
    With Me.RecordSet
        If .BOF And .EOF Then
            Debug.Print "No records"
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            For Each field In .fields()
                If Not IsNull(field) Then
                    c.Add (field)
                End If
            Next field
            .MoveNext
        Loop
    End With
    Set fetchCollection = c
End Function

' Fetch one record as object CECB4
Public Function fetchOneECB4(ByVal SQL As String) As CECB4
    Dim r As CECB4
    Me.prepare SQL
    With Me.RecordSet
        Set r = New CECB4
            r.ID = .fields("ID")
            r.CRGGlobalId = .fields("CRGGlobalId")
            r.CRGOneId = .fields("CRGOneId")
            r.CRGCBSId = .fields("CRGCBSId")
            r.CRGType = .fields("CRGType")
            r.CRGResp = .fields("CRGResp")
            r.regionCode = .fields("regionCode")
            r.marketCode = .fields("marketCode")
            r.countryCode = .fields("countryCode")
            r.GCUReportingGroupId = .fields("GCUReportingGroupId")
            r.GCUReportingGroupCode = .fields("GCUReportingGroupCode")
            r.GCUReportingGroupName = .fields("GCUReportingGroupName")
            r.parentType = .fields("parentType")
            r.CUReportingGroupId = .fields("CUReportingGroupId")
            r.CUReportingGroupName = .fields("CUReportingGroupName")
            r.CRGCode = .fields("CRGCode")
            r.CRGName = .fields("CRGName")
            r.CUManager = .fields("CUManager")
            r.comments = .fields("comments")
            r.directMarketController = .fields("directMarketController")
    End With
    Set fetchOneECB4 = r
End Function

' Native Constructor
Private Sub Class_Initialize()
    Debug.Print "CDatabase class instantiated."
End Sub

' Native Destructor - Close connection to the DB. Essential so that DB won't get locked.
Private Sub Class_Terminate()
    Me.Connection.Close
    Debug.Print "CDatabase class terminated. Connection is closed."
End Sub

' Below functions are intended as a simple querybuilder.
Public Function selectQuery(ByVal table As String, Optional ByRef fields As Variant, Optional ByVal distinct As Boolean) As String
    Dim i As Integer, SQL As String
    SQL = "SELECT "
    If Not IsMissing(distinct) Then
        If distinct = True Then
            SQL = SQL & "DISTINCT "
        End If
    End If
    If Not IsMissing(fields) Then
        If IsArray(fields) Then
            For i = 0 To UBound(fields) - 1
                SQL = SQL & fields(i) & ", "
            Next i
            SQL = SQL & fields(UBound(fields)) & " "
        Else
            SQL = SQL & fields & " "
        End If
    Else
        SQL = SQL & "* "
    End If
    SQL = SQL & "FROM " & table
    selectQuery = SQL
End Function

' Table ECB3 should be redesigned to store NUMBERS for fields with numbers. Else this won't work.
' Or, store everything everywhere as TEXT and reimplement old solution with apostrophes.

' Function to create the initial WHERE partial of the query.
' Can prepend "!" to the value argument to negligate the WHERE clause. If you need more WHERE conditions, use the function .andWhere().
Public Function where(ByVal column As String, ByVal value As Variant) As String
    Dim SQL As String
    SQL = " WHERE "
    If StrComp(CStr(Mid(value, 1, 1)), "!", vbTextCompare) = 0 Then ' returns 0 if the strings match (lol?)
        value = Mid(value, 2)
        SQL = SQL & "NOT "
    End If
    ' If is not numeric, prepend and append "'" to the value, making it a string for SQL.
    If Not IsNumeric(value) Or column = "pCode" Or column = "pCodeOne" Then ' dirty solution with pCode, stored as text due to one record with alphabetic chars.
        value = "'" & value & "'"
    End If
    SQL = SQL & column & " = " & value
    where = SQL
End Function

' Function to create the AND WHERE part of the query. (if you already created the first "WHERE" part).
' Can prepend "!" to the value argument to negligate the WHERE clause.
Public Function andWhere(ByVal column As String, ByVal value As Variant) As String
    Dim SQL As String
    SQL = " AND "
    If StrComp(CStr(Mid(value, 1, 1)), "!", vbTextCompare) = 0 Then
        value = Mid(value, 2)
        SQL = SQL & "NOT "
    End If
    If Not IsNumeric(value) Or column = "pCode" Or column = "pCodeOne" Then
        value = "'" & value & "'"
    End If
    SQL = SQL & column & " = " & value
    andWhere = SQL
End Function

Public Function orWhere(ByVal column As String, ByVal value As String) As String
    Dim SQL As String
    SQL = " OR "
    If StrComp(CStr(Mid(value, 1, 1)), "!", vbTextCompare) = 0 Then
        value = Mid(value, 2)
        SQL = SQL & "NOT "
    End If
    If Not IsNumeric(value) Or column = "pCode" Or column = "pCodeOne" Then
        value = "'" & value & "'"
    End If
    SQL = SQL & column & " = " & value
    orWhere = SQL
End Function

Public Function getControllersByCustomer(ByVal customerName As String) As Collection
    Dim res As Collection, SQL As String
    SQL = ""
    With Me
        SQL = .selectQuery("ECB4", "directMarketController", True)
        SQL = SQL & .where("CRGResp", customerName)
    End With
    Set res = Me.fetchCollection(SQL)
    Set getControllersByCustomer = res
End Function
