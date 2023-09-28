VERSION 1.4 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Http"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 1.4

Private winHttpReq As Object
    
Private url As String
Private urn As String
Private host As String
Private protokol As String
Private payload As String
Private headers As Object
Public autoLogonPolicy As Boolean

Private boundary As String

Private Sub class_initialize()

    Set headers = CreateObject("Scripting.Dictionary")
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    autoLogonPolicy = False
    
End Sub
Public Sub init(baseUrl As String)
    
    If Right(baseUrl, 1) = "/" Then
        url = Left(baseUrl, Len(baseUrl) - 1)
    Else
        url = baseUrl
    End If
    
    temp = Split(url, "://")
    protokol = temp(0)
    host = temp(1)
    
    setBaseHeaders
    
End Sub

Public Sub setBaseHeaders()
    
    setHeader Array("Accept", "*/*")
    setHeader Array("Accept-Language", "en-US")
    setHeader Array("User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)")
    setHeader Array("Content-Type", "text/html;charset=utf-8")
    setHeader Array("Host", host)

End Sub

Public Sub setHeader(header As Variant)

    Select Case TypeName(header)
        Case "String"
            temp = Split(header, ":")
            headers(Trim(temp(0))) = Trim(temp(1))
        Case "Variant()"
            headers(header(0)) = header(1)
    End Select
            
End Sub

Public Sub removeHeader(header As String)

    headers.Remove header
    
End Sub

Public Sub removeAllHeaders()

    headers.RemoveAll
    
End Sub

Private Sub setHeaders()

    For Each akey In headers.Keys
        winHttpReq.SetRequestHeader akey, headers(akey)
    Next
    
End Sub

Public Function methodGet(Optional urn As String = "", Optional params As Collection = Nothing) As Object
    
    If Left(urn, 1) = "/" Then urn = Right(urn, Len(urn) - 1)
    fullUrl = url & "/" & urn
    
    If Not params Is Nothing Then
        Dim paramsStr As String
        paramsStr = "?"
        For Each aparam In params
            paramsStr = paramsStr & aparam(0) & "=" & aparam(1) & "&"
        Next
        paramsStr = Left(paramsStr, Len(paramsStr) - 1)
        fullUrl = fullUrl & paramsStr
    End If
    
    winHttpReq.Open "GET", fullUrl, False

    setHeaders
    If autoLogonPolicy = True Then winHttpReq.SetAutoLogonPolicy (0)
    winHttpReq.Send

    Set methodGet = winHttpReq

End Function

Public Function methodPost(Optional urn As String = "", Optional payload As Variant = "") As Object
    
    If Left(urn, 1) = "/" Then urn = Right(urn, Len(urn) - 1)
    fullUrl = url & "/" & urn
    
    payloadStr = ""
    Select Case TypeName(payload)
        Case "Collection"
            If Not payload Is Nothing Then
                Dim paramsStr As String
                paramsStr = ""
                For Each aparam In payload
                    paramsStr = paramsStr & aparam(0) & "=" & aparam(1) & "&"
                Next
                payloadStr = Left(paramsStr, Len(paramsStr) - 1)
                setHeader Array("Content-Type", "application/x-www-form-urlencoded")
            End If
        Case "String"
            payloadStr = payload
            setHeader Array("Content-Type", "application/json")
        Case "Byte()"
            payloadStr = payload
            setHeader Array("Content-Type", "multipart/form-data; boundary=" & boundary)
    End Select
       
    
    winHttpReq.Open "POST", fullUrl, False
    
    setHeaders
    If autoLogonPolicy = True Then winHttpReq.SetAutoLogonPolicy (0)
    winHttpReq.Send payloadStr
    
    Set methodPost = winHttpReq
    
    setHeader Array("Content-Type", "text/html;charset=utf-8")
    
End Function


Public Function createForm(Optional params As Collection, Optional files As Collection) As Variant

    Dim bytFormData As Variant
    Dim bytFormEnd As String
    Dim q As Long
    
    Dim ado As Object
    Dim rs As Object
    
    boundary = getBoundary(40)
    
    Dim strFormEnd As String

    strFormEnd = "--" & boundary & "--" & vbCrLf

    Dim formPartsParams As New Collection
    If Not params Is Nothing Then
        For Each aparam In params
            tempStr = "--" & boundary & vbCrLf
            tempStr = tempStr & "Content-Disposition: form-data; " & "name=""" & aparam(0) & """"
            tempStr = tempStr & vbCrLf & vbCrLf
            tempStr = tempStr & aparam(1)
            tempStr = tempStr & vbCrLf
            formPartsParams.Add tempStr
        Next
    End If
    
    Dim formPartsFiles As New Collection
    If Not files Is Nothing Then
    
        Set ado = CreateObject("ADODB.Stream")
        ado.Type = 1
        ado.Open
        
        For Each aparam In files
            ado.LoadFromFile aparam(1)
        
            tempStr = "--" & boundary & vbCrLf
            tempStr = tempStr & "Content-Disposition: form-data; "
            tempStr = tempStr & "name=""" & aparam(0) & """; "
            tempStr = tempStr & "filename=""" & Mid(aparam(1), InStrRev(aparam(1), "\") + 1) & """"
            tempStr = tempStr & vbCrLf
            tempStr = tempStr & "Content-Type: application/upload"
            tempStr = tempStr & vbCrLf & vbCrLf
            
            formPartsFiles.Add Array(tempStr, ado.Read, vbCrLf)
    
        Next
    End If
    

    Set rs = CreateObject("ADODB.Recordset")
    
    fullLen = 0
    For Each partStr In formPartsFiles
        fullLen = fullLen + Len(partStr(0)) + LenB(partStr(1)) + Len(partStr(2))
    Next
    fullLen = fullLen + Len(strFormEnd)

    
    rs.fields.Append "FormData", 205, fullLen
    rs.Open
    rs.AddNew

    
    Dim bytsFiles As New Collection
    Dim bytsParams As New Collection
    
    For Each partStr In formPartsParams
        rs("FormData") = ""
        bytParam = ""
        For q = 1 To Len(partStr)
            bytParam = bytParam & ChrB(Asc(Mid(partStr, q, 1)))
        Next
        rs("FormData").AppendChunk bytParam & ChrB(0)
        bytParam = rs("formData").GetChunk(Len(partStr))
        
        bytsParams.Add bytParam
    Next
    
    
    For Each partStr In formPartsFiles
        rs("FormData") = ""
        bytStartFile = ""
        For q = 1 To Len(partStr(0))
            bytStartFile = bytStartFile & ChrB(Asc(Mid(partStr(0), q, 1)))
        Next
        rs("FormData").AppendChunk bytStartFile & ChrB(0)
        bytStartFile = rs("formData").GetChunk(Len(partStr(0)))
        
        
        rs("FormData") = ""
        bytEndFile = ""
        For q = 1 To Len(partStr(2))
            bytEndFile = bytEndFile & ChrB(Asc(Mid(partStr(2), q, 1)))
        Next
        rs("FormData").AppendChunk bytEndFile
        bytEndFile = rs("formData").GetChunk(Len(partStr(2)))
        
        bytsFiles.Add Array(bytStartFile, partStr(1), bytEndFile)
    Next

    
    rs("FormData") = ""
    For q = 1 To Len(strFormEnd)
        bytFormEnd = bytFormEnd & ChrB(Asc(Mid(strFormEnd, q, 1)))
    Next
    rs("FormData").AppendChunk bytFormEnd
    bytFormEnd = rs("formData").GetChunk(Len(strFormEnd))
       
    
    rs("FormData") = ""
    For Each byt In bytsParams
        rs("FormData").AppendChunk byt
    Next
    For Each byt In bytsFiles
        rs("FormData").AppendChunk byt(0)
        rs("FormData").AppendChunk byt(1)
        rs("FormData").AppendChunk byt(2)
    Next
    rs("FormData").AppendChunk bytFormEnd
    bytFormData = rs("FormData")
    rs.Close

    createForm = bytFormData
    
End Function

Private Function getBoundary(lenStr As Integer) As String
    Dim possibleSymbols As String
    Dim result As String
    Dim i As Integer

    possibleSymbols = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    Randomize
    
    For i = 1 To lenStr
        result = result & Mid(possibleSymbols, Int((Len(possibleSymbols) * Rnd) + 1), 1)
    Next i

    getBoundary = result
End Function

