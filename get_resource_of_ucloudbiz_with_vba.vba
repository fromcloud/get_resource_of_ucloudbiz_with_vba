Option Explicit
Option Base 0

Const dictKey = 1
Const dictItem = 2

Const common = 1
Const storage = 2

Public capikey As String
Public csecretkey As String


Sub get_full_resource()
    Call request
End Sub


Sub init_account(flag As Integer)

    If flag = 1 Then
        capikey = Worksheets("user_config").Range("b2").Value
        csecretkey = Worksheets("user_config").Range("b3").Value
    Else
    End If

End Sub


Sub request()
    Dim url As String
    Dim request As Object
    Set request = CreateObject("Scripting.Dictionary")
    Call init_account(common)
    
    Call clear_sheets("vm")
    url = Worksheets("url_config").Range("b2").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listVirtualMachines")
    Worksheets("vm").Select
    Call fill_data_as_querytable("vm", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/virtualmachine/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))

    Call clear_sheets("disk")
    url = Worksheets("url_config").Range("b2").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listVolumes")
    Worksheets("disk").Select
    Call fill_data_as_querytable("disk", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/volume/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Call clear_sheets("public ip")
    url = Worksheets("url_config").Range("b2").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listPublicIpAddresses")
    Worksheets("public ip").Select
    Call fill_data_as_querytable("public ip", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/publicipaddress/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))

    Call clear_sheets("template")
    url = Worksheets("url_config").Range("b2").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listTemplates", "templatefilter=self")
    Worksheets("template").Select
    Call fill_data_as_querytable("template", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/template/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))

    Call clear_sheets("snapshot")
    url = Worksheets("url_config").Range("b2").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listSnapshots")
    Worksheets("snapshot").Select
    Call fill_data_as_querytable("snapshot", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/snapshot/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))


    Call clear_sheets("network")
    url = Worksheets("url_config").Range("b2").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listNetworks")
    Worksheets("network").Select
    Call fill_data_as_querytable("network", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/network/")
    Set request = Nothing
    Columns("AD:AJ").EntireColumn.Delete
    Call delete_row("network", "D")
    Application.Wait (Now + TimeValue("0:00:01"))


    Call clear_sheets("nas")
    url = Worksheets("url_config").Range("b3").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listVolumes", "status=online")
    Worksheets("nas").Select
    Call fill_data_as_querytable("nas", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/response/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))


    Call clear_sheets("waf")
    url = Worksheets("url_config").Range("b4").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listWAFs")
    Worksheets("waf").Select
    Call fill_data_as_querytable("waf", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/wafservice/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))


    Call clear_sheets("lb")
    url = Worksheets("url_config").Range("b5").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listLoadBalancers")
    Worksheets("lb").Select
    Call fill_data_as_querytable("lb", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/loadbalancer/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))

    
    Call clear_sheets("gslb")
    url = Worksheets("url_config").Range("b6").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listGslbServer")
    Worksheets("gslb").Select
    Call fill_data_as_querytable("gslb", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/lists/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))

    Call clear_sheets("cdn")
    url = Worksheets("url_config").Range("b7").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "statusCdnService")
    Worksheets("cdn").Select
    Call fill_data_as_querytable("cdn", "A1", get_sig_request(request, csecretkey, url))
    Call make_pretty("/lists/")
    Set request = Nothing
    Application.Wait (Now + TimeValue("0:00:01"))

    
    Call clear_sheets("db")
    url = Worksheets("url_config").Range("b8").Value
    Set request = CreateObject("Scripting.Dictionary")
    Call make_dic_for_request(request, "listInstances")
    Worksheets("db").Select
    Dim str As String
    str = get_sig_request(request, csecretkey, url)
    Call fill_data_as_xml("db", "A1", str)
    
    Set request = Nothing
    
    On Error GoTo errhdl
    ActiveWorkbook.Connections("api").Delete
    ActiveWorkbook.Connections("__").Deletete
    Exit Sub
errhdl:
    Resume Next
    
End Sub


Sub make_dic_for_request(ByRef request As Object, command As String, ParamArray Args() As Variant)
    Dim tarr() As String
    Dim i As Integer
    For i = 0 To UBound(Args)
        tarr = Split(Args(i), "=")
        request.Add tarr(0), tarr(1)
    Next
    
    request.Add "command", command
    request.Add "response", "xml"
    request.Add "apikey", capikey
End Sub


Sub fill_data_as_xml(ByVal target_sheet As String, ByVal target_range As String, ByVal final_url As String)
On Error GoTo errhdl
    Worksheets(target_sheet).Select
    ActiveWorkbook.XmlImport url:=final_url, ImportMap:=Nothing, Overwrite:=True, Destination:=Range(target_range)
    Exit Sub
errhdl:
    Call MsgBox("Eror" + target_sheet + " : " + "apikey or secretkey / apiurl is missing", vbOKOnly, "")
    Resume Next
End Sub


Sub fill_data_as_querytable(ByVal target_sheet As String, ByVal target_range As String, ByVal final_url As String)
    Call clear_sheets(target_sheet)
    Worksheets(target_sheet).Range(target_range).Select
    Debug.Print final_url
On Error GoTo errhdl
    Dim conn As Variant
    conn = "URL;" + final_url
    Dim qt As Object
    Set qt = ActiveSheet.QueryTables.Add(Connection:=conn, Destination:=Range(target_range))
    With qt
        .Name = "tempgeturl"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        '.WebSelectionType = xlEntirePage
        '.WebFormatting = xlWebFormattingRTF
        .WebFormatting = xlWebFormattingAll
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = False
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = True
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    qt.Delete
    Set qt = Nothing
    Exit Sub
errhdl:
    Debug.Print err.Description
    
    Resume Next
End Sub


Function get_sig_request(request As Object, ByVal secretkey As String, ByVal baseurl As String)
    'create string from request dictionary that contain apikey, command, parameter
    Dim request_str As String
    
    Dim arr() As String
    ReDim arr(request.Count)
    Dim i As Integer
    i = 0
    Dim x As Variant
    
    For Each x In request
        arr(i) = x + "=" + encodeURL(request.Item(x))
        i = i + 1
    Next
    
    Dim tsig_str, sig_str As String
    tsig_str = Join(arr, "&")
    request_str = Mid(tsig_str, 1, Len(tsig_str) - 1)

    For Each x In request
        request.Item(x) = Replace(LCase((encodeURL(request.Item(x)))), "+", "%20")
    Next
    
    i = 0
    For Each x In request
        arr(i) = x + "=" + request.Item(x)
        i = i + 1
    Next
    
    
    Call Alphabetically_SortArray(arr)

    tsig_str = Join(arr, "&")
    sig_str = Mid(tsig_str, 2, Len(tsig_str) - 1)

    
    Dim sig As String
    sig = encodeURL(Base64_HMACSHA1(sig_str, secretkey))
    get_sig_request = baseurl + request_str + "&signature=" + sig
End Function


Public Function encodeURL(ByVal str As String)
    Dim ScriptEngine As Object
    Dim encoded As String

    Set ScriptEngine = CreateObject("scriptcontrol")
    ScriptEngine.Language = "JScript"

    encoded = ScriptEngine.Run("encodeURIComponent", str)

    encodeURL = encoded
End Function



Public Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey As String)
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = EncodeBase64(bytes)
    Set asc = Nothing
    Set enc = Nothing
End Function


Private Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML = New MSXML2.DOMDocument60

    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function


Sub clear_sheets(ByVal sheetnum As String)
    Worksheets(sheetnum).Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub


Sub make_pretty(deletestr As String)
    Cells.Select
    Selection.Font.Size = 10
    Range("A1").Select
    Cells.Replace What:=deletestr, Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub


Sub Alphabetically_SortArray(myArray As Variant)
    Dim x As Long, y As Long
    Dim TempTxt1 As String
    Dim TempTxt2 As String

    For x = LBound(myArray) To UBound(myArray)
        For y = x To UBound(myArray)
            If UCase(myArray(y)) < UCase(myArray(x)) Then
                TempTxt1 = myArray(x)
                TempTxt2 = myArray(y)
                myArray(x) = TempTxt2
                myArray(y) = TempTxt1
            End If
        Next y
    Next x
End Sub


Sub delete_row(sheet As String, col As String)
    Worksheets(sheet).Select
    Dim i As Integer
    i = 3
    Dim prvvalue, nxtvalue As String
    prvvalue = Range(col & i).Value
    nxtvalue = Range(col & (i + 1)).Value
    Do While nxtvalue <> ""
        If prvvalue = nxtvalue Then
            Rows(i + 1).EntireRow.Delete
            nxtvalue = Range(col & (i + 1)).Value
        Else
            prvvalue = nxtvalue
            i = i + 1
            nxtvalue = Range(col & (i + 1)).Value
        End If
    Loop
End Sub

