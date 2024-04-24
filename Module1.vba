Attribute VB_Name = "Module1"
Sub GETFromURL()
    Dim URL As String
    Dim WinHttpReq As Object
    Dim ws As Worksheet
    Dim wsExists As Boolean
    
    ' Define the URL of the CSV file
    URL = "https://radioid.net/static/user.csv"
    
    ' Check if the worksheet named "user" already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("user")
    On Error GoTo 0
    
    ' If the worksheet exists, delete it
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' Suppress alert messages
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create a new instance of WinHttpRequest
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Open a connection to the URL and download the file
    WinHttpReq.Open "GET", URL, False
    WinHttpReq.send
    
    ' Check if the request was successful
    If WinHttpReq.Status = 200 Then
        ' Create a new worksheet named "user"
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "user"
        
        ' Import the downloaded CSV data into the new worksheet
        With ws.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=ws.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh
        End With
        'MsgBox "CSV data has been imported into a new sheet named 'user' in the active workbook."
    Else
        MsgBox "Failed to download file from: " & URL
        Exit Sub
    End If
    Worksheets("user").Columns("A:G").AutoFit
End Sub

Sub DMRID()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim countryName As String
    Dim stateName As String
    
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Sheets("user")
    
    'remove nonsense data
    Columns("C:F").Select
    Selection.Replace What:="None", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
    Selection.Replace What:="All Regions", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    ' Find the last row with data in column G
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Loop through each row
    For i = 1 To lastRow
        ' Get the value in column G and F
        countryName = ws.Cells(i, "G").Value
        stateName = ws.Cells(i, "F").Value
        If Not IsEmpty(ws.Cells(i, "e")) And Not Application.WorksheetFunction.IsText(ws.Cells(i, "e")) Then
            ws.Cells(i, "e").Value = "Inv.City"
        End If
        
        ' processing data based on the value in column G
        Select Case countryName
        
            'For US
            Case "United States"
                Select Case stateName
                    Case "Alabama"
                        ws.Cells(i, "F").Value = "AL"
                    Case "Alaska"
                        ws.Cells(i, "F").Value = "AK"
                    Case "Arizona"
                        ws.Cells(i, "F").Value = "AZ"
                    Case "Arkansas"
                        ws.Cells(i, "F").Value = "AR"
                    Case "California"
                        ws.Cells(i, "F").Value = "CA"
                    Case "Colorado"
                        ws.Cells(i, "F").Value = "CO"
                    Case "Connecticut"
                        ws.Cells(i, "F").Value = "CT"
                    Case "District of Columbia"
                        ws.Cells(i, "F").Value = "DC"
                    Case "Delaware"
                        ws.Cells(i, "F").Value = "DE"
                    Case "Florida"
                        ws.Cells(i, "F").Value = "FL"
                    Case "Georgia"
                        ws.Cells(i, "F").Value = "GA"
                    Case "Hawaii"
                        ws.Cells(i, "F").Value = "HI"
                    Case "Idaho"
                        ws.Cells(i, "F").Value = "ID"
                    Case "Illinois"
                        ws.Cells(i, "F").Value = "IL"
                    Case "Indiana"
                        ws.Cells(i, "F").Value = "IN"
                    Case "Iowa"
                        ws.Cells(i, "F").Value = "IA"
                    Case "Kansas"
                        ws.Cells(i, "F").Value = "KS"
                    Case "Kentucky"
                        ws.Cells(i, "F").Value = "KY"
                    Case "Louisiana"
                        ws.Cells(i, "F").Value = "LA"
                    Case "Maine"
                        ws.Cells(i, "F").Value = "ME"
                    Case "Maryland"
                        ws.Cells(i, "F").Value = "MD"
                    Case "Massachusetts"
                        ws.Cells(i, "F").Value = "MA"
                    Case "Michigan"
                        ws.Cells(i, "F").Value = "MI"
                    Case "Minnesota"
                        ws.Cells(i, "F").Value = "MN"
                    Case "Mississippi"
                        ws.Cells(i, "F").Value = "MS"
                    Case "Missouri"
                        ws.Cells(i, "F").Value = "MO"
                    Case "Montana"
                        ws.Cells(i, "F").Value = "MT"
                    Case "Nebraska"
                        ws.Cells(i, "F").Value = "NE"
                    Case "Nevada"
                        ws.Cells(i, "F").Value = "NV"
                    Case "New Hampshire"
                        ws.Cells(i, "F").Value = "NH"
                    Case "New Jersey"
                        ws.Cells(i, "F").Value = "NJ"
                    Case "New Mexico"
                        ws.Cells(i, "F").Value = "NM"
                    Case "New York"
                        ws.Cells(i, "F").Value = "NY"
                    Case "North Carolina"
                        ws.Cells(i, "F").Value = "NC"
                    Case "North Dakota"
                        ws.Cells(i, "F").Value = "ND"
                    Case "Ohio"
                        ws.Cells(i, "F").Value = "OH"
                    Case "Oklahoma"
                        ws.Cells(i, "F").Value = "OK"
                    Case "Oregon"
                        ws.Cells(i, "F").Value = "OR"
                    Case "Pennsylvania"
                        ws.Cells(i, "F").Value = "PA"
                    Case "Rhode Island"
                        ws.Cells(i, "F").Value = "RI"
                    Case "South Carolina"
                        ws.Cells(i, "F").Value = "SC"
                    Case "South Dakota"
                        ws.Cells(i, "F").Value = "SD"
                    Case "Tennessee"
                        ws.Cells(i, "F").Value = "TN"
                    Case "Texas"
                        ws.Cells(i, "F").Value = "TX"
                    Case "Utah"
                        ws.Cells(i, "F").Value = "UT"
                    Case "Vermont"
                        ws.Cells(i, "F").Value = "VT"
                    Case "Virginia"
                        ws.Cells(i, "F").Value = "VA"
                    Case "Washington"
                        ws.Cells(i, "F").Value = "WA"
                    Case "West Virginia"
                        ws.Cells(i, "F").Value = "WV"
                    Case "Wisconsin"
                        ws.Cells(i, "F").Value = "WI"
                    Case "Wyoming"
                        ws.Cells(i, "F").Value = "WY"
                    Case Else
                        ' If the state is not in the list, do nothing or handle differently
                End Select
                
                'if city name is longer than 18 charecters make it 18
                'why 18? Because .?? is 3 charecters and 18+3 = 21
                If (Len(ws.Cells(i, "e").Value)) > 18 Then
                    ws.Cells(i, "e").Value = Left(ws.Cells(i, "e").Value, 18)
                End If
                If Not IsEmpty(ws.Cells(i, "e")) Then
                    ws.Cells(i, "G").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "f").Value
                End If
                
            'For UK
            Case "United Kingdom"
                ws.Cells(i, "G").Value = "GB"
                
                If (Len(ws.Cells(i, "e").Value) + Len(ws.Cells(i, "f").Value) + Len(ws.Cells(i, "g").Value)) < 20 And Not IsEmpty(ws.Cells(i, "e")) And Not IsEmpty(ws.Cells(i, "f")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "f").Value + "." + ws.Cells(i, "G").Value
                ElseIf (Len(ws.Cells(i, "e").Value) + Len(ws.Cells(i, "g").Value)) < 21 And Not IsEmpty(ws.Cells(i, "e")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "G").Value
                ElseIf (Len(ws.Cells(i, "f").Value) + Len(ws.Cells(i, "g").Value)) < 21 And Not IsEmpty(ws.Cells(i, "f")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "f").Value + "." + ws.Cells(i, "G").Value
                End If
                
            'For Thailand
            Case "Thailand"
                ws.Cells(i, "G").Value = "TH"
                
                If Not IsEmpty(ws.Cells(i, "f")) Then
                ws.Cells(i, "g").Value = ws.Cells(i, "f").Value + "." + ws.Cells(i, "g").Value
                End If
                
            'For Bosnia Hercegovina
            Case "Bosnia and Hercegovina"
                ws.Cells(i, "G").Value = "Bosnia.Hercegovina"
                
            'For Korea
            Case "Korea Republic of"
                ws.Cells(i, "G").Value = "Korea"
                
                If (Len(ws.Cells(i, "e").Value) + Len(ws.Cells(i, "f").Value) + Len(ws.Cells(i, "g").Value)) < 20 And Not IsEmpty(ws.Cells(i, "e")) And Not IsEmpty(ws.Cells(i, "f")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "f").Value + "." + ws.Cells(i, "g").Value
                ElseIf (Len(ws.Cells(i, "e").Value) + Len(ws.Cells(i, "g").Value)) < 21 And Not IsEmpty(ws.Cells(i, "e")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "g").Value
                ElseIf (Len(ws.Cells(i, "f").Value) + Len(ws.Cells(i, "g").Value)) < 21 And Not IsEmpty(ws.Cells(i, "f")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "f").Value + "." + ws.Cells(i, "g").Value
                End If
                
            'For Canada
            Case "Canada"
                ws.Cells(i, "G").Value = "CAN"
                Select Case stateName
                    Case "Alberta"
                        ws.Cells(i, "F").Value = "AB"
                    Case "British Columbia"
                        ws.Cells(i, "F").Value = "BC"
                    Case "Manitoba"
                        ws.Cells(i, "F").Value = "MB"
                    Case "New Brunswick"
                        ws.Cells(i, "F").Value = "NB"
                    Case "Newfoundland"
                        ws.Cells(i, "F").Value = "NL"
                    Case "Nova Scotia"
                        ws.Cells(i, "F").Value = "NS"
                    Case "Ontario"
                        ws.Cells(i, "F").Value = "ON"
                    Case "Prince Edward Island"
                        ws.Cells(i, "F").Value = "PE"
                    Case "Quebec"
                        ws.Cells(i, "F").Value = "QC"
                    Case "Saskatchewan"
                        ws.Cells(i, "F").Value = "SK"
                    Case "Northwest Territories"
                        ws.Cells(i, "F").Value = "NT"
                    Case "Nunavut"
                        ws.Cells(i, "F").Value = "NU"
                    Case "Yukon"
                        ws.Cells(i, "F").Value = "YT"
                    Case Else
                        ' If the state is not in the list, do nothing or handle differently
                End Select
                'if the city name is longer than 14 charecters make it 14
                'why 14? Because .??.CAN is 7 charecters and 14+7 = 21
                If (Len(ws.Cells(i, "e").Value)) > 14 Then
                    ws.Cells(i, "e").Value = Left(ws.Cells(i, "e").Value, 14)
                End If
                ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "f").Value + "." + ws.Cells(i, "g").Value

            
            'for the rest of the world
            Case Else
                
                If (Len(ws.Cells(i, "e").Value) + Len(ws.Cells(i, "f").Value) + Len(ws.Cells(i, "g").Value)) < 20 And Not IsEmpty(ws.Cells(i, "e")) And Not IsEmpty(ws.Cells(i, "f")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "f").Value + "." + ws.Cells(i, "g").Value
                ElseIf (Len(ws.Cells(i, "e").Value) + Len(ws.Cells(i, "g").Value)) < 21 And Not IsEmpty(ws.Cells(i, "e")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "e").Value + "." + ws.Cells(i, "g").Value
                ElseIf (Len(ws.Cells(i, "f").Value) + Len(ws.Cells(i, "g").Value)) < 21 And Not IsEmpty(ws.Cells(i, "f")) Then
                    ws.Cells(i, "g").Value = ws.Cells(i, "f").Value + "." + ws.Cells(i, "g").Value
                End If
                
        End Select
        'dealing with First_Name and Last_Name
        If Not IsEmpty(ws.Cells(i, "c")) And Not Application.WorksheetFunction.IsText(ws.Cells(i, "C")) Then
            ws.Cells(i, "c").Value = "Inv.F.Name"
        End If
        If Not IsEmpty(ws.Cells(i, "d")) And Not Application.WorksheetFunction.IsText(ws.Cells(i, "D")) Then
            ws.Cells(i, "D").Value = "Inv.L.Name"
        End If
        'if First_Name is longer than 21 charecters make it 21
        If (Len(ws.Cells(i, "C").Value)) > 21 Then
            ws.Cells(i, "C").Value = Left(ws.Cells(i, "C").Value, 21)
        End If
        'if First_Name + Last_Name is longer than one line, only disply First_Name
        If (Len(ws.Cells(i, "C").Value) + Len(ws.Cells(i, "D").Value)) > 20 Then
            ws.Cells(i, "D").Value = ""
        End If
    Next i
    
    'remove remaining (space) from column G
    Columns("G:G").Select
        Selection.Replace What:=" ", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'we no longer need column E and F
    Columns("E:F").Delete
    
    'make the sheet presentable
    Cells(1, 1).Select
    Application.ScreenUpdating = True
    Worksheets("user").Columns("A:B").AutoFit
    Worksheets("user").Columns("C:E").ColumnWidth = 21
        
End Sub

Sub ExportToCSV()
    On Error GoTo ErrorHandler
    
    Dim MyFileName As String
    Dim DesktopFilePath As String
    Dim WorkbookFilePath As String
    
    Application.ScreenUpdating = False
    ' Save the workbook
    ThisWorkbook.Save
    
    ' Define the file name
    MyFileName = "user.csv"
    
    ' Define the desktop file path
    DesktopFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & MyFileName
    
    ' Check if the file already exists on desktop, and delete it if it does
    If Dir(DesktopFilePath) <> "" Then
        Kill DesktopFilePath
    End If
    
    ' Save the worksheet as CSV on desktop
    ThisWorkbook.Sheets("user").Copy
    ActiveWorkbook.SaveAs Filename:=DesktopFilePath, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close SaveChanges:=False
    
    ' Define the workbook file path
    ' If ThisWorkbook.Path <> "" Then
    WorkbookFilePath = ThisWorkbook.Path & "\" & MyFileName
    
     ' Check if the file already exists on this folder, and delete it if it does
    If Dir(WorkbookFilePath) <> "" Then
        Kill WorkbookFilePath
    End If
        
    ' Save the worksheet as CSV in workbook's folder
    ThisWorkbook.Sheets("user").Copy
    ActiveWorkbook.SaveAs Filename:=WorkbookFilePath, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Sub Start()
    GETFromURL
    DMRID
    ExportToCSV
End Sub
