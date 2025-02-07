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
    'Worksheets("user").Columns("A:G").AutoFit
End Sub

Sub DMRID()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim countryName As String
    Dim stateName As String
    Dim cityName As String
    Dim firstName As String
    Dim lencountryName As Integer
    Dim lenstateName As Integer
    Dim lencityName As Integer
    
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
        ' Get the value in column G, F and E
        countryName = ws.Cells(i, "G").Value
        stateName = ws.Cells(i, "F").Value
        
        'some people put weird data in column e, we have to filter that out
        If IsEmpty(ws.Cells(i, "e")) Then
            cityName = ""
            ElseIf Application.WorksheetFunction.IsText(ws.Cells(i, "e")) Then
                cityName = ws.Cells(i, "e").Value
                Else
                    cityName = "Inv.City"
        End If
        
        ' processing data based on the value in column G
        Select Case countryName
        
            'For US
            Case "United States"
                Select Case stateName
                    Case "Alabama"
                        stateName = "AL"
                    Case "Alaska"
                        stateName = "AK"
                    Case "Arizona"
                        stateName = "AZ"
                    Case "Arkansas"
                        stateName = "AR"
                    Case "California"
                        stateName = "CA"
                    Case "Colorado"
                        stateName = "CO"
                    Case "Connecticut"
                        stateName = "CT"
                    Case "District of Columbia"
                        stateName = "DC"
                    Case "District Of Columbia"
                        stateName = "DC"
                    Case "Delaware"
                        stateName = "DE"
                    Case "Florida"
                        stateName = "FL"
                    Case "Georgia"
                        stateName = "GA"
                    Case "Hawaii"
                        stateName = "HI"
                    Case "Idaho"
                        stateName = "ID"
                    Case "Illinois"
                        stateName = "IL"
                    Case "Indiana"
                        stateName = "IN"
                    Case "Iowa"
                        stateName = "IA"
                    Case "Kansas"
                        stateName = "KS"
                    Case "Kentucky"
                        stateName = "KY"
                    Case "Louisiana"
                        stateName = "LA"
                    Case "Maine"
                        stateName = "ME"
                    Case "Maryland"
                        stateName = "MD"
                    Case "Massachusetts"
                        stateName = "MA"
                    Case "Michigan"
                        stateName = "MI"
                    Case "Minnesota"
                        stateName = "MN"
                    Case "Mississippi"
                        stateName = "MS"
                    Case "Missouri"
                        stateName = "MO"
                    Case "Montana"
                        stateName = "MT"
                    Case "Nebraska"
                        stateName = "NE"
                    Case "Nevada"
                        stateName = "NV"
                    Case "New Hampshire"
                        stateName = "NH"
                    Case "New Jersey"
                        stateName = "NJ"
                    Case "New Mexico"
                        stateName = "NM"
                    Case "New York"
                        stateName = "NY"
                    Case "North Carolina"
                        stateName = "NC"
                    Case "North Dakota"
                        stateName = "ND"
                    Case "Ohio"
                        stateName = "OH"
                    Case "Oklahoma"
                        stateName = "OK"
                    Case "Oregon"
                        stateName = "OR"
                    Case "Pennsylvania"
                        stateName = "PA"
                    Case "Rhode Island"
                        stateName = "RI"
                    Case "South Carolina"
                        stateName = "SC"
                    Case "South Dakota"
                        stateName = "SD"
                    Case "Tennessee"
                        stateName = "TN"
                    Case "Texas"
                        stateName = "TX"
                    Case "Utah"
                        stateName = "UT"
                    Case "Vermont"
                        stateName = "VT"
                    Case "Virginia"
                        stateName = "VA"
                    Case "Washington"
                        stateName = "WA"
                    Case "West Virginia"
                        stateName = "WV"
                    Case "Wisconsin"
                        stateName = "WI"
                    Case "Wyoming"
                        stateName = "WY"
                    Case Else
                        ' If the state is not in the list, do nothing or handle differently
                End Select
                
                If cityName <> "" Then
                    'if city name is longer than 18 charecters make it 18
                    'why 18? Because .?? is 3 charecters and 18+3 = 21
                    If Len(cityName) > 18 Then
                        cityName = Left(cityName, 18)
                    End If
                    ws.Cells(i, "G").Value = cityName + "." + stateName
                Else
                    ws.Cells(i, "G").Value = stateName
                End If
                
            'For Canada
            Case "Canada"
                countryName = "CAN"
                Select Case stateName
                    Case "Alberta"
                        stateName = "AB"
                    Case "British Columbia"
                        stateName = "BC"
                    Case "Manitoba"
                        stateName = "MB"
                    Case "New Brunswick"
                        stateName = "NB"
                    Case "Newfoundland"
                        stateName = "NL"
                    Case "Nova Scotia"
                        stateName = "NS"
                    Case "Ontario"
                        stateName = "ON"
                    Case "Prince Edward Island"
                        stateName = "PE"
                    Case "Quebec"
                        stateName = "QC"
                    Case "Saskatchewan"
                        stateName = "SK"
                    Case "Northern Territories"
                        stateName = "NT"
                    Case "Nunavut"
                        stateName = "NU"
                    Case "Yukon"
                        stateName = "YT"
                    Case Else
                        ' If the state is not in the list, do nothing or handle differently
                End Select
                
                If cityName <> "" Then
                    'if the city name is longer than 14 charecters make it 14
                    'why 14? Because .??.CAN is 7 charecters and 14+7 = 21
                    If Len(cityName) > 14 Then
                        cityName = Left(cityName, 14)
                    End If
                    ws.Cells(i, "g").Value = cityName + "." + stateName + "." + countryName
                Else
                    ws.Cells(i, "G").Value = stateName + "." + countryName
                End If
                
            'For UK
            Case "United Kingdom"
                countryName = "GB"
                
                lencountryName = Len(countryName)
                lenstateName = Len(stateName)
                lencityName = Len(cityName)
                If (lencityName + lenstateName + lencountryName) < 20 And cityName <> "" And stateName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + stateName + "." + countryName
                ElseIf (lencityName + lencountryName) < 21 And cityName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + countryName
                ElseIf (lenstateName + lencountryName) < 21 And stateName <> "" Then
                    ws.Cells(i, "g").Value = stateName + "." + countryName
                Else
                    ws.Cells(i, "g").Value = countryName
                End If
                
            'For Thailand
            Case "Thailand"
                countryName = "TH"
                
                If stateName = "Phra Nakhon Si Ayutthaya" Then
                    stateName = "Ayutthaya"
                End If
                If stateName <> "" Then
                    ws.Cells(i, "G").Value = stateName + "." + countryName
                Else
                    ws.Cells(i, "G").Value = countryName
                End If
                
            'For Bosnia Hercegovina
            Case "Bosnia and Hercegovina"
                ws.Cells(i, "G").Value = "Bosnia.Hercegovina"
                
            'For Trinidad and Tobago
            Case "Trinidad and Tobago"
                'doing nothing
            
            'For U.S. Virgin Islands
            Case "U.S. Virgin Islands"
                ws.Cells(i, "G").Value = "U.S.Virgin.Islands"
                
            'For United Arab Emirates
            Case "United Arab Emirates"
                countryName = "UAE"
                
                If stateName <> "" Then
                    ws.Cells(i, "G").Value = stateName + "." + countryName
                ElseIf cityName <> "" Then
                    If Len(cityName) > 17 Then
                        cityName = Left(cityName, 17)
                    End If
                    ws.Cells(i, "G").Value = cityName + "." + countryName
                Else
                    ws.Cells(i, "G").Value = countryName
                End If
                
            'For Korea
            Case "Korea Republic of"
                countryName = "Korea"
                
                lencountryName = Len(countryName)
                lenstateName = Len(stateName)
                lencityName = Len(cityName)
                If (lencityName + lenstateName + lencountryName) < 20 And cityName <> "" And stateName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + stateName + "." + countryName
                ElseIf (lencityName + lencountryName) < 21 And cityName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + countryName
                ElseIf (lenstateName + lencountryName) < 21 And stateName <> "" Then
                    ws.Cells(i, "g").Value = stateName + "." + countryName
                Else
                    ws.Cells(i, "g").Value = countryName
                End If
                
            'For Argentina Republic
            Case "Argentina Republic"
                countryName = "Argentina"
                
                lencountryName = Len(countryName)
                lenstateName = Len(stateName)
                lencityName = Len(cityName)
                If (lencityName + lenstateName + lencountryName) < 20 And cityName <> "" And stateName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + stateName + "." + countryName
                ElseIf (lencityName + lencountryName) < 21 And cityName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + countryName
                ElseIf (lenstateName + lencountryName) < 21 And stateName <> "" Then
                    ws.Cells(i, "g").Value = stateName + "." + countryName
                Else
                    ws.Cells(i, "g").Value = countryName
                End If
            
            'for the rest of the world
            Case Else
                
                lencountryName = Len(countryName)
                lenstateName = Len(stateName)
                lencityName = Len(cityName)
                If (lencityName + lenstateName + lencountryName) < 20 And cityName <> "" And stateName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + stateName + "." + countryName
                ElseIf (lencityName + lencountryName) < 21 And cityName <> "" Then
                    ws.Cells(i, "g").Value = cityName + "." + countryName
                ElseIf (lenstateName + lencountryName) < 21 And stateName <> "" Then
                    ws.Cells(i, "g").Value = stateName + "." + countryName
                Else
                    ws.Cells(i, "g").Value = countryName
                End If
                
        End Select
        
        'dealing with First_Name and Last_Name
        If IsEmpty(ws.Cells(i, "c")) Then
            firstName = ""
            ElseIf Application.WorksheetFunction.IsText(ws.Cells(i, "c")) Then
                firstName = ws.Cells(i, "c").Value
                Else
                    firstName = "Inv.F.Name"
                    ws.Cells(i, "C").Value = "Inv.F.Name"
        End If
        
        If Not IsEmpty(ws.Cells(i, "d")) And Not Application.WorksheetFunction.IsText(ws.Cells(i, "D")) Then
            ws.Cells(i, "D").Value = "Inv.L.Name"
        End If
        
        'if First_Name is longer than 21 charecters make it 21 and delete lastname
        If Len(firstName) > 21 Then
            ws.Cells(i, "C").Value = Left(firstName, 21)
            ws.Cells(i, "D").Value = ""
        'if First_Name + Last_Name is longer than one line, only disply First_Name
        ElseIf (Len(firstName) + Len(ws.Cells(i, "D").Value)) > 20 Then
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
    Range("E1").Value = "QTH"
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
