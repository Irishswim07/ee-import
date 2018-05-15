Attribute VB_Name = "CDOT_EE_Import"
'-----------------------------------------------------------------------------------
'CDOT EEMA Project Item Search and Import to Excel
'-----------------------------------------------------------------------------------
'
'By: Alan Carter
'Last Update: 05/07/2018
'
'-----------------------------------------------------------------------------------
'
'This code is used to automatically search the CDOT EEMA Project Item Search
'and import the results to an Excel table for flexible data interpretation
'
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'-----------------------
'New Table Code
'-----------------------

Public Sub CDOT_EEMA_ProjectItem()

Dim ie As SHDocVw.InternetExplorer
Dim doc As Object
Dim rng As Range
Dim tbl As Object
Dim rw As Object
Dim cl As Object
Dim tabno As Long
Dim nextrow As Long
Dim I As Long
Dim y As Long
Dim x As Long
Dim z As Integer
Dim tblrange As Long
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet

x = 5: y = 1: z = 1


Item = Range("C2").Value
item1 = Range("G2").Value
datefrom = Format(Range("C3").Text, "mm/dd/yy")
dateto = Format(Range("G3").Text, "mm/dd/yy")


Set ie = New SHDocVw.InternetExplorerMedium

ie.Visible = False 'Toggle on or off the display of IE
ie.Navigate2 ("https://apps.coloradodot.info/EEMA_ProjectItem/projectItem.cfm")

Do

DoEvents

Loop Until ie.ReadyState = 4

'ie READYSTATE has 5 different status codes, here we are using status 4:
'Uninitialized = 0
'Loading = 1
'Loaded = 2
'Interactive = 3
'Complete = 4

ie.Document.getElementById("ITEM").Value = Item
ie.Document.getElementById("ITEM1").Value = item1
ie.Document.getElementById("DATE_FROM").Value = datefrom
ie.Document.getElementById("DATE_TO").Value = dateto

Set AllInputs = ie.Document.getElementsByTagName("input")

    For Each hyper_link In AllInputs

        If hyper_link.Type = "submit" Then

            hyper_link.Click

            Exit For

        End If

    Next

Do

DoEvents

Loop Until ie.ReadyState = 3

Do

DoEvents

Loop Until ie.ReadyState = 4

'--------------------------------------------------------------------------
' Copy and Paste Table Information in Excel
'--------------------------------------------------------------------------

nextrow = 4

            For Each tbl In ie.Document.getElementsByTagName("TABLE")
                    tabno = tabno + 1
                    
                    If tabno <> 3 Then GoTo MyLabel
                    
                        nextrow = nextrow + 1
                        Set rng = ActiveSheet.Range("A" & nextrow)
                        'rng.Offset(, -1) = "Table " & tabno
                        totalRow = tbl.Rows.Length
                        
                        For Each rw In tbl.Rows
                            For Each cl In rw.Cells
                                rng.Value = cl.outerText
                                Set rng = rng.Offset(, 1)
                                I = I + 1
                            Next cl
                            
                            nextrow = nextrow + 1
                            Set rng = rng.Offset(1, -I)
                            I = 0
                            Application.StatusBar = "Progress: " & z & " of " & totalRow & " : " & Format(z / totalRow, "0%")
                            z = z + 1
            
                        Next rw
                           
MyLabel:
                Next tbl
                
Application.StatusBar = False

            'ws.Cells.ClearFormats
  
'ie.Quit

'--------------------------------------------------------------------------
' Create Table in Excel and Sort by Quantity
'--------------------------------------------------------------------------

tbltitle = InputBox("Enter Name for New Table", "Table Name Entry", "Table_")
tblrange = Range("A" & Rows.Count).End(xlUp).Row
ActiveSheet.ListObjects.Add(xlSrcRange, Range("A" & x, "Q" & tblrange), , xlYes).Name = _
        tbltitle
        
ActiveSheet.ListObjects(tbltitle).Sort.SortFields.Clear
ActiveSheet.ListObjects(tbltitle).Sort.SortFields.Add _
        Key:=Range("H5"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(tbltitle).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



End Sub

'-----------------------
'Append Table Code
'-----------------------

Public Sub CDOT_EEMA_ProjectItem_Append()

Dim ie As SHDocVw.InternetExplorer
Dim doc As Object
Dim rng As Range
Dim tbl As Object
Dim rw As Object
Dim cl As Object
Dim tabno As Long
Dim nextrow As Long
Dim I As Long
Dim y As Long
Dim x As Long
Dim z As Integer
Dim tblrange As Long
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim dateRange(0 To 1) As Long

x = Range("A" & Rows.Count).End(xlUp).Row: y = 1: z = 1


Item = Range("C2").Value
item1 = Range("G2").Value
datefrom = Format(Range("C3").Text, "mm/dd/yy")
dateto = Format(Range("G3").Text, "mm/dd/yy")

'-------------------------------------------------------------------------
'Search Excel Table for duplicates of user-selected item(s) AND date range
'-------------------------------------------------------------------------

Set tbl = Sheets("MAIN").ListObjects("Table_Main").ListColumns(7).DataBodyRange
itemNum = Replace(Item, "-", "")
item1Num = Replace(item1, "-", "")

For Each Cel In tbl
        'Cel.Select
    If Replace(Cel.Value, "-", "") >= itemNum And Replace(Cel.Value, "-", "") <= item1Num Then
        'Cel.Offset(0, 6).Select
        If CLng(DateValue(Cel.Offset(0, 6).Value)) >= CLng(DateValue(datefrom)) And CLng(DateValue(Cel.Offset(0, 6).Value)) <= CLng(DateValue(dateto)) Then
            response = MsgBox("Selected Item(s) and Time Period Exist in Table, Proceed with Import?", vbYesNoCancel)
                If response = vbNo Or response = vbCancel Then
                    Exit Sub
                Else
                    GoTo Import
                End If
        End If
    End If
    
    Next

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Import:

Set ie = New SHDocVw.InternetExplorerMedium

ie.Visible = False 'Toggle on or off the display of IE
ie.Navigate2 ("https://apps.coloradodot.info/EEMA_ProjectItem/projectItem.cfm")

Do

DoEvents

Loop Until ie.ReadyState = 4

'ie READYSTATE has 5 different status codes, here we are using status 4:
'Uninitialized = 0
'Loading = 1
'Loaded = 2
'Interactive = 3
'Complete = 4

ie.Document.getElementById("ITEM").Value = Item
ie.Document.getElementById("ITEM1").Value = item1
ie.Document.getElementById("DATE_FROM").Value = datefrom
ie.Document.getElementById("DATE_TO").Value = dateto

Set AllInputs = ie.Document.getElementsByTagName("input")

    For Each hyper_link In AllInputs

        If hyper_link.Type = "submit" Then

            hyper_link.Click

            Exit For

        End If

    Next

Do

DoEvents

Loop Until ie.ReadyState = 3

Do

DoEvents

Loop Until ie.ReadyState = 4

'--------------------------------------------------------------------------
' Copy and Paste Table Information in Excel
'--------------------------------------------------------------------------

nextrow = x

            For Each tbl In ie.Document.getElementsByTagName("TABLE")
                    tabno = tabno + 1
                    
                    If tabno <> 3 Then GoTo MyLabel
                    
                        nextrow = nextrow + 1
                        Set rng = ActiveSheet.Range("A" & nextrow)
                        'rng.Offset(, -1) = "Table " & tabno
                        totalRow = tbl.Rows.Length
                          
                        For Each rw In tbl.Rows
                        If totalRow = 1 Then GoTo NoData
                        
                            For Each cl In rw.Cells
                                rng.Value = cl.outerText
                                Set rng = rng.Offset(, 1)
                                I = I + 1
                            Next cl
                            
                            nextrow = nextrow + 1
                            Set rng = rng.Offset(1, -I)
                            I = 0
                            Application.StatusBar = "Progress: " & z & " of " & totalRow & " : " & Format(z / totalRow, "0%")
                            z = z + 1
            
MyLabel1:
                Next rw
MyLabel:
                Next tbl
                
                
MsgBox "Number of Records Imported: " & totalRow, vbOKOnly, "Results of Import"
Application.StatusBar = False

'    'Remove Duplicates from import
'    ActiveSheet.Range(ActiveSheet.ListObjects(1).Name).RemoveDuplicates Columns:=Array(12), Header:=xlYes

Exit Sub

            'ws.Cells.ClearFormats
  
'ie.Quit
NoData:
                MsgBox "No Records exist for the selected Bid Item(s) and Time Range", vbCritical
                Exit Sub
End Sub


