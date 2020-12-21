Attribute VB_Name = "modListViewControl"

Public Sub SetHeads()
With FrmData
    ' Create ListView Headers
    Set LVName = .ListView1
    Set Colx = .ListView1.ColumnHeaders.Add(, , "ID")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Year")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Title")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Purchased")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Cost")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Medium")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Awards")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Other")
    Set Colx = .ListView1.ColumnHeaders.Add(, , "Condition")

    LV_Load    'Use column widths saved from last session
End With

End Sub

Public Sub ListLoad()

lvRefresh

End Sub

Public Sub LV_Unload()

With LVName
    'Save the Position of each of the ColumnHeader
    'Objects so we can load them the next time
    'the user starts the program
    For Each Colx In .ColumnHeaders
        SaveSetting App.Title, "Settings", "Col" & Colx.Index, _
                    .ColumnHeaders(Colx.Index).Position
        SaveSetting App.Title, "Settings", "ColWidth" & Colx.Index, _
                    .ColumnHeaders(Colx.Index).Width
    Next
End With

End Sub

Public Sub LV_Load()

With LVName
    'Let the user reorder the columns
    .AllowColumnReorder = True
    
    'Set view to report
    .View = lvwReport

    'Loop though each ColumnHeader object and set the
    'position of it dependent on what the user did
    'the last time
    For Each Colx In .ColumnHeaders
        Colx.Position = GetSetting(App.Title, "Settings", "Col" & Colx.Index, Colx.Index)
        Colx.Width = GetSetting(App.Title, "Settings", "ColWidth" & Colx.Index, Colx.Width)
    Next
End With

End Sub

Public Sub SortListView(ByVal lvw As MSComctlLib.ListView, _
     ByVal colHdr As MSComctlLib.ColumnHeader)
     'Sort by clicked ListView Column
'--set the sortkey to the column header's index - 1
lvw.SortKey = colHdr.Index - 1
lvw.Sorted = True

'--toggle the sort order between ascending & descending
lvw.SortOrder = 1 Xor lvw.SortOrder
End Sub

Public Function ClearForm()
 Dim ctl As Control
    
    For Each ctl In FrmData.Controls
        If TypeOf ctl Is TextBox Then ctl.Text = ""
    Next ctl

    For Each ctl In FrmData.Controls
        If TypeOf ctl Is ListView Then LVName.ListItems.Clear
    Next ctl
    
    FrmData.SSTab1.Tab = 1
End Function

Public Function lvRefresh()
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\vidz.mdb")
Set rs = db.OpenRecordset("vid")

Dim TmpID As Double
Dim Tmp1999 As Double
Dim Tmp1996 As Double
Tmp1996 = 0
    ClearForm
    Do While Not rs.EOF
    Set Itm = LVName.ListItems.Add(1)
    TmpID = Format(Format(rs!id, "00000"), "@@@@@")    'Create leading Zeros
    Itm.Text = Format(Format(TmpID, "00000"), "@@@@@") 'Add leading Zeros for column sort
    Itm.SubItems(1) = rs!field1
    Itm.SubItems(2) = rs!field2
    Itm.SubItems(3) = rs!field3
    Itm.SubItems(4) = rs!field4
    Itm.SubItems(5) = rs!field5
    Itm.SubItems(6) = rs!field6
    Itm.SubItems(7) = rs!field7
    Itm.SubItems(8) = rs!field8
    
    'Add other years as needed
    If Trim(rs!field1) = "1996" Then        'Get 1996 totals
    Tmp1996 = Tmp1996 + Format(Format(rs!field4, "0.00"), "@@@@@@@@")
    FrmData.txtTot1996.Text = Format(Format(Tmp1996, "0.00"), "@@@@@@@")
    End If
    
    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function

'In case you have a need to send a Null to a file field
'This can also be achieved by not moving anything to the field
Public Function NullIt(ctl As Control) As Variant
    If TypeOf ctl Is TextBox Then
        If ctl = "" Then
            NullIt = Null
        Else
            NullIt = ctl
        End If
    End If
End Function

