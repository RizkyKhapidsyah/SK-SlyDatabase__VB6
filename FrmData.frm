VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmData 
   Caption         =   "Sly's Database"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "Totals"
      TabPicture(0)   =   "FrmData.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(1)=   "txtTot1996"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Does the Donut surround the hole?"
      TabPicture(1)   =   "FrmData.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Add/Update/Delete"
      TabPicture(2)   =   "FrmData.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(3)=   "Label5"
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "Label8"
      Tab(2).Control(7)=   "Label9"
      Tab(2).Control(8)=   "Label10"
      Tab(2).Control(9)=   "txtYear"
      Tab(2).Control(10)=   "txtTitle"
      Tab(2).Control(11)=   "txtPurch"
      Tab(2).Control(12)=   "txtCost"
      Tab(2).Control(13)=   "txtMedium"
      Tab(2).Control(14)=   "txtAward"
      Tab(2).Control(15)=   "txtOther"
      Tab(2).Control(16)=   "txtCondition"
      Tab(2).Control(17)=   "cmdAdd"
      Tab(2).Control(18)=   "txtID"
      Tab(2).Control(19)=   "cmdDelete"
      Tab(2).Control(20)=   "cmdUpd"
      Tab(2).ControlCount=   21
      Begin VB.TextBox txtTot1996 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74040
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdUpd 
         Caption         =   "Update Record"
         Height          =   375
         Left            =   -71400
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Record"
         Height          =   375
         Left            =   -71400
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   -73680
         TabIndex        =   21
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Record"
         Height          =   375
         Left            =   -71400
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtCondition 
         Height          =   285
         Left            =   -73680
         TabIndex        =   19
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox txtOther 
         Height          =   285
         Left            =   -73680
         TabIndex        =   18
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox txtAward 
         Height          =   285
         Left            =   -73680
         TabIndex        =   17
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtMedium 
         Height          =   285
         Left            =   -73680
         TabIndex        =   16
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtCost 
         Height          =   285
         Left            =   -73680
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtPurch 
         Height          =   285
         Left            =   -73680
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   -73680
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   -73680
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Movies to Die for, or Snarly Flicks           Click Line Item to select for UpDate/Delete, or use as base for Add"
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11535
         Begin MSComctlLib.ListView ListView1 
            Height          =   5175
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   9128
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Label Label11 
         Caption         =   "1996"
         Height          =   375
         Left            =   -74520
         TabIndex        =   26
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Record Number"
         Height          =   375
         Left            =   -74760
         TabIndex        =   22
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Condition"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Other ........."
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Awards ......"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Medium ....."
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cost ............"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Purchased"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Title ............."
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Year ............."
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Title"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "FrmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Delete a record
Private Sub cmdDelete_Click()
'Open the database
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\vidz.mdb")
'Open the file
Set rs = db.OpenRecordset("vid")
        
While Not rs.EOF

If txtID = rs!id Then
    rs.Delete           'Delete from file
    GoTo Exit01
End If

rs.MoveNext
Wend

Exit01:

'Get File as is now
rs.Close
Set rs = Nothing
main
Set rs = db.OpenRecordset("vid")
rs.Close
Set rs = Nothing
main

End Sub

Private Sub Form_Activate()
SSTab1.Tab = 1   'Show Catolog
End Sub

Private Sub Form_Unload(Cancel As Integer)
LV_Unload    'Save ListView column width as set by user
End Sub

'Load the file records into the listview
Private Sub ListView1_Click()
With LVName

Dim Numfield As Integer
Numfield = Val(.SelectedItem)   'Remove Leading Zeros

FrmData.txtID.Text = Numfield
FrmData.txtYear.Text = .SelectedItem.SubItems(1)
FrmData.txtTitle.Text = .SelectedItem.SubItems(2)
FrmData.txtPurch.Text = .SelectedItem.SubItems(3)
FrmData.txtCost.Text = .SelectedItem.SubItems(4)
FrmData.txtMedium.Text = .SelectedItem.SubItems(5)
FrmData.txtAward.Text = .SelectedItem.SubItems(6)
FrmData.txtOther.Text = .SelectedItem.SubItems(7)
FrmData.txtCondition.Text = .SelectedItem.SubItems(8)

End With
End Sub

Private Sub cmdAdd_Click()
'Open the DataBase
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\vidz.mdb")
'Open the File
Set rs = db.OpenRecordset("vid")

  'Load the fields to be added to your file
            rs.AddNew
            rs!field1 = txtYear
            rs!field2 = txtTitle
            rs!field3 = txtPurch
            rs!field4 = txtCost
            rs!field5 = txtMedium
            rs!field6 = txtAward
            rs!field7 = txtOther
            rs!field8 = txtCondition
            rs.Update
     
    'Get the file as updated
    rs.Close
    Set rs = Nothing
    main
    Set rs = db.OpenRecordset("vid")
    rs.Close
    Set rs = Nothing
    main
    
End Sub
Private Sub cmdUpd_Click()
'Open the Database
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\vidz.mdb")
'Open the File
Set rs = db.OpenRecordset("vid")
        
 'Find the record to be updated
        While Not rs.EOF
     If rs!id = txtID.Text Then
 'Move data into the file fields
            rs.Edit
            rs!field1 = txtYear.Text
            rs!field2 = txtTitle.Text
            rs!field3 = txtPurch.Text
            rs!field4 = txtCost.Text
            rs!field5 = txtMedium.Text
            rs!field6 = txtAward
            rs!field7 = txtOther
            rs!field8 = txtCondition
            rs.Update
            rs.MoveLast
            rs.MoveNext
            GoTo Exit02
     Else
            rs.MoveNext
     End If
        Wend
     
'Get the updated file
Exit02:
    UpDated = False
    rs.Close
    Set rs = Nothing
    main
    Set rs = db.OpenRecordset("vid")
    rs.Close
    Set rs = Nothing
    main
    
End Sub

'Sort/ReSort ListView by the clicked column
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
SortListView ListView1, ColumnHeader
End Sub
