Attribute VB_Name = "modData"
Option Explicit
'Common variables used in this program
Public UpDated  As Boolean
Public db       As Database
Public rs       As Recordset
Public LVName   As ListView
Public Itm      As ListItem
Public Colx     As ColumnHeader

' Though the code is incomplete, all of the components
' you need to open, read, write, and update an .mdb
' are present.               Sly June 22, 2002

Sub main()
    SetHeads                'modListViewControl
    ClearForm               'modListViewControl
    ListLoad                'modListViewControl
    FrmData.Show
End Sub
