VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSearchName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Result"
   ClientHeight    =   7500
   ClientLeft      =   8700
   ClientTop       =   915
   ClientWidth     =   3075
   Icon            =   "frmSearchName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   3075
   Begin MSAdodcLib.Adodc adoqry_Clients 
      Height          =   330
      Left            =   360
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmSearchName.frx":0442
      OLEDBString     =   $"frmSearchName.frx":04D3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "qryBorrowers"
      Caption         =   "adoqry_Clients"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstSearchResult 
      Height          =   7500
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   13229
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fullname"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ClientID"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmSearchName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call loadname
End Sub
Private Sub lstSearchResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstSearchResult.SortKey = ColumnHeader.Index - 1
    lstSearchResult.Sorted = True
End Sub

Private Sub lstSearchResult_ItemClick(ByVal Item As MSComctlLib.ListItem)
        frmBorrow.adoClients_Profile.Refresh
        frmBorrow.adoClients_Profile.Recordset.MoveFirst
        frmBorrow.adoClients_Profile.Recordset.Find "RefNo = '" & lstSearchResult.SelectedItem.SubItems(1) & "'"
        
        Call listbooks
        On Error Resume Next
        Dim n As Integer
        frmBorrow.adoqrySTatus.Refresh
        n = frmBorrow.adoqrySTatus.Recordset.Fields!SumofNoCopyBorrowed
        frmBorrow.lblNoUnreturned.Caption = n
End Sub
Private Sub lstSearchResult_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then
        Unload Me
        frmBorrow.cmdNew.SetFocus
     End If
End Sub
