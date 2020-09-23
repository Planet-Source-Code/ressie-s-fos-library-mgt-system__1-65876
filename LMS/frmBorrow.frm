VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBorrow 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Logbook - Borrowing Books"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoqrySTatus 
      Height          =   330
      Left            =   9360
      Top             =   3120
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":0000
      OLEDBString     =   $"frmBorrow.frx":0091
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "qryStatus"
      Caption         =   "adoqrySTatus"
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
   Begin MSAdodcLib.Adodc adoqryListBooks 
      Height          =   330
      Left            =   9360
      Top             =   2835
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":0122
      OLEDBString     =   $"frmBorrow.frx":01B3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "qryListBooks"
      Caption         =   "adoqryListBooks"
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
   Begin MSAdodcLib.Adodc adoTransact 
      Height          =   330
      Left            =   9345
      Top             =   1395
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":0244
      OLEDBString     =   $"frmBorrow.frx":02D5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Transaction_Borrow"
      Caption         =   "adoTransact"
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
   Begin MSAdodcLib.Adodc adoLName 
      Height          =   330
      Left            =   9330
      Top             =   1170
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Connect         =   $"frmBorrow.frx":0366
      OLEDBString     =   $"frmBorrow.frx":03F7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "LName"
      Caption         =   "adoLName"
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
   Begin MSAdodcLib.Adodc adoClients_Profile 
      Height          =   330
      Left            =   9315
      Top             =   750
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":0488
      OLEDBString     =   $"frmBorrow.frx":0519
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Clients_Profile"
      Caption         =   "adoClients_Profile"
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
   Begin MSAdodcLib.Adodc adoqryBooks 
      Height          =   330
      Left            =   9330
      Top             =   435
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":05AA
      OLEDBString     =   $"frmBorrow.frx":063B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "qryBooks"
      Caption         =   "adoqryBooks"
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
   Begin MSAdodcLib.Adodc adoBooks_Details 
      Height          =   330
      Left            =   9330
      Top             =   75
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":06CC
      OLEDBString     =   $"frmBorrow.frx":075D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Book_Details"
      Caption         =   "adoBooks_Details"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   6330
      Left            =   120
      ScaleHeight     =   6270
      ScaleWidth      =   8835
      TabIndex        =   11
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ca&ncel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to cancel data entry"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to delete transaction details"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Click to edit transaction details"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&New"
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to make new transaction"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click to view all borrowed books"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click to save transaction details"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Transaction Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   8415
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFC0C0&
            Height          =   1815
            Left            =   240
            ScaleHeight     =   1755
            ScaleWidth      =   7875
            TabIndex        =   18
            Top             =   2220
            Visible         =   0   'False
            Width           =   7935
            Begin VB.TextBox txtNoCopies 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6240
               MaxLength       =   1
               TabIndex        =   4
               Top             =   1200
               Width           =   975
            End
            Begin VB.ComboBox cboTitle 
               Height          =   315
               Left            =   2160
               Locked          =   -1  'True
               Sorted          =   -1  'True
               TabIndex        =   3
               Top             =   75
               Width           =   5535
            End
            Begin VB.Label lblCN 
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "lblCN"
               DataField       =   "CN"
               DataSource      =   "adoTransact"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   32
               Top             =   1380
               Width           =   2775
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CN:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   31
               Top             =   1380
               Width           =   330
            End
            Begin VB.Shape Shape1 
               Height          =   975
               Left            =   4200
               Top             =   675
               Width           =   3495
            End
            Begin VB.Label lblReturnDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "lblReturnDate"
               DataSource      =   "adoBooks_Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   30
               Top             =   960
               Width           =   1185
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Return Date:"
               Height          =   255
               Left            =   5160
               TabIndex        =   29
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblDateBorrowed 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "lblDateBorrowed"
               DataSource      =   "adoBooks_Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   28
               Top             =   720
               Width           =   1410
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Borrowed:"
               Height          =   255
               Left            =   4920
               TabIndex        =   27
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. of Copies Borrowed*:"
               Height          =   195
               Left            =   4350
               TabIndex        =   26
               Top             =   1200
               Width           =   1785
            End
            Begin VB.Label lblTotalCopies 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "lblTotalCopies"
               DataField       =   "NoCopies"
               DataSource      =   "adoqryBooks"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2160
               TabIndex        =   25
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total No. of Copies:"
               Height          =   255
               Left            =   600
               TabIndex        =   24
               Top             =   705
               Width           =   1455
            End
            Begin VB.Label lblAvailablCopies 
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "lblAvailablCopies"
               DataField       =   "Available"
               DataSource      =   "adoqryBooks"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   23
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblAuthor 
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "lblAuthor"
               DataField       =   "Author"
               DataSource      =   "adoqryBooks"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   22
               Top             =   435
               Width           =   5535
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Book Title/Book Code*:"
               Height          =   195
               Left            =   345
               TabIndex        =   21
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Author:"
               Height          =   195
               Left            =   1530
               TabIndex        =   20
               Top             =   420
               Width           =   510
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Available Copies:"
               Height          =   195
               Left            =   825
               TabIndex        =   19
               Top             =   960
               Width           =   1215
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00FF8080&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   315
               Left            =   105
               Top             =   1335
               Width           =   3975
            End
         End
         Begin MSComctlLib.ListView lstBooksBorrowed 
            Height          =   3255
            Left            =   240
            TabIndex        =   2
            Top             =   720
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "BookCode"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Title"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Author"
               Object.Width           =   4233
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Date Borrowed"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "# Copy"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Status"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblNoUnreturned 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   36
            Top             =   4080
            Width           =   1395
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Unreturned Books:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   4080
            Width           =   2535
         End
         Begin VB.Label lblFullname 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Fullname"
            DataSource      =   "adoClients_Profile"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4440
            TabIndex        =   17
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lblRefNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            DataField       =   "RefNo"
            DataSource      =   "adoClients_Profile"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2160
            TabIndex        =   16
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Borrower:"
            Height          =   195
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Look up Borrower"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Width           =   6855
         Begin VB.CommandButton cmdFind 
            BackColor       =   &H00FFC0C0&
            Height          =   495
            Left            =   5640
            Picture         =   "frmBorrow.frx":07EE
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Click to start searching names"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   2160
            TabIndex        =   0
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter surname of Borrower:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1920
         End
      End
   End
   Begin MSAdodcLib.Adodc adoRefNo 
      Height          =   330
      Left            =   9345
      Top             =   1755
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Connect         =   $"frmBorrow.frx":08F0
      OLEDBString     =   $"frmBorrow.frx":0981
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "RefNo"
      Caption         =   "adoRefNo"
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
   Begin VB.Label lblBookCode1 
      Caption         =   "lblBookCode1"
      DataField       =   "BookCode"
      DataSource      =   "adoqryListBooks"
      Height          =   255
      Left            =   9480
      TabIndex        =   37
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblBookCode 
      Caption         =   "lblBookCode"
      DataField       =   "BookCode"
      DataSource      =   "adoqryBooks"
      Height          =   255
      Left            =   9480
      TabIndex        =   34
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Label lblTitle 
      Caption         =   "lblTitle"
      DataField       =   "Title"
      DataSource      =   "adoqryBooks"
      Height          =   255
      Left            =   9495
      TabIndex        =   33
      Top             =   2190
      Width           =   1155
   End
End
Attribute VB_Name = "frmBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboTitle_Click()
    adoqryBooks.Refresh
    adoqryBooks.Recordset.Find "Code = '" & cboTitle.Text & "'"
    '=======
    
    On Error Resume Next
    adoqryListBooks.Refresh
    adoqryListBooks.Recordset.Find "BookCode ='" & lblBookCode.Caption & "'"
    
    If Not lblBookCode1.Caption = "" Then
        MsgBox "Transaction not allowed! The borrower has still unreturned book/s of this kind.", vbCritical, "Warning!"
        cboTitle.Text = ""
        cboTitle.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure to cancel data entry?", vbQuestion + vbYesNo, "Confirm Cancel") = vbYes Then
        lstBooksBorrowed.Height = 3255
        Picture2.Visible = False
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        adoTransact.Refresh
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Function loadbooks()
    Me.adoBooks_Details.Refresh
    Me.cboTitle.Clear
    Do While Not Me.adoBooks_Details.Recordset.EOF
        Me.cboTitle.AddItem (Me.adoBooks_Details.Recordset.Fields!Title & "-" & Me.adoBooks_Details.Recordset.Fields!BookCode)
        Me.adoBooks_Details.Recordset.MoveNext
    Loop
    Me.adoBooks_Details.Refresh
End Function

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure to delete the selected record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        adoTransact.Refresh
        adoTransact.Recordset.Find "CN='" & lblCN.Caption & "'"
        adoTransact.Recordset.Delete
        adoTransact.Recordset.Update
        adoTransact.Recordset.MovePrevious
        
        '================
        Dim n As Integer
        Dim m As Integer
        adoqryBooks.Refresh
        adoqryBooks.Recordset.Find "BookCode ='" & lstBooksBorrowed.SelectedItem & "'"
        n = adoqryBooks.Recordset.Fields!NoBorrowed
        m = lstBooksBorrowed.SelectedItem.SubItems(4)
        adoqryBooks.Recordset.Fields!NoBorrowed = n - m
        adoqryBooks.Recordset.Update
        adoqryBooks.Refresh
        '=================
        
        MsgBox "Selected record has been successfully deleted.", vbInformation, "Delete Successfull"
        Call listbooks
        On Error Resume Next
        Dim o As Integer
        frmBorrow.adoqrySTatus.Refresh
        o = frmBorrow.adoqrySTatus.Recordset.Fields!SumofNoCopyBorrowed
        frmBorrow.lblNoUnreturned.Caption = o
        
     Else
         Exit Sub
     End If
    
End Sub

Private Sub cmdFind_Click()
    If txtFind.Text = "" Then
        MsgBox "Please enter borrower's Last Name first.", vbExclamation, "Message"
        txtFind.SetFocus
    Else
        '=====================================
        'updates last name for query purposes
        adoLName.Refresh
        adoLName.Recordset.Fields!LName = txtFind.Text
        adoLName.Recordset.Update
        adoLName.Refresh
        '=====================================
        Me.adoClients_Profile.Refresh
        Call loadname
        
        'gives message if client's surname is not found
        If frmSearchName.lstSearchResult.ListItems.Count = 0 Then
            MsgBox "No record found matching your search criteria!", vbInformation, "Message"
            txtFind.SetFocus
            SendKeys "{home}+{end}"
        Else
            frmSearchName.Show vbModal
        End If
    End If
    
End Sub
Private Sub cmdNew_Click()
        
    If lblRefNo.Caption = "" Then
           MsgBox "Please select name of borrower first.", vbExclamation, "Message"
           Exit Sub
           txtFind.SetFocus
    Else
        If Me.lblNoUnreturned.Caption > 1 Then
            MsgBox "Borrower is allowed to borrow up to 2 books only.", vbCritical, "Warning!"
            Exit Sub
        Else
            If MsgBox("Are you sure to make new borrowing transaction?", vbQuestion + vbYesNo, "New Borrowing") = vbYes Then
                Picture2.Visible = True
                cboTitle.SetFocus
                cmdSave.Enabled = True
                cboTitle.Locked = False
                txtNoCopies.Locked = False
                lstBooksBorrowed.Height = 1575
                cmdNew.Enabled = False
                cmdCancel.Enabled = True
                cmdEdit.Enabled = False
                
                'loads system date and return date_
                'return date is set at 5 days after date borrowed
                lblDateBorrowed.Caption = Format(Date, "mm/dd/yyyy")
                lblReturnDate.Caption = DateAdd("d", 5, lblDateBorrowed.Caption)
                
                
                
                'prepares transaction table to store new records
                adoTransact.Refresh
                adoTransact.Recordset.AddNew
                
                'sets CN
                lblCN.Caption = Format(Date, "dmyy") & Format(Time, "hns") & "-" & Format(adoTransact.Recordset.RecordCount, "0000")
                
                
            Else
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo res
    If cboTitle.Text = "" Or txtNoCopies.Text = "" Then
        MsgBox "All fields must be filled in.", vbExclamation, "Message"
        Exit Sub
    Else
            
        If MsgBox("Are you sure to save the information?", vbQuestion + vbYesNo, "Confirm Save") = vbYes Then
            adoTransact.Recordset.Fields!RefNo = lblRefNo.Caption
            adoTransact.Recordset.Fields!Fullname = lblFullname.Caption
            adoTransact.Recordset.Fields!Title = lblTitle.Caption
            adoTransact.Recordset.Fields!Author = lblAuthor.Caption
            adoTransact.Recordset.Fields!DateBorrowed = lblDateBorrowed.Caption
            adoTransact.Recordset.Fields!ReturnDate = lblReturnDate.Caption
            adoTransact.Recordset.Fields!NoCopyBorrowed = txtNoCopies.Text
            adoTransact.Recordset.Fields!CN = lblCN.Caption
            adoTransact.Recordset.Fields!BookCode = lblBookCode.Caption
            'update ado
            adoTransact.Recordset.Update
            '===========
            adoqryBooks.Recordset.Fields!NoBorrowed = adoqryBooks.Recordset.Fields!NoBorrowed + txtNoCopies.Text
            adoqryBooks.Recordset.Update
            adoqryBooks.Refresh
            Call cboTitle_Click
            '==========
            cboTitle.Locked = True
            txtNoCopies.Locked = True
            cmdSave.Enabled = False
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            
            '
            lstBooksBorrowed.Height = 3255
            Picture2.Visible = False
            cmdNew.Enabled = True
            
            
            MsgBox "Information has been successfully saved.", vbInformation, "Save Successful"
            
            'load list of books borrowed
            Call listbooks
            
            'updates no of unreturned
            Dim n As Integer
            frmBorrow.adoqrySTatus.Refresh
            n = frmBorrow.adoqrySTatus.Recordset.Fields!SumofNoCopyBorrowed
            frmBorrow.lblNoUnreturned.Caption = n
            
            
        Else
            Exit Sub
        End If
        Me.Refresh
res:
    'traps error cuased by unfilling up of required fields
    If Err.Number = -2147467259 Then
        MsgBox "Please fill in all fields!", vbExclamation, "Message"
        Exit Sub
    End If
    End If
End Sub

Private Sub Form_Load()
    Call loadbooks
    lblRefNo.Caption = ""
    lblFullname.Caption = ""
    
    On Error Resume Next
    adoqryBooks.Refresh
    adoqryBooks.Recordset.MoveLast
    adoqryBooks.Recordset.MoveNext
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If cboTitle.Locked = False Then
        Dim msg
        msg = MsgBox("Closing will lost unsaved information. Save first before closing?", vbExclamation + vbYesNoCancel, "Warning!")
        
        If msg = vbYes Then
            Call cmdSave_Click 'cmdsave is hidden and with the same function as the save button
            Exit Sub
        Else
           If msg = vbNo Then
                Unload Me
                Exit Sub
           Else
                Cancel = True
           End If
        End If
    Else
        If MsgBox("Are you sure to close window?", vbQuestion + vbYesNo, "Close Window") = vbYes Then
            Unload Me
        Else
            Cancel = True
        End If
    End If

End Sub

Private Sub lstBooksBorrowed_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdDelete.Enabled = True
    cmdEdit.Enabled = True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdFind_Click
    Else
        Exit Sub
    End If
End Sub

Private Sub txtNoCopies_LostFocus()
    If Not txtNoCopies.Text = "" Then
       If txtNoCopies.Text > 2 Then
            MsgBox "Borrower is allowed to borrow up to 2 books only.", vbCritical, "Warning!"
            txtNoCopies.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
       End If
       '==========================
        Dim l As Integer
        Dim m As Integer
        Dim n As Integer
        l = lblNoUnreturned.Caption
        m = txtNoCopies.Text
        n = l + m
        If n > 2 Then
            MsgBox "Borrower is allowed to borrow up to 2 books only.", vbCritical, "Warning!"
            txtNoCopies.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        '
        If txtNoCopies.Text = "0" Then
            MsgBox "Invalid value! Default value is 1.", vbCritical, "Warning!"
            txtNoCopies.Text = "1"
            Exit Sub
        End If
        '
        If txtNoCopies.Text > Me.lblAvailablCopies.Caption Then
            MsgBox "There is only " & Me.lblAvailablCopies.Caption & " copy/ies in the stock.", vbCritical, "Warning!"
            txtNoCopies.Text = Me.lblAvailablCopies.Caption
            txtNoCopies.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
    Else
        Exit Sub
    End If
End Sub
