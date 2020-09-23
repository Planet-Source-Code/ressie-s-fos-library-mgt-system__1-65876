VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStudProfile 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Students/Borrowers Profile"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc adoClients_Profile 
      Height          =   330
      Left            =   2040
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   $"frmStudProfile.frx":0000
      OLEDBString     =   $"frmStudProfile.frx":0091
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   4815
      Left            =   105
      ScaleHeight     =   4755
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "C&lose"
         Height          =   495
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click to close window"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Click to cancel data entry"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Borrowers Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   165
         TabIndex        =   1
         Top             =   765
         Width           =   9375
         Begin VB.TextBox txtCourse 
            Alignment       =   2  'Center
            DataField       =   "Course"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   4920
            TabIndex        =   21
            Top             =   2520
            Width           =   4215
         End
         Begin VB.TextBox txtSchool 
            Alignment       =   2  'Center
            DataField       =   "School"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   480
            TabIndex        =   20
            Top             =   2520
            Width           =   4215
         End
         Begin VB.TextBox txtAddress 
            Alignment       =   2  'Center
            DataField       =   "Address"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   2400
            TabIndex        =   19
            Top             =   1920
            Width           =   6735
         End
         Begin VB.TextBox txtBDate 
            Alignment       =   2  'Center
            DataField       =   "BDate"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   480
            TabIndex        =   18
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtLName 
            Alignment       =   2  'Center
            DataField       =   "LName"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   6240
            TabIndex        =   17
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtMName 
            Alignment       =   2  'Center
            DataField       =   "MName"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   3360
            TabIndex        =   16
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtFName 
            Alignment       =   2  'Center
            DataField       =   "FName"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   480
            TabIndex        =   15
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtIDNo 
            Alignment       =   2  'Center
            DataField       =   "IDNo"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   7200
            TabIndex        =   14
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtDateReg 
            Alignment       =   2  'Center
            DataField       =   "DateReg"
            DataSource      =   "adoClients_Profile"
            Height          =   285
            Left            =   5160
            TabIndex        =   13
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Course/Degree/Position:"
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
            Left            =   5940
            TabIndex        =   22
            Top             =   2280
            Width           =   2130
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lib Ref No:"
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
            Left            =   1440
            TabIndex        =   10
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID No:"
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
            Left            =   7845
            TabIndex        =   9
            Top             =   480
            Width           =   585
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Registered:"
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
            Left            =   5160
            TabIndex        =   8
            Top             =   480
            Width           =   1965
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
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
            Left            =   1425
            TabIndex        =   7
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name:"
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
            Left            =   4215
            TabIndex        =   6
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Surname:"
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
            Left            =   7275
            TabIndex        =   5
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date:"
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
            Left            =   840
            TabIndex        =   4
            Top             =   1680
            Width           =   945
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
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
            Left            =   5385
            TabIndex        =   3
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School/Office:"
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
            Left            =   1980
            TabIndex        =   2
            Top             =   2280
            Width           =   1260
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   540
         Left            =   210
         TabIndex        =   11
         Top             =   120
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   953
         ButtonWidth     =   1005
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Description     =   "New"
               Object.ToolTipText     =   "Add new record"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Description     =   "Edit"
               Object.ToolTipText     =   "Edit current record"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               Description     =   "Save"
               Object.ToolTipText     =   "Save current record"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Description     =   "Delete"
               Object.ToolTipText     =   "Delete current record"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "First"
               Description     =   "First Record"
               Object.ToolTipText     =   "Go to the first record"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Description     =   "Previous Record"
               Object.ToolTipText     =   "Go to the previous record"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Next"
               Description     =   "Next Record"
               Object.ToolTipText     =   "Go to the next record"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Last"
               Description     =   "Last Record"
               Object.ToolTipText     =   "Go to the last record"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Find"
               Description     =   "Find Borrowers"
               Object.ToolTipText     =   "Click to find name of borrowers"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":0122
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":023E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":035A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":0476
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":0592
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":09E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":0E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":128E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":16E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":17F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStudProfile.frx":190A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStudProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdSave_Click()
    On Error GoTo res  'traps error caused by unfilled up required fields
        If txtFName.Locked = True Then
            MsgBox "No changes to save.", vbInformation, "Message"
            Exit Sub
        Else
            
            If lblRefNo.Caption = "" Then
                MsgBox "No current record to save.", vbInformation, "Message"
                cmdCancel.Enabled = False
            Else
                If MsgBox("Are you sure to save the current record?", vbQuestion + vbYesNo, "Confirm Save") = vbYes Then
                    'sets fullname = LName, FName MName (initial)
                    
                    adoClients_Profile.Recordset.Fields!Fullname = txtLName.Text & ", " & txtFName.Text & " " & UCase(Left(txtMName.Text, 1))
                    adoClients_Profile.Recordset.Update
                    cmdCancel.Enabled = False
                    MsgBox "Record has been successfully saved.", vbInformation, "Save Successful"
                    'lock all fields
                    Call lockfields
                    'enables and disables toolbar buttons
                    Toolbar1.Buttons(1).Enabled = True     'new
                    Toolbar1.Buttons(2).Enabled = True     'edit
                    Toolbar1.Buttons(3).Enabled = False    'save
                    Toolbar1.Buttons(4).Enabled = True     'delete
                    Toolbar1.Buttons(6).Enabled = True     'first
                    Toolbar1.Buttons(7).Enabled = True     'previous
                    Toolbar1.Buttons(8).Enabled = True     'next
                    Toolbar1.Buttons(9).Enabled = True     'last
                    
                End If
            End If
res:
            'traps error cuased by unfilling up of required fields
            If Err.Number = -2147467259 Then
                MsgBox "Please fill in all fields!", vbExclamation, "Message"
                Exit Sub
            End If
        End If
End Sub

Private Sub Form_Load()
    Call lockfields 'lock all fields
    'displays no data in fields at form load
    adoClients_Profile.Refresh
    adoClients_Profile.Recordset.MoveLast
    adoClients_Profile.Recordset.MoveNext
    
    'places the form in center
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 4
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtFName.Locked = False Then
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'the following codes invoke the function of each button in the toolbar_
    'since there are many buttons,each button is being recognized through its index
    'and the select case statement was used
    
    Select Case Button.Index
    Case 1
        If MsgBox("Are you sure to add new record?", vbQuestion + vbYesNo, "New Record") = vbYes Then
            adoClients_Profile.Recordset.AddNew
            Call unlockfields
            txtDateReg.SetFocus
            cmdCancel.Enabled = True
            
            'automatically loads system date
            txtDateReg.Text = Format(Date, "mm/dd/yyyy")
            'sets unique client code out of the system dates and database recordcount
            lblRefNo.Caption = Format(Date, "dmyy") & Format(Time, "hns") & adoClients_Profile.Recordset.RecordCount
            
            'enables and disables toolbar buttons
            Toolbar1.Buttons(1).Enabled = False     'new
            Toolbar1.Buttons(2).Enabled = False     'edit
            Toolbar1.Buttons(3).Enabled = True      'save
            Toolbar1.Buttons(4).Enabled = False     'delete
            Toolbar1.Buttons(6).Enabled = False     'first
            Toolbar1.Buttons(7).Enabled = False     'previous
            Toolbar1.Buttons(8).Enabled = False     'next
            Toolbar1.Buttons(9).Enabled = False     'last
        End If
    Case 2
        If lblRefNo.Caption = "" Then
            MsgBox "No current record to edit. Please select a record first.", vbInformation, "Message"
        Else
            If MsgBox("Are you sure to edit the current record?", vbQuestion + vbYesNo, "Confirm Edit") = vbYes Then
                'unlocks all editable fields
                Call unlockfields
                '
                'enables and disables toolbar buttons
                Toolbar1.Buttons(1).Enabled = False     'new
                Toolbar1.Buttons(2).Enabled = False     'edit
                Toolbar1.Buttons(3).Enabled = True      'save
                Toolbar1.Buttons(4).Enabled = False     'delete
                Toolbar1.Buttons(6).Enabled = False     'first
                Toolbar1.Buttons(7).Enabled = False     'previous
                Toolbar1.Buttons(8).Enabled = False     'next
                Toolbar1.Buttons(9).Enabled = False     'last
                cmdCancel.Enabled = True
                
            End If
        End If
    Case 3
        On Error GoTo res  'traps error caused by unfilled up required fields
        If txtFName.Locked = True Then
            MsgBox "No changes to save.", vbInformation, "Message"
            Exit Sub
        Else
            
            If lblRefNo.Caption = "" Then
                MsgBox "No current record to save.", vbInformation, "Message"
                cmdCancel.Enabled = False
            Else
                If MsgBox("Are you sure to save the current record?", vbQuestion + vbYesNo, "Confirm Save") = vbYes Then
                    'sets fullname = LName, FName MName (initial)
                    
                    adoClients_Profile.Recordset.Fields!Fullname = txtLName.Text & ", " & txtFName.Text & " " & UCase(Left(txtMName.Text, 1))
                    adoClients_Profile.Recordset.Update
                    cmdCancel.Enabled = False
                    MsgBox "Record has been successfully saved.", vbInformation, "Save Successful"
                    'lock all fields
                    Call lockfields
                    'enables and disables toolbar buttons
                    Toolbar1.Buttons(1).Enabled = True     'new
                    Toolbar1.Buttons(2).Enabled = True     'edit
                    Toolbar1.Buttons(3).Enabled = False    'save
                    Toolbar1.Buttons(4).Enabled = True     'delete
                    Toolbar1.Buttons(6).Enabled = True     'first
                    Toolbar1.Buttons(7).Enabled = True     'previous
                    Toolbar1.Buttons(8).Enabled = True     'next
                    Toolbar1.Buttons(9).Enabled = True     'last
                    
                End If
            End If
res:
            'traps error cuased by unfilling up of required fields
            If Err.Number = -2147467259 Then
                MsgBox "Please fill in all fields!", vbExclamation, "Message"
                Exit Sub
            End If
        End If
    Case 4
        On Error Resume Next
        If lblRefNo.Caption = "" Then
            MsgBox "No current record to delete. Please select a record first.", vbInformation, "Message"
        Else
            If MsgBox("Are you sure to delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                adoClients_Profile.Recordset.Delete
                adoClients_Profile.Recordset.MovePrevious
                adoClients_Profile.Recordset.Update
                MsgBox "The current record has been deleted.", vbInformation, "Delete Successful"
                
                'enables and disables toolbar buttons
                Toolbar1.Buttons(1).Enabled = True     'new
                Toolbar1.Buttons(2).Enabled = True     'edit
                Toolbar1.Buttons(3).Enabled = False    'save
                Toolbar1.Buttons(4).Enabled = True     'delete
                Toolbar1.Buttons(6).Enabled = True     'first
                Toolbar1.Buttons(7).Enabled = True     'previous
                Toolbar1.Buttons(8).Enabled = True     'next
                Toolbar1.Buttons(9).Enabled = True     'last
            End If
        End If
    Case 5
        'separator
    'database browser (arrow) buttons
    Case 6
        On Error Resume Next
        If adoClients_Profile.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
        Else
            adoClients_Profile.Recordset.MoveFirst
        End If
    Case 7
        On Error Resume Next
        If adoClients_Profile.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
            Exit Sub
        Else
            adoClients_Profile.Recordset.MovePrevious
            If adoClients_Profile.Recordset.BOF Then
                adoClients_Profile.Recordset.MoveNext
                MsgBox "Beginning of records", vbInformation, "Message"
            End If
        End If
    Case 8
        On Error Resume Next
        'sees first whether or not database is empty, if empty_
        'msgbox appears informing user that database is empty else
        'the database proceed to the next record
        
        If adoClients_Profile.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
            Exit Sub
        Else
            adoClients_Profile.Recordset.MoveNext
            If adoClients_Profile.Recordset.EOF Then
                adoClients_Profile.Recordset.MovePrevious
                MsgBox "End of records.", vbInformation, "Message"
            End If
        End If
    Case 9
        'sees first whether or not database is empty
        On Error Resume Next
        If adoClients_Profile.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
            Exit Sub
        Else
            adoClients_Profile.Recordset.MoveLast
        End If
    End Select

End Sub
Function lockfields()
    Me.txtAddress.Locked = True
    Me.txtBDate.Locked = True
    Me.txtCourse.Locked = True
    Me.txtDateReg.Locked = True
    Me.txtFName.Locked = True
    Me.txtIDNo.Locked = True
    Me.txtLName.Locked = True
    Me.txtMName.Locked = True
    Me.txtSchool.Locked = True
    
End Function

Function unlockfields()
    Me.txtAddress.Locked = False
    Me.txtBDate.Locked = False
    Me.txtCourse.Locked = False
    Me.txtDateReg.Locked = False
    Me.txtFName.Locked = False
    Me.txtIDNo.Locked = False
    Me.txtLName.Locked = False
    Me.txtMName.Locked = False
    Me.txtSchool.Locked = False
    
End Function
