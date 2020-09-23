VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBooks 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Books Database Entry Form"
   ClientHeight    =   5325
   ClientLeft      =   1575
   ClientTop       =   1935
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   960
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   4935
      Left            =   150
      ScaleHeight     =   4875
      ScaleWidth      =   7275
      TabIndex        =   10
      Top             =   180
      Width           =   7335
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to cancel data entry"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "C&lose"
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to close window"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Book Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   6975
         Begin VB.ComboBox cboType 
            DataField       =   "Type"
            DataSource      =   "adoBooks_Details"
            Height          =   315
            ItemData        =   "frmBooks.frx":0000
            Left            =   1680
            List            =   "frmBooks.frx":0016
            TabIndex        =   2
            Top             =   1440
            Width           =   3255
         End
         Begin VB.TextBox txtCopies 
            DataField       =   "CatalogNo"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   5160
            TabIndex        =   7
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txtShelfNo 
            DataField       =   "ShelfNo"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   5160
            TabIndex        =   6
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtCatalogNo 
            DataField       =   "NoCopies"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Top             =   2520
            Width           =   2055
         End
         Begin VB.TextBox txtEditionNo 
            DataField       =   "EditionNo"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Top             =   2160
            Width           =   2055
         End
         Begin VB.TextBox txtISBN 
            DataField       =   "ISBN"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox txtAuthor 
            DataField       =   "Author"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   1680
            TabIndex        =   1
            Top             =   1080
            Width           =   5175
         End
         Begin VB.TextBox txtTitle 
            DataField       =   "Title"
            DataSource      =   "adoBooks_Details"
            Height          =   285
            Left            =   1680
            TabIndex        =   0
            Top             =   720
            Width           =   5175
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Book Code:"
            Height          =   195
            Left            =   735
            TabIndex        =   25
            Top             =   360
            Width           =   840
         End
         Begin VB.Label lblBookCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblBookCode"
            DataField       =   "BookCode"
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
            Left            =   1680
            TabIndex        =   24
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
            Height          =   195
            Left            =   1080
            TabIndex        =   22
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edition No:"
            Height          =   195
            Left            =   810
            TabIndex        =   21
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shelf No:"
            Height          =   255
            Left            =   4320
            TabIndex        =   20
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Catalog No:"
            Height          =   255
            Left            =   4200
            TabIndex        =   19
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Copies*:"
            Height          =   195
            Left            =   540
            TabIndex        =   18
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label lblDatePosted 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblDatePosted"
            DataField       =   "DatePosted"
            DataSource      =   "adoBooks_Details"
            Height          =   195
            Left            =   5280
            TabIndex        =   17
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Posted:"
            Height          =   255
            Left            =   4200
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Author*:"
            Height          =   195
            Left            =   1005
            TabIndex        =   15
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN:"
            Height          =   255
            Left            =   1080
            TabIndex        =   14
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title*:"
            Height          =   195
            Left            =   1170
            TabIndex        =   13
            Top             =   720
            Width           =   405
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   540
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
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
               Description     =   "Find Books"
               Object.ToolTipText     =   "Click to find books by criteria"
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
            Picture         =   "frmBooks.frx":007C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":0198
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":02B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":03D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":04EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":0940
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":0D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":11E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":163C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":1750
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBooks.frx":1864
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoBooks_Details 
      Height          =   330
      Left            =   2880
      Top             =   5880
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
      Connect         =   $"frmBooks.frx":1976
      OLEDBString     =   $"frmBooks.frx":1A07
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
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If MsgBox("Are you sure to cancel data entry?", vbQuestion + vbYesNo, "Confirm Cancel") = vbYes Then
        Me.adoBooks_Details.Refresh
        Call lockfields
        '
        Toolbar1.Buttons(1).Enabled = True     'new
        Toolbar1.Buttons(2).Enabled = True     'edit
        Toolbar1.Buttons(3).Enabled = False    'save
        Toolbar1.Buttons(4).Enabled = True     'delete
        Toolbar1.Buttons(6).Enabled = True     'first
        Toolbar1.Buttons(7).Enabled = True     'previous
        Toolbar1.Buttons(8).Enabled = True     'next
        Toolbar1.Buttons(9).Enabled = True     'last
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo res  'traps error caused by unfilled up required fields
        If txtTitle.Locked = True Then
            MsgBox "No changes to save.", vbInformation, "Message"
            Exit Sub
        Else
            
            If lblBookCode.Caption = "" Then
                MsgBox "No current record to save.", vbInformation, "Message"
                cmdCancel.Enabled = False
            Else
                If MsgBox("Are you sure to save the current record?", vbQuestion + vbYesNo, "Confirm Save") = vbYes Then
                    'sets fullname = LName, FName MName (initial)
                    adoBooks_Details.Recordset.Fields!Code = txtTitle.Text & "-" & lblBookCode.Caption
                    adoBooks_Details.Recordset.Update
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
    'places the form in center
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 4
    
    Call lockfields     'lock fields
    
    'displays no data in fields on form load
    On Error Resume Next
    adoBooks_Details.Refresh
    adoBooks_Details.Recordset.MoveLast
    adoBooks_Details.Recordset.MoveNext
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtTitle.Locked = False Then
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
Function lockfields()
    Me.txtAuthor.Locked = True
    Me.txtCatalogNo.Locked = True
    Me.txtCopies.Locked = True
    Me.txtEditionNo.Locked = True
    Me.txtISBN.Locked = True
    Me.txtShelfNo.Locked = True
    Me.txtTitle.Locked = True
    Me.cboType.Locked = True
End Function

Function unlockfields()
    Me.txtAuthor.Locked = False
    Me.txtCatalogNo.Locked = False
    Me.txtCopies.Locked = False
    Me.txtEditionNo.Locked = False
    Me.txtISBN.Locked = False
    Me.txtShelfNo.Locked = False
    Me.txtTitle.Locked = False
    Me.cboType.Locked = False
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'the following codes invoke the function of each button in the toolbar_
    'since there are many buttons,each button is being recognized through its index
    'and the select case statement was used
    
    Select Case Button.Index
    Case 1
        If MsgBox("Are you sure to add new record?", vbQuestion + vbYesNo, "New Record") = vbYes Then
            adoBooks_Details.Recordset.AddNew
            Call unlockfields
            txtTitle.SetFocus
            cmdCancel.Enabled = True
            
            'automatically loads system date
            lblDatePosted.Caption = Format(Date, "mm/dd/yyyy")
            'sets unique client code out of the system dates and database recordcount
            lblBookCode.Caption = Format(Date, "dmyy") & Format(Time, "hns") & "-" & Format(adoBooks_Details.Recordset.RecordCount, "0000")
            
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
        If lblBookCode.Caption = "" Then
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
        If txtTitle.Locked = True Then
            MsgBox "No changes to save.", vbInformation, "Message"
            Exit Sub
        Else
            
            If lblBookCode.Caption = "" Then
                MsgBox "No current record to save.", vbInformation, "Message"
                cmdCancel.Enabled = False
            Else
                If MsgBox("Are you sure to save the current record?", vbQuestion + vbYesNo, "Confirm Save") = vbYes Then
                    'sets fullname = LName, FName MName (initial)
                    adoBooks_Details.Recordset.Fields!Code = txtTitle.Text & "-" & lblBookCode.Caption
                    adoBooks_Details.Recordset.Update
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
        If lblBookCode.Caption = "" Then
            MsgBox "No current record to delete. Please select a record first.", vbInformation, "Message"
        Else
            If MsgBox("Are you sure to delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                adoBooks_Details.Recordset.Delete
                adoBooks_Details.Recordset.MovePrevious
                adoBooks_Details.Recordset.Update
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
        If adoBooks_Details.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
        Else
            adoBooks_Details.Recordset.MoveFirst
        End If
    Case 7
        On Error Resume Next
        If adoBooks_Details.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
            Exit Sub
        Else
            adoBooks_Details.Recordset.MovePrevious
            If adoBooks_Details.Recordset.BOF Then
                adoBooks_Details.Recordset.MoveNext
                MsgBox "Beginning of records", vbInformation, "Message"
            End If
        End If
    Case 8
        On Error Resume Next
        'sees first whether or not database is empty, if empty_
        'msgbox appears informing user that database is empty else
        'the database proceed to the next record
        
        If adoBooks_Details.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
            Exit Sub
        Else
            adoBooks_Details.Recordset.MoveNext
            If adoBooks_Details.Recordset.EOF Then
                adoBooks_Details.Recordset.MovePrevious
                MsgBox "End of records.", vbInformation, "Message"
            End If
        End If
    Case 9
        'sees first whether or not database is empty
        On Error Resume Next
        If adoBooks_Details.Recordset.RecordCount = 0 Then
            MsgBox "Database is empty.", vbInformation, "Message"
            Exit Sub
        Else
            adoBooks_Details.Recordset.MoveLast
        End If
    
    
    
    End Select

End Sub
