VERSION 5.00
Begin VB.Form frmAboutMe 
   BorderStyle     =   0  'None
   Caption         =   "About the Programmer"
   ClientHeight    =   3435
   ClientLeft      =   1005
   ClientTop       =   2100
   ClientWidth     =   5460
   ControlBox      =   0   'False
   Icon            =   "frmAboutMe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAboutMe.frx":0CCE
   ScaleHeight     =   3435
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2145
      Left            =   1560
      Picture         =   "frmAboutMe.frx":3E178
      Top             =   360
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "frmAboutMe.frx":4D782
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click to see me!"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutMe.frx":4D8D4
      ForeColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   3420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "email: ressiefos@yahoo.com"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2220
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:  09103354370; 09173736879"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "frmAboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    Image1.Visible = False
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label3_Click()
    Image1.Visible = True
End Sub

Private Sub lblDescription_Click()
    Unload Me
End Sub
