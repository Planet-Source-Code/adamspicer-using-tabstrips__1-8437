VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   Caption         =   "Tabstrip Example"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExitCMD 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is Tab #1...  Put whatever you desire here!"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   3300
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5280
   End
   Begin VB.Image Image2 
      Height          =   3300
      Left            =   0
      Picture         =   "Form1.frx":33462
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5280
   End
   Begin MSForms.TabStrip TabStrip 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      ListIndex       =   0
      Size            =   "9340;6376"
      Items           =   "Tab1;Tab2;"
      TipStrings      =   ";;"
      Names           =   "Tab1;Tab2;"
      NewVersion      =   -1  'True
      TabsAllocated   =   2
      Tags            =   ";;"
      TabData         =   2
      Accelerator     =   ";;"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      TabState        =   "3;3"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(_.,._)  One Computer Software  (_.,._)'
'(_.,._)          2000           (_.,._)'

Private Sub ExitCMD_Click()
End

End Sub

Private Sub Form_Load()
'changes the tabs name
TabStrip.Tabs(0).Caption = "Adam's Tab #1"
TabStrip.Tabs(1).Caption = "Adam's Tab #2"
'======================
Image2.Visible = False

End Sub

Private Sub TabStrip_Click(Index As Long)
'this changes/hides things for the second tab!
If TabStrip.SelectedItem.Caption = "Adam's Tab #2" Then
    Label1.Caption = "This is Tab #2...  Put whatever you desire here also!"
    Image1.Visible = False
    Image2.Visible = True
End If
'=============================
'this changes/hides things for the first tab!
If TabStrip.SelectedItem.Caption = "Adam's Tab #1" Then
    Label1.Caption = "This is Tab #1...  Put whatever you desire!"
    Image1.Visible = True
    Image2.Visible = False
End If

End Sub
