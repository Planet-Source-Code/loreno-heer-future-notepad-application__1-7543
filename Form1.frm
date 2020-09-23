VERSION 5.00
Object = "{02E51CD9-1709-11D4-88D8-444553540001}#1.0#0"; "FLATBUTTON.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Edit RTF"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin FlatButton.AXButtonCtl AXButtonCtl3 
      Height          =   555
      Left            =   3165
      TabIndex        =   4
      Top             =   3600
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   979
      URLPicture      =   ""
      Picture         =   "Form1.frx":0000
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl2 
      Height          =   555
      Left            =   1965
      TabIndex        =   3
      Top             =   3600
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   979
      URLPicture      =   ""
      Picture         =   "Form1.frx":2146
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl1 
      Height          =   555
      Left            =   765
      TabIndex        =   2
      Top             =   3600
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   979
      URLPicture      =   ""
      Picture         =   "Form1.frx":428C
   End
   Begin VB.Frame Frame1 
      Caption         =   "RTF Code"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AXButtonCtl1_Click()
info.RichTextBox1.TextRTF = Text1.Text
End Sub
Private Sub AXButtonCtl2_Click()
Unload Me
End Sub
Private Sub AXButtonCtl3_Click()
Text1.Text = info.RichTextBox1.TextRTF
End Sub
Private Sub Form_Load()
Text1.Text = info.RichTextBox1.TextRTF
End Sub
