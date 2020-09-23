VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{02E51CD9-1709-11D4-88D8-444553540001}#1.0#0"; "FLATBUTTON.OCX"
Begin VB.Form info 
   BorderStyle     =   0  'Kein
   Caption         =   "F@t_F|sh's Notepad"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   ScaleHeight     =   2535
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows-Standard
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"info.frx":0000
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl4 
      Height          =   210
      Left            =   4920
      TabIndex        =   3
      Top             =   25
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      URLPicture      =   ""
      Picture         =   "info.frx":00AE
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl3 
      Height          =   210
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Minimize"
      Top             =   25
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      URLPicture      =   ""
      Picture         =   "info.frx":0368
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl2 
      Height          =   210
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Maximize"
      Top             =   25
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      URLPicture      =   ""
      Picture         =   "info.frx":0622
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl1 
      Height          =   210
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   25
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      URLPicture      =   ""
      Picture         =   "info.frx":08DC
   End
   Begin FlatButton.AXButtonCtl AXButtonCtl5 
      Height          =   210
      Left            =   5520
      TabIndex        =   4
      ToolTipText     =   "Normalize"
      Top             =   25
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      URLPicture      =   ""
      Picture         =   "info.frx":0B96
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2415
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":0E50
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":0F62
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1074
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1186
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1298
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":13AA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":14BC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":15CE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":16E0
            Key             =   "Strike Through"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":17F2
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1904
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1A16
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1B28
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1C3A
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1D4C
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1E5E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "info.frx":1F70
            Key             =   "Macro"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   420
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   741
      BandCount       =   1
      _CBWidth        =   5415
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   1440
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
         ButtonWidth     =   609
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Neu"
               Object.ToolTipText     =   "Neu"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Öffnen"
               Object.ToolTipText     =   "Öffnen"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Speichern"
               Object.ToolTipText     =   "Speichern"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Drucken"
               Object.ToolTipText     =   "Drucken"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Kopieren"
               Object.ToolTipText     =   "Kopieren"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ausschneiden"
               Object.ToolTipText     =   "Ausschneiden"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Einfügen"
               Object.ToolTipText     =   "Einfügen"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Löschen"
               Object.ToolTipText     =   "Löschen"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Durchgestrichen"
               Object.ToolTipText     =   "Durchgestrichen"
               ImageKey        =   "Strike Through"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fett"
               Object.ToolTipText     =   "Fett"
               ImageKey        =   "Bold"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Kursiv"
               Object.ToolTipText     =   "Kursiv"
               ImageKey        =   "Italic"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Unterstrichen"
               Object.ToolTipText     =   "Unterstrichen"
               ImageKey        =   "Underline"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Links ausrichten"
               Object.ToolTipText     =   "Links ausrichten"
               ImageKey        =   "Align Left"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zentrieren"
               Object.ToolTipText     =   "Zentrieren"
               ImageKey        =   "Center"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rechts ausrichten"
               Object.ToolTipText     =   "Rechts ausrichten"
               ImageKey        =   "Align Right"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Suchen"
               Object.ToolTipText     =   "Suchen"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Makro"
               Object.ToolTipText     =   "Makro"
               ImageKey        =   "Macro"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F@t_F|sh's Notepad"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   15
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   25
      Picture         =   "info.frx":2082
      Stretch         =   -1  'True
      Top             =   25
      Width           =   210
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6000
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   0
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   360
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   4800
      X2              =   4800
      Y1              =   240
      Y2              =   0
   End
End
Attribute VB_Name = "info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const HTBOTTOMRIGHT = 17
Private Const HTBOTTOMLEFT = 16
Private Const HTTOPLEFT = 13
Private Const HTTOPRIGHT = 14
Private Const HTBOTTOM = 15
Private Const HTLEFT = 10
Private Const HTRIGHT = 11
Private Const HTTOP = 12
Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Neu"
                RichTextBox1.Text = ""
        Case "Öffnen"
            CD1.ShowOpen
            RichTextBox1.LoadFile (CD1.FileName)
        Case "Speichern"
            CD1.ShowSave
            RichTextBox1.SaveFile (CD1.FileName)
        Case "Drucken"
            
        Case "Kopieren"
        
        Case "Ausschneiden"
            
        Case "Einfügen"
            
        Case "Löschen"
            
        Case "Durchgestrichen"
            
        Case "Fett"
            
        Case "Kursiv"
           
        Case "Unterstrichen"
            
        Case "Links ausrichten"
            
        Case "Zentrieren"
           
        Case "Rechts ausrichten"
            
        Case "Suchen"
            
        Case "Makro"
            Form1.Show
    End Select
End Sub

Private Sub AXButtonCtl1_Click()
Unload Me
Unload Menu
End Sub
Private Sub AXButtonCtl5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub
Private Sub AXButtonCtl1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub

Private Sub AXButtonCtl2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub

Private Sub AXButtonCtl3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub

Private Sub AXButtonCtl4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub

Private Sub AXButtonCtl5_Click()
If Me.WindowState = vbMaximized Then
    Me.WindowState = 0
    AXButtonCtl5.Visible = False
    AXButtonCtl2.Visible = True
End If
End Sub
Public Sub AXButtonCtl2_Click()
If Me.WindowState = 0 Then
    Me.WindowState = vbMaximized
    AXButtonCtl5.Visible = True
    AXButtonCtl2.Visible = False
End If
End Sub
Public Sub AXButtonCtl3_Click()
Me.WindowState = vbMinimized
End Sub



Private Sub Form_Load()
If Me.Width < 1680 Then
    Me.Width = 1680
End If
If Me.Height < 300 Then
    Me.Height = 300
End If
Line4.X1 = 0
Line4.X2 = Me.Width
Line4.Y1 = 0
Line4.Y2 = 0
Line1.Y1 = 0
Line1.Y2 = Me.Height
Line3.Y1 = Me.Height
Line3.Y2 = Me.Height
Line3.X1 = 0
Line3.X2 = Me.Width
Line2.X1 = Me.Width
Line2.X2 = Me.Width
Line2.Y1 = Me.Height
Line2.Y2 = Me.Height
AXButtonCtl1.Left = Me.Width - 270
AXButtonCtl2.Left = AXButtonCtl1.Left - 240
AXButtonCtl5.Left = AXButtonCtl1.Left - 240
AXButtonCtl3.Left = AXButtonCtl5.Left - 240
AXButtonCtl4.Left = AXButtonCtl3.Left - 360
Label1.Left = Line6.X1 + 25
Label1.Width = Line5.X2 - Line5.X1 - 25
CoolBar1.Left = 0
CoolBar1.Width = Me.Width
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.WindowState = 0 Then
If x > -1 And x < 50 And y > -1 And y < 50 Then
    Me.MousePointer = 8
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTTOPLEFT, 0&
    End If
ElseIf x > Me.Width - 50 And x < Me.Width + 1 And y > -1 And y < 50 Then
    Me.MousePointer = 6
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTTOPRIGHT, 0&
    End If
ElseIf x > Me.Width - 50 And x < Me.Width + 1 And y > Me.Height - 50 And y < Me.Height + 1 Then
    Me.MousePointer = 8
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
    End If
ElseIf x > -1 And x < 50 And y > Me.Height - 50 And y < Me.Height + 1 Then
    Me.MousePointer = 6
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0&
    End If
ElseIf x > -1 And x < Me.Width + 1 And y > -1 And y < 50 Then
    Me.MousePointer = 7
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTTOP, 0&
    End If
ElseIf x > -1 And x < Me.Width + 1 And y > Me.Height - 50 And y < Me.Height + 1 Then
    Me.MousePointer = 7
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0&
    End If
ElseIf x > -1 And x < 50 And y > -1 And y < Me.Height + 1 Then
    Me.MousePointer = 9
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
    End If
ElseIf x > Me.Width - 50 And x < Me.Width + 1 And y > -1 And y < Me.Height + 1 Then
    Me.MousePointer = 9
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
    End If
Else
    Me.MousePointer = 0
End If
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
If Me.Width < 6700 Then
    Me.Width = 6700
End If
If Me.Height < 1000 Then
    Me.Height = 1000
End If
AXButtonCtl1.Left = Me.Width - 270
AXButtonCtl2.Left = AXButtonCtl1.Left - 240
AXButtonCtl5.Left = AXButtonCtl1.Left - 240
AXButtonCtl3.Left = AXButtonCtl5.Left - 240
AXButtonCtl4.Left = AXButtonCtl3.Left - 360
Line4.X1 = 0
Line4.X2 = Me.Width
Line4.Y1 = 5
Line4.Y2 = 5
Line1.Y1 = 0
Line1.Y2 = Me.Height
Line1.X1 = 5
Line1.X2 = 5
Line3.Y1 = Me.Height - 20
Line3.Y2 = Me.Height - 20
Line3.X1 = 0
Line3.X2 = Me.Width
Line2.X1 = Me.Width - 20
Line2.X2 = Me.Width - 20
Line2.Y1 = 0
Line2.Y2 = Me.Height
Line6.X1 = Image1.Left + Image1.Width + 125
Line6.X2 = Image1.Left + Image1.Width + 125
Line6.Y1 = 0
Line6.Y2 = 240
Line5.X1 = Image1.Left + Image1.Width + 125
Line5.Y1 = 240
Line5.X2 = AXButtonCtl4.Left - 120
Line5.Y2 = 240
Line7.X1 = AXButtonCtl4.Left - 120
Line7.X2 = AXButtonCtl4.Left - 120
Line7.Y1 = 0
Line7.Y2 = 240
Label1.Left = Line6.X1 + 25
Label1.Width = Line5.X2 - Line5.X1 - 25
CoolBar1.Left = 100
CoolBar1.Width = Me.Width - 200
Toolbar1.Width = CoolBar1.Width - 120
RichTextBox1.Left = 100
RichTextBox1.Width = Me.Width - 200
RichTextBox1.Top = 900
RichTextBox1.Height = Me.Height - 1000
End If
End Sub

Private Sub Image1_DblClick()
AXButtonCtl1_Click
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
Me.PopupMenu Menu.SysMenu, , 0, 240, Menu.Schliessen
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub

Private Sub Label1_DblClick()
If Me.WindowState = vbMaximized Then
    AXButtonCtl5_Click
ElseIf Me.WindowState = 0 Then
    AXButtonCtl2_Click
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 0
End Sub
