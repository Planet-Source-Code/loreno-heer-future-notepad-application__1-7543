VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Form2"
   ClientHeight    =   450
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   450
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Menu SysMenu 
      Caption         =   "SysMenu"
      Begin VB.Menu Maximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu Minimieren 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu About 
         Caption         =   "A&bout"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu Schliessen 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Maximieren_Click()
info.AXButtonCtl2_Click
End Sub

Private Sub Calc_Click()
Shell "calc.exe"
End Sub

Private Sub About_Click()
frmAbout.Show 1, Me
End Sub

Private Sub Maximize_Click()
info.AXButtonCtl2_Click
End Sub

Private Sub Minimieren_Click()
info.AXButtonCtl3_Click
End Sub

Private Sub Schliessen_Click()
Unload info
Unload Me
End Sub
