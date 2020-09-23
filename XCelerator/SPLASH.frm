VERSION 5.00
Begin VB.Form SPLASH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XCelerator"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "SPLASH.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "SPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    'Hide the splash screen and show the Xcelerator
    frmMain.Visible = True
    Unload Me
End Sub
