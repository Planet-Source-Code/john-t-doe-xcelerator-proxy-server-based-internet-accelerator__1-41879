VERSION 5.00
Begin VB.Form HELP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Adapting Your Browser"
      Height          =   2535
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton Command8 
         Caption         =   "Other"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Opera"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Internet Explorer"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   $"options.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Common Problems"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "When to use Xcelerator"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Optimizing Xcelerator for your connection"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adapting Your Browser"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Common Problems"
      Height          =   2535
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   4455
      Begin VB.Label Label5 
         Caption         =   $"options.frx":00B0
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "When To Use Xcelerator"
      Height          =   2535
      Left            =   1800
      TabIndex        =   12
      Top             =   960
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   $"options.frx":01CA
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Optimizing Xcelerator:"
      Height          =   2535
      Left            =   1800
      TabIndex        =   10
      Top             =   960
      Width           =   4455
      Begin VB.Label Label3 
         Caption         =   $"options.frx":030F
         Height          =   2175
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Xcelerator"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "HELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.ZOrder 0
End Sub

Private Sub Command2_Click()
Frame2.ZOrder 0
End Sub

Private Sub Command3_Click()
Frame3.ZOrder 0
End Sub

Private Sub Command4_Click()
Frame4.ZOrder 0
End Sub

Private Sub Command5_Click()
IEADAPT.Visible = True
End Sub

Private Sub Command6_Click()
OPERAADAPT.Visible = True
End Sub

Private Sub Command8_Click()
MsgBox "For any browser, simply look up how to use proxy servers.  If you don't change anything, the proxy servers address is 127.0.0.1 and its port is 6026."
End Sub
