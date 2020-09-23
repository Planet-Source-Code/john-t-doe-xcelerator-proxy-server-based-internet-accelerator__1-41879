VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Xcelerator"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3625
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Xcelerator"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PROXYLISTEN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "HTTPCONNECT(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "PL"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Proxy Options"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).Control(1)=   "PORTN"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Acceleration Options"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Command12"
      Tab(2).ControlCount=   4
      Begin MSWinsockLib.Winsock PL 
         Left            =   1680
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock HTTPCONNECT 
         Index           =   0
         Left            =   3000
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock PROXYLISTEN 
         Left            =   840
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Save Options"
         Height          =   375
         Left            =   -69600
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Content Protection"
         Height          =   1215
         Left            =   -71040
         TabIndex        =   21
         Top             =   0
         Width           =   3015
         Begin VB.CommandButton Command11 
            Caption         =   "Delete"
            Height          =   255
            Left            =   2280
            TabIndex        =   25
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton Command10 
            Caption         =   "ADD"
            Height          =   255
            Left            =   1800
            TabIndex        =   24
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox WORDTXT 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1575
         End
         Begin VB.ListBox WORDLIST 
            Height          =   450
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bandwidth Protection"
         Height          =   1095
         Left            =   -75000
         TabIndex        =   15
         Top             =   600
         Width           =   3975
         Begin VB.CommandButton Command9 
            Caption         =   "Block Ads"
            Height          =   255
            Left            =   2760
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Delete"
            Height          =   255
            Left            =   2160
            TabIndex        =   19
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Add"
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox SITETXT 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1335
         End
         Begin VB.ListBox SITELIST 
            Height          =   450
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Connection Acceleration"
         Height          =   615
         Left            =   -75000
         TabIndex        =   12
         Top             =   0
         Width           =   3975
         Begin VB.TextBox MULTICONNUM 
            Height          =   285
            Left            =   3480
            TabIndex        =   14
            Text            =   "20"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Maximum Connection Attempts:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save Options"
         Height          =   375
         Left            =   -69600
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox PORTN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -69000
         TabIndex        =   8
         Text            =   "8026"
         Top             =   120
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Proxy Security"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   3
         Top             =   0
         Width           =   5175
         Begin VB.CommandButton Command6 
            Caption         =   "Delete"
            Height          =   255
            Left            =   4440
            TabIndex        =   11
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add"
            Height          =   255
            Left            =   3720
            TabIndex        =   7
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox IPTXT 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "66.189.40.235"
            Top             =   1080
            Width           =   3495
         End
         Begin VB.ListBox IPLIST 
            Height          =   450
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   4935
         End
         Begin VB.CheckBox IPALLOW 
            Caption         =   "Allow only the following IP's to connect:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   6735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69720
         TabIndex        =   9
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldport As Integer
Dim oldnum As Integer
Dim datamess As String
Dim webpagedata() As Byte
Dim bwaiting As Boolean
Private Sub Command1_Click()
    'You can't press start now because its started
    Command1.Enabled = False
    'Load the HTTP Connecters(Winsock)
    For i% = 1 To MULTICONNUM.Text
        Load HTTPCONNECT(i%)
    Next
    'Make proxy server listen for requests
    PROXYLISTENNOW

End Sub
Sub PROXYLISTENNOW()
    'Make the winsock that handles requests stop what its doing
    PROXYLISTEN.Close
    PL.Close 'Make the winsock that listens for requests listen for them
    PL.LocalPort = PORTN.Text
    PL.Listen
    bwaiting = False 'We are not handling anything
End Sub
Private Sub Command10_Click()
    WORDLIST.AddItem WORDTXT.Text 'Add a word to filter from the url
End Sub

Private Sub Command11_Click()
    'Remove a word to filter from the URL
    On Error GoTo eh:
    WORDLIST.RemoveItem WORDLIST.ListIndex
eh:
    MsgBox "Select an Item"
End Sub

Private Sub Command12_Click()
    'Save acceleration options
    SaveSetting "XCEL", "ACCEL", "MAXCON", MULTICONNUM.Text
    SaveSetting "XCEL", "ACCEL", "SITENUM", SITELIST.ListCount
    For i% = 0 To SITELIST.ListCount - 1
        SaveSetting "XCEL", "ACCEL", "SITE" & i%, SITELIST.List(i%)
    Next
    SaveSetting "XCEL", "ACCEL", "WORDNUM", WORDLIST.ListCount
    For i% = 0 To WORDLIST.ListCount - 1
        SaveSetting "XCEL", "ACCEL", "WORD" & i%, WORDLIST.List(i%)
    Next
End Sub
Sub ACCELLOAD()
'Load Acceleration Options
If GetSetting("XCEL", "ACCEL", "RAN") = "YES" Then 'Have they been saved before?
    'If so, load them
    MULTICONNUM.Text = GetSetting("XCEL", "ACCEL", "MAXCON")
    For i% = 0 To GetSetting("XCEL", "ACCEL", "SITENUM") - 1
        SITELIST.AddItem GetSetting("XCEL", "ACCEL", "SITE" & i%)
        DoEvents
    Next
    For i% = 0 To GetSetting("XCEL", "ACCEL", "WORDNUM") - 1
        WORDLIST.AddItem GetSetting("XCEL", "ACCEL", "WORD" & i%)
        DoEvents
    Next
Else
    'If not, create them
    SaveSetting "XCEL", "ACCEL", "MAXCON", 50
    SaveSetting "XCEL", "ACCEL", "SITENUM", 0
    SaveSetting "XCEL", "ACCEL", "WORDNUM", 0
    SaveSetting "XCEL", "ACCEL", "RAN", "YES"
    MULTICONNUM.Text = 50
    
    answ = MsgBox("This is your first time running Xcelerator.  Would you like to view the help file?", vbYesNo)
    If answ = vbYes Then
        HELP.Visible = True
    End If

End If
End Sub

Private Sub Command3_Click()
    HELP.Visible = True 'Show the help dialog
End Sub

Private Sub Command4_Click()
    IPLIST.AddItem IPTXT.Text 'Add an IP to the IPList
End Sub

Private Sub Command5_Click()
'Save the proxy settings/options

SaveSetting "XCEL", "PROXY", "IPALLOW", IPALLOW.Value
SaveSetting "XCEL", "PROXY", "PORT", PORTN.Text
SaveSetting "XCEL", "PROXY", "IPNUM", IPLIST.ListCount
For i% = 0 To IPLIST.ListCount - 1
    SaveSetting "XCEL", "PROXY", "IP" & i%, IPLIST.List(i%)
Next
End Sub
Sub LoadProxy()
'Load the proxy options
    If GetSetting("XCEL", "PROXY", "RAN") = "YES" Then 'Have they been saved before?
        IPALLOW.Value = GetSetting("XCEL", "PROXY", "IPALLOW")
        PORTN.Text = GetSetting("XCEL", "PROXY", "PORT")
        For i% = 0 To GetSetting("XCEL", "PROXY", "IPNUM") - 1
            IPLIST.AddItem GetSetting("XCEL", "PROXY", "IP" & i%)
            DoEvents
        Next
    Else
        SaveSetting "XCEL", "PROXY", "IPALLOW", 0
        SaveSetting "XCEL", "PROXY", "PORT", 2086
        SaveSetting "XCEL", "PROXY", "RAN", "YES"
    End If
End Sub

Private Sub Command6_Click()
    'remove an IP from the IPList
    On Error Resume Next
    IPLIST.RemoveItem IPLIST.ListIndex
End Sub

Private Sub Command7_Click()
    'Add a website to the blocked site list
    SITELIST.AddItem SITETXT.Text
End Sub

Private Sub Command8_Click()
    SITELIST.RemoveItem SITELIST.ListIndex 'Remove an item from the site list
End Sub

Private Sub Form_Load()
    PORTN = 8026 'Set the port to default port
    LoadProxy 'Load proxy options
    ACCELLOAD 'Load acceleration options
End Sub

Private Sub HTTPCONNECT_Close(Index As Integer)
    PROXYLISTENNOW 'Reset everything
End Sub

Private Sub HTTPCONNECT_Connect(Index As Integer)
    'One of our many attempts to connect has succeeded.  Tell the other winsocks to stop
    For i% = 0 To MULTICONNUM.Text
        If Index <> i% Then
            HTTPCONNECT(i%).Close
        End If
    Next
    'Send the website the data sent by the browser to the proxy server(this program)
    HTTPCONNECT(Index).SendData datamess
    webpagedata = ""
End Sub

Private Sub HTTPCONNECT_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Since this works like a proxy server, it sends data back to the browser so we do it
    On Error Resume Next
    Dim datax As String
    HTTPCONNECT(Index).GetData webpagedata()
    PROXYLISTEN.SendData webpagedata()
End Sub

Private Sub HTTPCONNECT_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'If we have an error, reset
    HTTPCONNECT(Index).Close
    PROXYLISTENNOW
End Sub

Private Sub IPALLOW_Click()
    'If we limit the IPs, allow the user to enter IP's
    If IPALLOW.Value = 1 Then
        IPTXT.Enabled = True
        IPLIST.Enabled = True
        Command1.Enabled = True
    Else
        IPTXT.Enabled = False
        IPLIST.Enabled = False
        Command1.Enabled = False
    End If
End Sub

Private Sub MULTICONNUM_Change()
    'Only allow the text box to store numeric values between 1 and 100
    On Error GoTo eh:
    If Len(MULTICONNUM.Text) = 0 Then Exit Sub
    Dim checknum As Integer
    checknum = MULTICONNUM.Text
    If checknum > 100 Then
        GoTo eh
    End If
    oldnum = MULTICONNUM.Text
    Exit Sub
eh:
    MULTICONNUM.Text = oldport
    MsgBox "The number of connections must be a number between 1 and 100"
End Sub

Private Sub PL_ConnectionRequest(ByVal requestID As Long)
    'If were busing, wait until we are done doing something
    Do While bwaiting = True
        DoEvents
    Loop
    'Make the computer know we are doing something
    bwaiting = True
    'Make proxylisten handle the request
    PROXYLISTEN.Close
    PROXYLISTEN.Accept requestID
    'Check that if we limit IP's, that the computer connecting is one of them
    If IPALLOW.Value = 1 Then
        For i% = 0 To IPLIST.ListCount - 1
            If IPLIST.List(i%) = PROXYLISTEN.RemoteHostIP Then
                Exit Sub 'If so, exit the sub
            End If
        Next
    Else
        Exit Sub 'If we don't limit IP's exit the sub
    End If
    'If we don't exit the sub it means we limit IP's and the connecting computer wasn't one of them: reset the connection
    PROXYLISTENNOW
End Sub


Private Sub PORTN_Change()
    'Limit text box to numbers between 1 and 600000
    On Error GoTo eh:
    If Len(PORTN.Text) = 0 Then Exit Sub
    Dim checknum As Integer
    checknum = PORTN.Text
    If checknum > 60000 Then
        GoTo eh
    End If
    oldport = PORTN.Text
    Exit Sub
eh:
    PORTN.Text = oldport
    MsgBox "The port must be a number between 1 and 60000"
End Sub

Private Sub PROXYLISTEN_Close()
    'If the browser closes the connection, listen again
    bwaiting = False
    PROXYLISTENNOW
End Sub
Private Sub PROXYLISTEN_DataArrival(ByVal bytesTotal As Long)
    'Parse the request and get the website address
    Dim URL As String
    Dim datax As String
    PROXYLISTEN.GetData datax
    For i% = 1 To Len(datax)
        If Mid(datax, i%, 2) = "//" Then 'The website starts after the //(HTTP://)
            Exit For
        End If
    Next
    
    For Y% = i% + 2 To Len(datax)
        If Mid(datax, Y%, 1) = "/" Then 'The website ends at the /(www.planet-source-code.com/) 'In the case of www.planet-source-code.com/Junk/Cheese.html  , the /junk/Cheese.html is part of the request, not part of the website URL, so we connect to the website URL and send the request.
            URL = Mid(datax, i% + 2, Y% - (i% + 2))
            Exit For
        End If
    Next
    FastConnect URL, datax 'Get the website
End Sub

Private Sub ProxyListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'If we have an error reset the connection
    bwaiting = False
    PROXYLISTENNOW
End Sub
Sub FastConnect(URL As String, DT As String)
        'Connects to a url
        'If the website is on the blocked list, do not connect
        For i% = 0 To SITELIST.ListCount - 1
            If InStr(1, URL, SITELIST.List(i%)) Then
                PROXYLISTENNOW
                Exit Sub
            End If
        Next
        'Store the request in a variable that doesn't loose its data
        datamess = DT
        'Make each HTTPConnect attempt to connect to the url
        For i% = 0 To MULTICONNUM.Text
            HTTPCONNECT(i%).Close
            HTTPCONNECT(i%).Connect URL, 80
        Next
End Sub
