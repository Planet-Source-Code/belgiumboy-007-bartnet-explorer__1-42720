VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3585
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser w1 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1065
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.bartnet.freeservers.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "E-mail : webmaster@bartnet.freeservers.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form2.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Created By BelgiumBoy_007"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "BartNet Explorer"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''############################################################''
''##                                                        ##''
''##    Code Created By BelgiumBoy_007                      ##''
''##                                                        ##''
''##    E-mail : webmaster@bartnet.freeservers.com          ##''
''##                                                        ##''
''##    WebSite: http://www.bartnet.freeservers.com         ##''
''##                                                        ##''
''##    Copyright 2003 BartNet Corp. All Rights Reserved    ##''
''##                                                        ##''
''############################################################''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label3_Click()
    w1.Navigate "mailto:webmaster@bartnet.freeservers.com?subject=feedback"
End Sub

Private Sub Label4_Click()
    w1.Navigate "http://www.bartnet.freeservers.com", , "_new"
End Sub
