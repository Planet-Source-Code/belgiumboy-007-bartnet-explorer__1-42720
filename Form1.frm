VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   -180
   ClientTop       =   735
   ClientWidth     =   13575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   13575
   Begin SHDocVwCtl.WebBrowser w1 
      Height          =   1335
      Left            =   3240
      TabIndex        =   3
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   2355
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
   Begin VB.PictureBox picTemp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   840
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   1560
   End
   Begin ComctlLib.ListView l1 
      Height          =   7000
      Left            =   4920
      TabIndex        =   1
      Top             =   50000
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12356
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ilsFileIconsLarge"
      SmallIcons      =   "ilsFileIconsSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.TreeView t1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   12303
      _Version        =   327682
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilsTreeview"
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   4920
      ScaleHeight     =   3255
      ScaleWidth      =   5295
      TabIndex        =   4
      Top             =   0
      Width           =   5295
      Begin VB.Label Label1 
         Caption         =   "Working ..."
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3615
      End
   End
   Begin ComctlLib.ImageList ilsFileIconsLarge 
      Left            =   6240
      Top             =   7275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ilsFileIconsSmall 
      Left            =   5640
      Top             =   7275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ilsTreeview 
      Left            =   5040
      Top             =   7275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5C12
            Key             =   "Desktop"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5F64
            Key             =   "MyDocuments"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":62B6
            Key             =   "MyComputer"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFolder 
         Caption         =   "New Folder"
      End
      Begin VB.Menu mnuNothing1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuIcons 
         Caption         =   "Icons"
      End
      Begin VB.Menu mnuSmallIcons 
         Caption         =   "Small Icons"
      End
      Begin VB.Menu mnuList 
         Caption         =   "List"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Report"
      End
      Begin VB.Menu mnuNothing4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrangeIconsBy 
         Caption         =   "Arrange Icons By"
         Begin VB.Menu mnuArrangeName 
            Caption         =   "Name"
         End
         Begin VB.Menu mnuArrangeSize 
            Caption         =   "Size"
         End
         Begin VB.Menu mnuArrangeType 
            Caption         =   "Type"
         End
         Begin VB.Menu mnuArrangeDateModified 
            Caption         =   "Date Modified"
         End
      End
      Begin VB.Menu mnuNothing2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuVisitBartNetOnline 
         Caption         =   "Visit BartNet Online"
      End
      Begin VB.Menu mnuNothing3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupFolderView 
         Caption         =   "View"
         Begin VB.Menu mnuPopupFolderIcons 
            Caption         =   "Icons"
         End
         Begin VB.Menu mnuPopupFolderSmallIcons 
            Caption         =   "Small Icons"
         End
         Begin VB.Menu mnuPopupFolderList 
            Caption         =   "List"
         End
         Begin VB.Menu mnuPopupFolderReport 
            Caption         =   "Report"
         End
      End
      Begin VB.Menu mnuPopupFolderNothing1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupFolderArrangeIconsBy 
         Caption         =   "Arrange Icons By"
         Begin VB.Menu mnuPopupFolderName 
            Caption         =   "Name"
         End
         Begin VB.Menu mnuPopupFolderSize 
            Caption         =   "Size"
         End
         Begin VB.Menu mnuPopupFolderType 
            Caption         =   "Type"
         End
         Begin VB.Menu mnuPopupFolderDateModified 
            Caption         =   "Date Modified"
         End
      End
      Begin VB.Menu mnuPopupFolderRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuPopupFolderNothing2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupFolderFolder 
         Caption         =   "New Folder"
      End
      Begin VB.Menu mnuPopupFolderNothing3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupFolderDeleteSelectedFile 
         Caption         =   "Delete Selected File / Folder"
      End
      Begin VB.Menu mnuPopupFolderNothing4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupFolderProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "Form1"
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
Private FirstName As String
Private FirstKey As String

Private Sub SetTreeview(ByVal WhichFolder As Folder, ByVal TheType As String, ByVal ParentName As String, ByVal FirstTime As Boolean)
    Dim Folders As Folders
    Dim Folder As Folder
    
    Set Folders = WhichFolder.SubFolders
    
    If FirstTime = True Then
        For Each Folder In Folders
            t1.Nodes.Add ParentName, tvwChild, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
        Next
    Else
        For Each Folder In Folders
            t1.Nodes.Add ParentName, tvwChild, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
        Next
    End If
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - 13665) / 2, (Screen.Height - 7755) / 2, 13665, 7755
    
    l1.ColumnHeaders(1).Width = 3000
    l1.ColumnHeaders(2).Width = 1000
    l1.ColumnHeaders(3).Width = 1800
    l1.ColumnHeaders(4).Width = 1620
    
    l1.Left = 4920
    l1.Top = -15
    
    Picture1.Height = l1.Height
    Picture1.Width = l1.Width
    
    With Label1
        .Left = (Picture1.Width - Label1.Width) / 2
        .Top = (Picture1.Height - Label1.Height) / 2
        .FontBold = True
        .FontSize = 20
    End With
    
    App.Title = ProgName
    Me.Caption = ProgName
    
    RefreshView
End Sub

Private Sub RefreshView()
    t1.Visible = False
    l1.Visible = False

    t1.Nodes.Clear
    l1.ListItems.Clear
    
    mnuSmallIcons.Checked = False
    mnuIcons.Checked = False
    mnuList.Checked = False
    mnuReport.Checked = True
    
    mnuArrangeName.Checked = True
    mnuArrangeSize.Checked = False
    mnuArrangeType.Checked = False
    mnuArrangeDateModified.Checked = False
    
    mnuPopupFolderName.Checked = True
    mnuPopupFolderSize.Checked = False
    mnuPopupFolderType.Checked = False
    mnuPopupFolderDateModified.Checked = False
    
    mnuPopupFolderIcons.Checked = False
    mnuPopupFolderSmallIcons.Checked = False
    mnuPopupFolderList.Checked = False
    mnuPopupFolderReport.Checked = True
    
    l1.View = lvwReport
    
    Dim fso As New FileSystemObject
    Dim Drives As Drives
    Dim Drive As Drive
    Dim Folders As Folders
    Dim Folder As Folder
    Dim Files As Files
    Dim File As File
    Dim Item As ListItem
    
    t1.Nodes.Add , , "Desktop", "Desktop", "Desktop", "Desktop"
    t1.Nodes.Add "Desktop", tvwChild, "Folder" & GetSpecialFolder.MyDocuments & "Before", "My Documents", "MyDocuments", "MyDocuments"
    
    Set Folder = fso.GetFolder(GetSpecialFolder.MyDocuments)
    Set Folders = Folder.SubFolders
    Set Files = Folder.Files
    
    For Each Folder In Folders
        t1.Nodes.Add "Folder" & GetSpecialFolder.MyDocuments & "Before", tvwChild, "Folder" & Folder.Path & "Before", Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
        Set Item = l1.ListItems.Add(, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(Folder.Path, ilsFileIconsSmall, picTemp, 16))
        Item.SubItems(1) = ""
        Item.SubItems(2) = Folder.Type
        Item.SubItems(3) = Folder.DateLastModified
        
        SetTreeview Folder, "Folder", "Folder" & Folder.Path & "Before", False
        
        DoEvents
    Next
    
    For Each File In Files
        Set Item = l1.ListItems.Add(, "File" & File.Path, File.Name, ExtractIcon(File.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(File.Path, ilsFileIconsSmall, picTemp, 16))
        Item.SubItems(1) = GetSize(File.Size)
        Item.SubItems(2) = File.Type
        Item.SubItems(3) = File.DateLastModified
        
        DoEvents
    Next
    
    t1.Nodes.Item(2).Selected = True
    t1.Nodes.Item(2).Expanded = True
    
    t1.Nodes.Add "Desktop", tvwChild, "MyComputer", "My Computer", "MyComputer", "MyComputer"
    
    Set Drives = fso.Drives
    For Each Drive In Drives
        If Drive.IsReady = True Then
            Set Folder = fso.GetFolder(Drive.RootFolder)
            Set Folders = Folder.SubFolders

            If Drive.DriveType = Fixed Then
                t1.Nodes.Add "MyComputer", tvwChild, "Drive" & Drive.DriveLetter, Drive.VolumeName & " (" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", ilsTreeview, picTemp, 16), ExtractIcon(Drive.Path, ilsTreeview, picTemp, 16)
                
                For Each Folder In Folders
                    t1.Nodes.Add "Drive" & Drive.DriveLetter, tvwChild, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
                Next
            Else
                t1.Nodes.Add "MyComputer", tvwChild, "Removable" & Drive.DriveLetter, Drive.VolumeName & " (" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", ilsTreeview, picTemp, 16), ExtractIcon(Drive.Path, ilsTreeview, picTemp, 16)
                
                For Each Folder In Folders
                    t1.Nodes.Add "Removable" & Drive.DriveLetter, tvwChild, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
                Next
            End If
        Else
            If Drive.DriveType = Fixed Then
                t1.Nodes.Add "MyComputer", tvwChild, "Drive" & Drive.DriveLetter, Drive.DriveLetter, ExtractIcon(Drive.Path & "\", ilsTreeview, picTemp, 16), ExtractIcon(Drive.Path, ilsTreeview, picTemp, 16)
            Else
                t1.Nodes.Add "MyComputer", tvwChild, "Removable" & Drive.DriveLetter, "(" & Drive.DriveLetter & ")", ExtractIcon(Drive.Path & "\", ilsTreeview, picTemp, 16), ExtractIcon(Drive.Path, ilsTreeview, picTemp, 16)
            End If
        End If
        
        DoEvents
    Next
    
    t1.Visible = True
    l1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub l1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo a
    Dim Folder As Folder
    Dim fso As New FileSystemObject
    Dim File As File
    Dim Item As ListItem
    
    If Mid(l1.SelectedItem.Key, 1, 6) = "Folder" Then
        Set Folder = fso.GetFolder(Mid(l1.SelectedItem.Key, 7, Len(l1.SelectedItem.Key) - 6))
        Folder.Name = NewString
        l1.SelectedItem.Key = "Folder" & Folder.Path
    Else
        Set File = fso.GetFile(Mid(l1.SelectedItem.Key, 5, Len(l1.SelectedItem.Key) - 4))
        File.Name = NewString
        l1.SelectedItem.Key = "File" & File.Path
        l1.ListItems.Remove (l1.SelectedItem.Index)
        Set Item = l1.ListItems.Add(, "File" & File.Path, File.Name, ExtractIcon(File.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(File.Path, ilsFileIconsSmall, picTemp, 16))
        Item.SubItems(1) = GetSize(File.Size)
        Item.SubItems(2) = File.Type
        Item.SubItems(3) = File.DateLastModified
    End If
    
    Exit Sub
    
a:
    MsgBox "Unable to rename", vbOKOnly + vbExclamation, "Error"
    Cancel = 1
    l1.SelectedItem.Text = FirstName
    l1.SelectedItem.Key = FirstKey
End Sub

Private Sub l1_BeforeLabelEdit(Cancel As Integer)
    FirstName = t1.SelectedItem.Text
    FirstKey = t1.SelectedItem.Key
End Sub

Private Sub l1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    If l1.SortKey = ColumnHeader.Index - 1 Then
        If l1.SortOrder = lvwAscending Then l1.SortOrder = lvwDescending Else l1.SortOrder = lvwAscending
    Else
        mnuArrangeName.Checked = False
        mnuArrangeSize.Checked = False
        mnuArrangeType.Checked = False
        mnuArrangeDateModified.Checked = False
        
        mnuPopupFolderName.Checked = False
        mnuPopupFolderSize.Checked = False
        mnuPopupFolderType.Checked = False
        mnuPopupFolderDateModified.Checked = False
        
        l1.SortKey = ColumnHeader.Index - 1
        
        l1.SortOrder = lvwAscending
        
        Select Case ColumnHeader.Text
            Case "Name"
                mnuArrangeName.Checked = True
                mnuPopupFolderName.Checked = True
            Case "Size"
                mnuArrangeSize.Checked = True
                mnuPopupFolderSize.Checked = True
            Case "Type"
                mnuArrangeType.Checked = True
                mnuPopupFolderType.Checked = True
            Case "Date Modified"
                mnuArrangeDateModified.Checked = True
                mnuPopupFolderDateModified.Checked = True
        End Select
    End If
End Sub

Private Sub l1_DblClick()
    Dim Node As Node
    Dim Key1 As String
    Dim Key2 As String
    Dim Index As Long
    
    If Mid(l1.SelectedItem.Key, 1, 6) = "Folder" Then  'IT'S A FOLDER
        Key1 = l1.SelectedItem.Key
        Key2 = l1.SelectedItem.Key & "Before"
        
        For Each Node In t1.Nodes
            Select Case Node.Key
                Case Key1
                    Index = Node.Index
                    Exit For
                Case Key2
                    Index = Node.Index
                    Exit For
            End Select
        Next
        
        Set Node = t1.Nodes.Item(Index)
        t1_NodeClick Node
    Else
        mnuPopupFolderProperties_Click
    End If
End Sub

Private Sub l1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu mnuPopup
End Sub


Private Sub mnuAbout_Click()
    Form2.Show vbModal, Me
End Sub

Private Sub mnuArrangeDateModified_Click()
    l1_ColumnClick l1.ColumnHeaders(4)
End Sub

Private Sub mnuArrangeName_Click()
    l1_ColumnClick l1.ColumnHeaders(1)
End Sub


Private Sub mnuArrangeSize_Click()
    l1_ColumnClick l1.ColumnHeaders(2)
End Sub


Private Sub mnuArrangeType_Click()
    l1_ColumnClick l1.ColumnHeaders(3)
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFolder_Click()
On Error GoTo a
    Dim fso As New FileSystemObject
    Dim b As String
    
    b = InputBox("Enter a name for the new folder", "Create New Folder", "New Folder", Me.Left + 2000, Me.Top + 2000)
    If b = "" Then Exit Sub
    
    If Mid(t1.SelectedItem.Key, 1, 6) = "Folder" Then
        If Right(t1.SelectedItem.Key, 6) = "Before" Then
            fso.CreateFolder (Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 12) & "\" & b)
            
            t1.Nodes.Add t1.SelectedItem.Key, tvwChild, Mid(t1.SelectedItem.Key, 1, Len(t1.SelectedItem.Key) - 6) & "\" & b, b, ExtractIcon(Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 12) & "\" & b, ilsTreeview, picTemp, 16), ExtractIcon(Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 12) & "\" & b, ilsTreeview, picTemp, 16)
            t1_NodeClick t1.SelectedItem
        Else
            fso.CreateFolder (Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 6) & "\" & b)
            
            t1.Nodes.Add t1.SelectedItem.Key, tvwChild, t1.SelectedItem.Key & "\" & b, b, ExtractIcon(Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 6) & "\" & b, ilsTreeview, picTemp, 16), ExtractIcon(Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 6) & "\" & b, ilsTreeview, picTemp, 16)
            t1_NodeClick t1.SelectedItem
        End If
    Else
        If Mid(t1.SelectedItem.Key, 1, 4) = "Drive" Then
            fso.CreateFolder (Mid(t1.SelectedItem.Key, 6, Len(t1.SelectedItem.Key) - 5) & "\" & b)
            
            t1.Nodes.Add t1.SelectedItem.Key, tvwChild, "Folder" & Mid(t1.SelectedItem.Key, 6, Len(t1.SelectedItem.Key) - 5) & "\" & b, b, ExtractIcon(Mid(t1.SelectedItem.Key, 6, Len(t1.SelectedItem.Key) - 5) & "\" & b, ilsTreeview, picTemp, 16), ExtractIcon(Mid(t1.SelectedItem.Key, 6, Len(t1.SelectedItem.Key) - 5) & "\" & b, ilsTreeview, picTemp, 16)
            t1_NodeClick t1.SelectedItem
        Else
            fso.CreateFolder (Mid(t1.SelectedItem.Key, 10, Len(t1.SelectedItem.Key) - 9) & "\" & b)
            
            t1.Nodes.Add t1.SelectedItem.Key, tvwChild, "Folder" & Mid(t1.SelectedItem.Key, 10, Len(t1.SelectedItem.Key) - 9) & "\" & b, b, ExtractIcon(Mid(t1.SelectedItem.Key, 10, Len(t1.SelectedItem.Key) - 9) & "\" & b, ilsTreeview, picTemp, 16), ExtractIcon(Mid(t1.SelectedItem.Key, 10, Len(t1.SelectedItem.Key) - 9) & "\" & b, ilsTreeview, picTemp, 16)
            t1_NodeClick t1.SelectedItem
        End If
    End If
    
    Exit Sub
    
a:
    MsgBox "An error has occurred : " & Err.Description, vbOKOnly + vbExclamation, ProgName
End Sub

Private Sub mnuIcons_Click()
    l1.View = lvwIcon
    
    mnuIcons.Checked = True
    mnuSmallIcons.Checked = False
    mnuList.Checked = False
    mnuReport.Checked = False
    
    mnuPopupFolderIcons.Checked = True
    mnuPopupFolderSmallIcons.Checked = False
    mnuPopupFolderList.Checked = False
    mnuPopupFolderReport.Checked = False
End Sub

Private Sub mnuList_Click()
    l1.View = lvwList
    
    mnuIcons.Checked = False
    mnuSmallIcons.Checked = False
    mnuList.Checked = True
    mnuReport.Checked = False
    
    mnuPopupFolderIcons.Checked = False
    mnuPopupFolderSmallIcons.Checked = False
    mnuPopupFolderList.Checked = True
    mnuPopupFolderReport.Checked = False
End Sub



Private Sub mnuPopupFolderDateModified_Click()
    mnuArrangeDateModified_Click
End Sub


Private Sub mnuPopupFolderDeleteSelectedFile_Click()
On Error GoTo a
    Dim File As File
    Dim Folder As Folder
    Dim fso As New FileSystemObject
    Dim Node As Node
    
    If Mid(l1.SelectedItem.Key, 1, 6) = "Folder" Then
        If MsgBox("Are you sure you want to delete the selected folder with all it's contents ?", vbYesNo + vbQuestion, ProgName) = vbYes Then
            Set Folder = fso.GetFolder(Mid(l1.SelectedItem.Key, 7, Len(l1.SelectedItem.Key) - 6))
            fso.DeleteFolder (Folder.Path)
            
            For Each Node In t1.Nodes
                If Node.Key = l1.SelectedItem.Key Or Node.Key = l1.SelectedItem.Key & "Before" Then t1.Nodes.Remove (Node.Index): Exit For
            Next
            
            l1.ListItems.Remove (l1.SelectedItem.Index)
        Else
            Exit Sub
        End If
    Else
        If MsgBox("Are you sure you want to delete the selected file ?", vbYesNo + vbQuestion, ProgName) = vbYes Then
            Set File = fso.GetFile(Mid(l1.SelectedItem.Key, 5, Len(l1.SelectedItem.Key) - 4))
            fso.DeleteFile (File.Path)
            
            l1.ListItems.Remove (l1.SelectedItem.Index)
        Else
            Exit Sub
        End If
    End If
    
    Exit Sub
    
a:
    MsgBox "An error has occurred : " & Err.Description, vbOKOnly + vbCritical, ProgName
End Sub


Private Sub mnuPopupFolderFolder_Click()
    mnuFolder_Click
End Sub

Private Sub mnuPopupFolderIcons_Click()
    mnuIcons_Click
End Sub

Private Sub mnuPopupFolderList_Click()
    mnuList_Click
End Sub

Private Sub mnuPopupFolderName_Click()
    mnuArrangeName_Click
End Sub

Private Sub mnuPopupFolderProperties_Click()
    If Mid(l1.SelectedItem.Key, 1, 6) = "Folder" Then
        ShowFileProperties Mid(l1.SelectedItem.Key, 7, Len(l1.SelectedItem.Key) - 6), Me.hwnd
    Else
        ShowFileProperties Mid(l1.SelectedItem.Key, 5, Len(l1.SelectedItem.Key) - 4), Me.hwnd
    End If
End Sub

Private Sub mnuPopupFolderRefresh_Click()
    mnuRefresh_Click
End Sub

Private Sub mnuPopupFolderReport_Click()
    mnuReport_Click
End Sub

Private Sub mnuPopupFolderSize_Click()
    mnuArrangeSize_Click
End Sub

Private Sub mnuPopupFolderSmallIcons_Click()
    mnuSmallIcons_Click
End Sub

Private Sub mnuPopupFolderType_Click()
    mnuArrangeType_Click
End Sub

Private Sub mnuRefresh_Click()
    If MsgBox("This will reset the view to the original starting position, do you wish to continue ?", vbYesNo + vbQuestion, ProgName) = vbYes Then RefreshView
End Sub

Private Sub mnuReport_Click()
    l1.View = lvwReport
    
    mnuIcons.Checked = False
    mnuSmallIcons.Checked = False
    mnuList.Checked = False
    mnuReport.Checked = True
    
    mnuPopupFolderIcons.Checked = False
    mnuPopupFolderSmallIcons.Checked = False
    mnuPopupFolderList.Checked = False
    mnuPopupFolderReport.Checked = True
End Sub

Private Sub mnuSmallIcons_Click()
    l1.View = lvwSmallIcon
    
    mnuIcons.Checked = False
    mnuSmallIcons.Checked = True
    mnuList.Checked = False
    mnuReport.Checked = False
    
    mnuPopupFolderIcons.Checked = False
    mnuPopupFolderSmallIcons.Checked = True
    mnuPopupFolderList.Checked = False
    mnuPopupFolderReport.Checked = False
End Sub

Private Sub mnuVisitBartNetOnline_Click()
    w1.Navigate "http://www.bartnet.freeservers.com", , "_new"
End Sub

Private Sub t1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo a
    Dim Folder As Folder
    Dim fso As New FileSystemObject
    
    If Mid(t1.SelectedItem.Key, 1, 6) = "Folder" Then
        If Right(t1.SelectedItem.Key, 6) = "Before" Then
            Set Folder = fso.GetFolder(Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 12))
            Folder.Name = NewString
            t1.SelectedItem.Key = "Folder" & Folder.Path & "Before"
        Else
            Set Folder = fso.GetFolder(Mid(t1.SelectedItem.Key, 7, Len(t1.SelectedItem.Key) - 6))
            Folder.Name = NewString
            t1.SelectedItem.Key = "Folder" & Folder.Path
        End If
    Else
        GoTo a
    End If
    
    Exit Sub
    
a:
    MsgBox "Unable to rename", vbOKOnly + vbExclamation, "Error"
    Cancel = 1
    t1.SelectedItem.Text = FirstName
    t1.SelectedItem.Key = FirstKey
End Sub

Private Sub t1_BeforeLabelEdit(Cancel As Integer)
    FirstName = t1.SelectedItem.Text
    FirstKey = t1.SelectedItem.Key
End Sub

Private Sub t1_NodeClick(ByVal Node As ComctlLib.Node)
On Error GoTo a
    
    If Node.Key = "MyComputer" Then Node.Expanded = True: Exit Sub
    If Node.Key = "Desktop" Then l1.ListItems.Clear: Exit Sub: Node.Expanded = False
    
    Dim File As File
    Dim Files As Files
    Dim Folder As Folder
    Dim Folders As Folders
    Dim fso As New FileSystemObject
    Dim Item As ListItem
    Dim Drive As Drive
    
    If Mid(Node.Key, 1, 6) = "Folder" Then  'IT'S A FOLDER
        If Right(Node.Key, 6) = "Before" Then   'IT'S BEEN ACCESSED BEFORE
            Set Folder = fso.GetFolder(Mid(Node.Key, 7, Len(Node.Key) - 12))
            Set Folders = Folder.SubFolders
            Set Files = Folder.Files
            
            l1.Visible = False
            l1.ListItems.Clear
            
            For Each Folder In Folders
                Set Item = l1.ListItems.Add(, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(Folder.Path, ilsFileIconsSmall, picTemp, 16))
                Item.SubItems(1) = ""
                Item.SubItems(2) = Folder.Type
                Item.SubItems(3) = Folder.DateLastModified
                
                DoEvents
            Next
            
            For Each File In Files
                Set Item = l1.ListItems.Add(, "File" & File.Path, File.Name, ExtractIcon(File.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(File.Path, ilsFileIconsSmall, picTemp, 16))
                Item.SubItems(1) = GetSize(File.Size)
                Item.SubItems(2) = File.Type
                Item.SubItems(3) = File.DateLastModified
                
                DoEvents
            Next
            
            l1.Visible = True
            
            Node.Expanded = True
        Else    'IT HASN'T BEEN ACCESSED BEFORE
            Set Folder = fso.GetFolder(Mid(Node.Key, 7, Len(Node.Key) - 6))
            Set Folders = Folder.SubFolders
            Set Files = Folder.Files
            
            l1.Visible = False
            l1.ListItems.Clear
            
            For Each Folder In Folders
                t1.Nodes.Add Node.Key, tvwChild, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
                Set Item = l1.ListItems.Add(, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(Folder.Path, ilsFileIconsSmall, picTemp, 16))
                Item.SubItems(1) = ""
                Item.SubItems(2) = Folder.Type
                Item.SubItems(3) = Folder.DateLastModified
                
                DoEvents
            Next
            
            For Each File In Files
                Set Item = l1.ListItems.Add(, "File" & File.Path, File.Name, ExtractIcon(File.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(File.Path, ilsFileIconsSmall, picTemp, 16))
                Item.SubItems(1) = GetSize(File.Size)
                Item.SubItems(2) = File.Type
                Item.SubItems(3) = File.DateLastModified
                
                DoEvents
            Next
            
            Node.Key = Node.Key & "Before"
            
            l1.Visible = True
            
            Node.Expanded = True
        End If
    Else
        If Mid(Node.Key, 1, 5) = "Drive" Then    'IT'S A FIXED DRIVE
            Set Folder = fso.GetFolder(Mid(Node.Key, 6, Len(Node.Key) - 5) & ":\")
            Set Folders = Folder.SubFolders
            Set Files = Folder.Files
            
            l1.Visible = False
            l1.ListItems.Clear
            
            For Each Folder In Folders
                Set Item = l1.ListItems.Add(, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(Folder.Path, ilsFileIconsSmall, picTemp, 16))
                Item.SubItems(1) = ""
                Item.SubItems(2) = Folder.Type
                Item.SubItems(3) = Folder.DateLastModified
                
                DoEvents
            Next
            
            For Each File In Files
                Set Item = l1.ListItems.Add(, "File" & File.Path, File.Name, ExtractIcon(File.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(File.Path, ilsFileIconsSmall, picTemp, 16))
                Item.SubItems(1) = GetSize(File.Size)
                Item.SubItems(2) = File.Type
                Item.SubItems(3) = File.DateLastModified
                
                DoEvents
            Next
            
            l1.Visible = True
            
            Node.Expanded = True
        Else    'IT A NON-FIXED DRIVE
On Error Resume Next
            Set Drive = fso.GetDrive(Mid(Node.Key, 10, Len(Node.Key) - 9) & ":\")
            
            Dim sNode As Node
            Dim sNodes As Nodes
            
            Set Nodes = t1.Nodes
            
            For Each sNode In sNodes
                If sNode.Parent = Node.Key Then t1.Nodes.Remove sNode.Index
            Next
                
            If Drive.IsReady = True Then
                Set Folder = Drive.RootFolder
                Set Folders = Folder.SubFolders
                Set Files = Folder.Files
                
                l1.Visible = False
                l1.ListItems.Clear
                
                For Each Folder In Folders
                    t1.Nodes.Add Node.Key, tvwChild, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16), ExtractIcon(Folder.Path, ilsTreeview, picTemp, 16)
                    Set Item = l1.ListItems.Add(, "Folder" & Folder.Path, Folder.Name, ExtractIcon(Folder.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(Folder.Path, ilsFileIconsSmall, picTemp, 16))
                    Item.SubItems(1) = ""
                    Item.SubItems(2) = Folder.Type
                    Item.SubItems(3) = Folder.DateLastModified
                    
                    DoEvents
                Next
                
                For Each File In Files
                    Set Item = l1.ListItems.Add(, "File" & File.Path, File.Name, ExtractIcon(File.Path, ilsFileIconsLarge, picTemp, 32), ExtractIcon(File.Path, ilsFileIconsSmall, picTemp, 16))
                    Item.SubItems(1) = GetSize(File.Size)
                    Item.SubItems(2) = File.Type
                    Item.SubItems(3) = File.DateLastModified
                    
                    DoEvents
                Next
                
                l1.Visible = True
                
                Node.Expanded = True
            Else
                MsgBox "Drive is not ready", vbOKOnly + vbCritical, ProgName
                l1.ListItems.Clear
            End If
        End If
    End If
    
    DoEvents
    
    Exit Sub

a:
    MsgBox "An error has occurred : " & Err.Description, vbOKOnly + vbCritical, ProgName
End Sub


