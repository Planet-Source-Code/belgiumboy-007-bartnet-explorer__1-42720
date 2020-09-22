Attribute VB_Name = "Module1"
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
Public Const ProgName As String = "BartNet Explorer"

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private FileInfo As typSHFILEINFO

Public Function ExtractIcon(filename As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function

Public Function GetSize(ByVal FileSize As Long)
    Select Case FileSize
        Case 0 To 999
            GetSize = Round(FileSize, 2) & " Bytes"
            Exit Function
        Case 1000 To 999999
            GetSize = Round(FileSize / 1000, 2) & " KB"
            Exit Function
        Case 1000000 To 999999999
            GetSize = Round(FileSize / 1000000, 2) & " MB"
            Exit Function
        Case Is >= 1000000000
            GetSize = Round(FileSize / 1000000000, 2) & " GB"
            Exit Function
    End Select
End Function

Function ShowFileProperties(filename As String, OwnerhWnd As Long) As Long
    Dim SEI As SHELLEXECUTEINFO
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    ShellExecuteEX SEI
    ShowFileProperties = SEI.hInstApp
End Function
