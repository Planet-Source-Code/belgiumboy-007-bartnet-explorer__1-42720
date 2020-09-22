Attribute VB_Name = "GetSpecialFolder"
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
Private Function ReadKey(Value As String) As String
On Error Resume Next

    Dim b As Object
    
    Set b = CreateObject("wscript.shell")
    r = b.RegRead(Value)
    ReadKey = r
End Function

Public Function AdministrativeTools()
    AdministrativeTools = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Administrative Tools")
End Function

Public Function ApplicationData()
    ApplicationData = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\AppData")
End Function

Public Function TemporaryInternetFiles()
    TemporaryInternetFiles = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Cache")
End Function

Public Function CDBurning()
    CDBurning = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\CD Burning")
End Function

Public Function Cookies()
    Cookies = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Cookies")
End Function

Public Function Desktop()
    Desktop = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop")
End Function

Public Function Favorites()
    Favorites = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites")
End Function

Public Function Fonts()
    Fonts = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Fonts")
End Function

Public Function History()
    History = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\History")
End Function

Public Function LocalApplicationData()
    LocalApplicationData = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Local AppData")
End Function

Public Function LocalSettings()
    LocalSettings = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Local Settings")
End Function

Public Function MyMusic()
    MyMusic = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Music")
End Function

Public Function MyPictures()
    MyPictures = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Pictures")
End Function

Public Function MyVideos()
    MyVideos = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\My Video")
End Function

Public Function NetHood()
    NetHood = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\NetHood")
End Function

Public Function MyDocuments()
    MyDocuments = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")
End Function

Public Function PrintHood()
    PrintHood = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\PrintHood")
End Function

Public Function Programs()
    Programs = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Programs")
End Function

Public Function Recent()
    Recent = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Recent")
End Function

Public Function SendTo()
    SendTo = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\SendTo")
End Function

Public Function StartMenu()
    StartMenu = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Start Menu")
End Function

Public Function StartUp()
    StartUp = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\StartUp")
End Function

Public Function Templates()
    Templates = ReadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Templates")
End Function
