Option Explicit
' ==============================================================================
' Script Name  : vbs-ftp-get-file.vbs
' Tool Version : 1.0.0
' Argument     : -
' Usage        : $0
' Return    0= : Normal Return Code
'           2= : Abnormal Return Code
' OS Version   : Windows 7 Home Premium 32bit
' ==============================================================================
' Date          Author        Changes
' 2018/04/09    Yuta Akama    New Creation
' ==============================================================================
' --+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' On error statement enabled
On Error Resume Next

' Variable definition
Dim objSH
Dim objFS
Dim TxtCmd
Dim FtpCmd
Dim RC

' Constant definition
Const RC_INF=0                                   ' Normal Return Code
Const RC_ERR=2                                   ' Abnormal Rerurn Code
Const FtpServer=""                               ' FTP Server
Const FtpUserName=""                             ' Username
Const FtpUserPass=""                             ' Password
Const RemoteDirectory=""                         ' Remote Directory
Const LocalDirectory=""                          ' Local Directory
Const GetFileName=""                             ' Get Filename
Const FtpExePath="%Systemroot%\System32\ftp.exe" ' ftp.exe
Const TempFileName="vbs-ftp-get-file.tmp"        ' Temporary Filename

' Object declaration
Set objSH=CreateObject("WScript.Shell")
Set objFS=CreateObject("Scripting.FileSystemObject")

' Get a file from FTP server
If objFS.FileExists(LocalDiretory & GetFileName) Then
  objFS.DeleteFile(LocalDiretory & GetFileName)
End If
Set TxtCmd = objFS.GetFolder(LocalDirectory).CreateTextFile(TempFileName)
TxtCmd.WriteLine "open " & FtpServer
TxtCmd.WriteLine FtpUserName
TxtCmd.WriteLine FtpUserPass
TxtCmd.WriteLine "cd " & RemoteDirectory
TxtCmd.WriteLine "lcd """ & LocalDirectory & """"
TxtCmd.WriteLine "type binary"
TxtCmd.WriteLine "get " & GetFileName
TxtCmd.WriteLine "bye"
TxtCmd.Close
FtpCmd = FtpExePath & " -s:""" & _
objFs.BuildPath(LocalDirectory,TempFileName) & """"
objSH.Run FtpCmd, 0,True

If Err.Number <> 0 Then
  WScript.Echo(Date() & " " & Right("0" & Hour(Now), 2) & Right(Time(), 6) _
  & " " & "[ERR] " & Err.Description)
  RC = RC_ERR
  WScript.Echo(Date() & " " & Right("0" & Hour(Now), 2) & Right(Time(), 6) _
  & " " & "[ERR] Script finished RC=" & RC)
  WScript.Quit(RC)
Else
  If Not objFS.FileExists(LocalDirectory & GetFileName) Then
    WScript.Echo(Date() & " " & Right("0" & Hour(Now), 2) & Right(Time(), 6) _
    & " " & "[ERR] Could not connect to FTP.")
    RC = RC_ERR
    WScript.Echo(Date() & " " & Right("0" & Hour(Now), 2) & Right(Time(), 6) _
    & " " & "[ERR] Script finished RC=" & RC)
    WScript.Quit(RC)
  End If
End If
objFS.DeleteFile TempFileName

Set objSH=Nothing
Set objFs=Nothing
Set TxtCmd=Nothing
Set FtpCmd=Nothing
Set RC=RC_INF
WScript.Quit(RC)
