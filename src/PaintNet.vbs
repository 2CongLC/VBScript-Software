'Coder : 2CongLC.Vn
'File Name : PaintNet.vbs
'Copyright © 2017 By 2CongLC.Vn | All Rights Reserved.

Option Explicit

Dim WS : Set WS = CreateObject("WSCript.Shell")
Dim SA : Set SA = CreateObject("Shell.Application")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Dim WMI: Set WMI = GetObject("winmgmts:\\.\root\cimv2")
RunasAdmin()

Const InsDir = "paint.net"
Const ZIP = "PaintNET.zip"

Dim Root : Root = FSO.GetParentFolderName(WScript.ScriptFullName)
Dim Files : Files = Root & "\Files\" & ZIP
Dim Mode : If IsProc = "x86" Then Mode = "PaintDotNet.exe" Else Mode = "PaintDotNet.exe"


If FSO.FileExists(Files) Then 
 TaskKill(Mode)
 Call ProcessUnzip(Files, InsDir)
 If FSO.FileExists(Envs("%ProgramFiles%") & "\" & InsDir & "\" & Mode) Then
  Call DesktopShortCut(Envs("%ProgramFiles%") & "\" & InsDir, Mode, "Paint.Net")
  WS.Exec(Envs("%ProgramFiles%") & "\" & InsDir & "\" & Mode)
 End if
Else
 If FSO.FolderExists(Root & "\Extract\" & InsDir & "\" ) Then
  Call DesktopShortCut(Root & "\Extract\" & InsDir, Mode, "Paint.Net")
  WS.Exec(Root & "\Extract\" & InsDir & "\" & Mode)
 End if
End if

Private Function IsOS()
 Dim i
 For Each i in WMI.execquery("Select * from Win32_OperatingSystem")
  IsOS = i.caption
  Next 
 End Function

Private Function IsProc()'https://msdn.microsoft.com/en-us/library/aa394373(v=vs.85).aspx
 Dim i
 For Each i in WMI.execquery("Select * From Win32_Processor")
  If i.AddressWidth = 32 Then IsProc = "x86"
  If i.AddressWidth = 64 Then IsProc = "x64"
  Next 
End Function
 
Private Sub ProcessUnzip(File, ToPath)
 Call UnZip(File,Envs("%ProgramFiles%") & "\" & ToPath)
 End Sub

Private Function Envs(cmd)
 Envs = WS.ExpandEnvironmentStrings(cmd)
 End Function

Private Sub BuildFullPath(Path)
 If Not FSO.FolderExists(Path) Then
  BuildFullPath FSO.GetParentFolderName(Path)
  FSO.CreateFolder(Path)
  End if
 End Sub
 
Private Sub UnZip(File, ToPath)
 BuildFullPath(ToPath)
 Dim FileZip: set FileZip = SA.NameSpace(File).items
 SA.NameSpace(ToPath).CopyHere(FileZip)
 End Sub 

Private Sub DesktopShortCut(LocalFolder, FileName, DeskName)
 Dim link: Set link = WS.CreateShortcut(WS.SpecialFolders("Desktop") & "\" &  DeskName & ".lnk")
 With link
 .TargetPath =  LocalFolder & "\" & FileName
 .Arguments = ""
 .Description = ""
 .HotKey = ""
 .IconLocation = LocalFolder & "\" & FileName & ", 0"
 .WindowStyle = "1"
 .WorkingDirectory = LocalFolder
 .Save
 End With
 End Sub 

Private Sub PinToTaskbar(File)
 Dim FileName: FileName = FSO.GetFileName(File)
 Dim Folder: Set Folder = SA.Namespace(Left(File,Len(File)-Len(FileName)))
 Dim i
 For Each i in Folder.ParseName(FileName).Verbs
  If Instr(IsOS,"Microsoft Windows 10") <> 0 Then
   If Replace(i.name, "&", "") = "Pin to taskbar" Then i.Doit
   Else
   If Replace(i.name, "&", "") = "Pin to Taskbar" Then i.Doit
   End if
  Next
 End Sub 

Private Function TaskKill(Name)
 Dim Item
 For Each Item in WMI.ExecQuery("Select * From Win32_Process Where Name = " & "'" & Name & "'")
  TaskKill = Item.Terminate() 
  Next
 End Function 
 
 
Private Sub RunasAdmin()
 If Err.Number = 0 Then
  If WSCript.Arguments.Length = 0 Then
   SA.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & Chr(32) & "/2CongLC.Vn", , "runas", 1
   WSCript.Quit
   End if
  End if
 End Sub

Set WS = Nothing
Set SA = Nothing
Set FSO = Nothing
Set WMI = Nothing