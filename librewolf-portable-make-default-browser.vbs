'Modified by Astroprogs for Librewolf Portable

'Registers LibreWolf Portable with Default Programs or Default Apps in Windows
'firefoxportable.vbs - created by Ramesh Srinivasan for Winhelponline.com
'v1.0 17-July-2022 - Initial release. Tested on Mozilla Firefox 102.0.1.0.
'v1.1 23-July-2022 - Minor bug fixes.
'v1.2 27-July-2022 - Minor revision. Cleaned up the code.
'Suitable for all Windows versions, including Windows 10/11.
'Tutorial: https://www.winhelponline.com/blog/register-firefox-portable-with-default-apps/

Option Explicit
Dim sAction, sAppPath, sExecPath, sIconPath, objFile, sbaseKey, sbaseKey2, sAppDesc
Dim sClsKey, ArrKeys, regkey
Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = oFSO.GetFile(WScript.ScriptFullName)
sAppPath = oFSO.GetParentFolderName(objFile)
sExecPath = "Y:\Apps\librewolf\LibreWolf-Portable.exe"
sIconPath = "Y:\Apps\librewolf\LibreWolf-Portable.exe"
sAppDesc = "Librewolf- An independent fork of Firefox, with the primary" & _
" goals of privacy, security and user freedom."

'Quit if LibreWolfPortable.exe is missing in the current folder!
If Not oFSO.FileExists (sExecPath) Then
   MsgBox "Please run this script from LibreWolf Portable folder. The script will now quit.", _
   vbOKOnly + vbInformation, "Register LibreWolf Portable with Default Apps"
   WScript.Quit
End If

If InStr(sExecPath, " ") > 0 Then
   sExecPath = """" & sExecPath & """"
   sIconPath = """" & sIconPath & """"
End If

sbaseKey = "HKCU\Software\"
sbaseKey2 = sbaseKey & "Clients\StartmenuInternet\LibreWolf Portable\"
sClsKey = sbaseKey & "Classes\"

If WScript.Arguments.Count > 0 Then
   If UCase(Trim(WScript.Arguments(0))) = "-REG" Then Call RegisterFirefoxPortable
   If UCase(Trim(WScript.Arguments(0))) = "-UNREG" Then Call UnRegisterFirefoxPortable
Else
   sAction = InputBox ("Type REGISTER to add LibreWolf Portable to Default Apps. " & _
   "Type UNREGISTER To remove.", "LibreWolf Portable Registration", "REGISTER")
   If UCase(Trim(sAction)) = "REGISTER" Then Call RegisterFirefoxPortable
   If UCase(Trim(sAction)) = "UNREGISTER" Then Call UnRegisterFirefoxPortable
End If

Sub RegisterFirefoxPortable   
   WshShell.RegWrite sbaseKey & "RegisteredApplications\LibreWolf Portable", _
   "Software\Clients\StartMenuInternet\LibreWolf Portable\Capabilities", "REG_SZ"
   
   'FirefoxHTML registration
   WshShell.RegWrite sClsKey & "LibrewolfHTM\", "LibreWolf HTML Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\FriendlyTypeName", "LibreWolf HTML Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\DefaultIcon\", sIconPath & ",1", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\shell\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\shell\open\command\", sExecPath & _
   " -url " & """" & "%1" & """", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\shell\open\ddeexec\", "", "REG_SZ"
   
   'FirefoxPDF registration
   WshShell.RegWrite sClsKey & "LibrewolfPDF\", "LibreWolf PDF Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfPDF\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "LibrewolfPDF\FriendlyTypeName", "LibreWolf PDF Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfPDF\DefaultIcon\", sIconPath & ",5", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfPDF\shell\open\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfPDF\shell\open\command\", sExecPath & _
   " -url " & """" & "%1" & """", "REG_SZ"
   
   'FirefoxURL registration
   WshShell.RegWrite sClsKey & "LibrewolfURL\", "LibreWolf URL", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfURL\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "LibrewolfURL\FriendlyTypeName", "LibreWolf URL", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfURL\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfURL\DefaultIcon\", sIconPath & ",1", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfURL\shell\open\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfURL\shell\open\command\", sExecPath & _
   " -url " & """" & "%1" & """", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfURL\shell\open\ddeexec\", "", "REG_SZ"   
   
   'Default Apps Registration/Capabilities
   WshShell.RegWrite sbaseKey2, "LibreWolf Portable", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationDescription", sAppDesc, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationName", "LibreWolf Portable", "REG_SZ" 
   WshShell.RegWrite sbaseKey2 & "Capabilities\FileAssociations\.pdf", "LibrewolfPDF", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\StartMenu", "LibreWolf Portable", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\open\command\", sExecPath, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\properties\", "LibreWolf &Options", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\properties\command\", sExecPath & " -preferences", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\safemode\", "LibreWolf &Safe Mode", "REG_SZ"   
   WshShell.RegWrite sbaseKey2 & "shell\safemode\command\", sExecPath & " -safe-mode", "REG_SZ"
   
   ArrKeys = Array ( _
   "FileAssociations\.avif", _
   "FileAssociations\.htm", _
   "FileAssociations\.html", _
   "FileAssociations\.shtml", _
   "FileAssociations\.svg", _
   "FileAssociations\.webp", _
   "FileAssociations\.xht", _
   "FileAssociations\.xhtml", _
   "URLAssociations\http", _
   "URLAssociations\https", _
   "URLAssociations\mailto" _
   )
   
   For Each regkey In ArrKeys
      WshShell.RegWrite sbaseKey2 & "Capabilities\" & regkey, "LibrewolfHTM", "REG_SZ"
   Next      
   
   'Override the default app name by which the program appears in Default Apps  (*Optional*)
   '(i.e., -- "Mozilla Firefox, Portable Edition" Vs. "Firefox Portable")
   'The official Mozilla Firefox setup doesn't add this registry key.
   WshShell.RegWrite sClsKey & "LibrewolfHTM\Application\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sClsKey & "LibrewolfHTM\Application\ApplicationName", "LibreWolf Portable", "REG_SZ"
   
   'Launch Default Programs or Default Apps after registering Firefox Portable   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"
End Sub


Sub UnRegisterFirefoxPortable
   sbaseKey = "HKCU\Software\"
   sbaseKey2 = "HKCU\Software\Clients\StartmenuInternet\LibreWolf Portable"   
   
   On Error Resume Next
   WshShell.RegDelete sbaseKey & "RegisteredApplications\LibreWolf Portable"
   On Error GoTo 0
   
   WshShell.Run "reg.exe delete " & sClsKey & "LibrewolfHTM" & " /f", 0
   WshShell.Run "reg.exe delete " & sClsKey & "LibrewolfPDF" & " /f", 0
   WshShell.Run "reg.exe delete " & sClsKey & "LibrewolfURL" & " /f", 0
   WshShell.Run "reg.exe delete " & chr(34) & sbaseKey2 & chr(34) & " /f", 0
   
   'Launch Default Apps after unregistering Firefox Portable   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"   
End Sub