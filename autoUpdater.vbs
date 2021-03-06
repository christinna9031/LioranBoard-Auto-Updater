'LioranBoard Auto Updater by Christina K, V1.0
set xHttp = CreateObject("Microsoft.XMLHTTP")
set bStrm = CreateObject("Adodb.Stream")
set filesys = CreateObject("Scripting.FileSystemObject")
set objShell = CreateObject("Shell.Application")
set objWMIService = GetObject ("winmgmts:")

'function to force console and show progress bar
Function printi(txt)
    WScript.StdOut.Write txt
End Function    

Function printr(txt)
    back(Len(txt))
    printi txt
End Function

Function back(n)
    Dim i
    For i = 1 To n
        printi chr(08)
    Next
End Function   

Function percent(x, y, d)
    percent = FormatNumber((x / y) * 100, d) & "%"
End Function

Function progress(x, y)
    Dim intLen, strPer, intPer, intProg, intCont
    intLen  = 22
    strPer  = percent(x, y, 1)
    intPer  = FormatNumber(Replace(strPer, "%", ""), 0)
    intProg = intLen * (intPer / 100)
    intCont = intLen - intProg
    printr String(intProg, ChrW(9608)) & String(intCont, ChrW(9618)) & " " & strPer
End Function

Function ForceConsole()
    Set oWSH = CreateObject("WScript.Shell")
    vbsInterpreter = "cscript.exe"

    If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
        oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
        WScript.Quit
    End If
End Function

'Set Path
path = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
'Set new folder name to current date
d = Date()
dateStr = Year(d) & "-" & Right("00" & Month(d), 2) & "-" & Right("00" & Day(d), 2)
filePath = path & "\LioranBoard" & "-" & dateStr

'Check if user placed the script in the right folder
If NOT (filesys.FolderExists(path & "\LioranBoard Receiver(PC)")) Then
WScript.Echo "LioranBoard Receiver(PC) folder not found. Place autoUpdater.vbs in lioranboard folder and try again."
WScript.Quit
End If

'Check if LB Receiver or Stream Deck is running and ask user to close it
Set proc = objWMIService.ExecQuery("select * from Win32_Process Where Name='LioranBoard Receiver.exe'")
If proc.count > 0 Then 
WScript.Echo "Please close LioranBoard Receiver and try again!"
WScript.Quit
End If
Set proc = objWMIService.ExecQuery("select * from Win32_Process Where Name='LioranBoard Stream Deck.exe'")
If proc.count > 0 Then 
WScript.Echo "Please close LioranBoard Stream Deck and try again!"
WScript.Quit
End If

ForceConsole()
WScript.StdOut.WriteLine "Downloading the files. Please wait..."
Call progress(0, 100)

'Download zip file
xHttp.Open "GET", "http://lioran.servehttp.com/share/lioranboard/lioranboard.zip", False
on error resume next
xHttp.Send

Call progress(10, 100)

If(xHttp.Status <> 200) Then
WScript.StdOut.WriteLine " "
WScript.StdOut.WriteLine "Error downloading the file: " & xHttp.statusText & ". Server might be temporarily down. Please try again later!"
WScript.StdOut.WriteLine "Press [ENTER] to close this window..."
WScript.StdIn.ReadLine
WScript.Quit
End If

With bStrm
    .type = 1 
    .open
    .write xHttp.responseBody
    .savetofile filePath & ".zip", 2
end With

Call progress(20, 100)

'Extract zip file
ZipFile=filePath & ".zip"
ExtractTo=filePath
If NOT filesys.FolderExists(ExtractTo) Then
   filesys.CreateFolder(ExtractTo)
End If

Call progress(30, 100)

'Extract the contents of the zip file.
set FilesInZip=objShell.NameSpace(ZipFile).items
objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)

Call progress(50, 100)

'Copy files
If filesys.FileExists(path & "\LioranBoard Receiver(PC)\LioranBoard Receiver.exe") Then
  filesys.DeleteFile path & "\LioranBoard Receiver(PC)\LioranBoard Receiver.exe", True
End If
If filesys.FileExists(path & "\LioranBoard Receiver(PC)\data.win") Then
  filesys.DeleteFile path & "\LioranBoard Receiver(PC)\data.win", True
End If
filesys.CopyFile filePath & "\LioranBoard Receiver(PC)\LioranBoard Receiver.exe", path & "\LioranBoard Receiver(PC)\LioranBoard Receiver.exe", True 
Call progress(60, 100)
filesys.CopyFile filePath & "\LioranBoard Receiver(PC)\data.win", path & "\LioranBoard Receiver(PC)\data.win", True 
Call progress(70, 100)
filesys.CopyFolder filePath & "\LioranBoard Stream deck(PC)", path & "\LioranBoard Stream deck(PC)", True   
Call progress(80, 100)
filesys.CopyFolder filePath & "\LioranBoard Stream deck(Android)", path & "\LioranBoard Stream deck(Android)", True 
Call progress(90, 100)

'Let user know it's finished
Call progress(100, 100)
WScript.StdOut.WriteLine " "
WScript.StdOut.WriteLine " "
WScript.StdOut.WriteLine "LioranBoard Update Complete!"

'Clean up
filesys.DeleteFolder filePath, True
filesys.DeleteFile filePath & ".zip", True
set xHttp = Nothing
set bStrm = Nothing
set filesys = Nothing
set objShell = Nothing
set objWMIService = Nothing
WScript.StdOut.WriteLine " "
WScript.StdOut.WriteLine "Press [ENTER] to close this window..."
WScript.StdIn.ReadLine