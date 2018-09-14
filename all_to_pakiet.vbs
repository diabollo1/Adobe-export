call back_script("C:\Users\Tomek\AppData\Roaming\Adobe\InDesign\Version 9.0\pl_PL\Scripts\Scripts Panel\Adobe-export\all_to_pakiet.vbs", "all_to_pakiet.vbs")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set myInDesign = CreateObject("InDesign.Application")

copyingFonts = true
copyingLinkedGraphics = true
copyingProfiles = true
updatingGraphics = true
includingHiddenLayers = true
ignorePreflightErrors = true
creatingReport = false
versionComments = "comment"
forceSave = true

temp_msg = ""

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dim window, i
 
set window = createwindow()
	window.document.write "<html><body bgcolor=buttonface>Processing...<br><span id='output' style='font-size: 10px;'></span></body></html>"
	'window.document.title = "pakietowanie [" & myInDesign.Documents.Count * 2 & "]"
	window.resizeto 600, 900
	window.moveto 20, 20

temp_msg = temp_msg & "Do przerobienia: " & myInDesign.Documents.Count & "<br>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'save all
For myDocumentCounter = 1 To myInDesign.Documents.Count
	myInDesign.Documents.Item(myDocumentCounter).Save
Next

For myDocumentCounter = 1 To myInDesign.Documents.Count

	name = myInDesign.Documents.Item(myDocumentCounter).name
	name_bez = replace(myInDesign.Documents.Item(myDocumentCounter).name, ".indd", "")

	myPackageFolder = "D:\InD\zzz_pakiety" & "\PAKIET - " & name_bez	

	
	if myInDesign.Documents.Count = 1 then
		myPackageFolder = myInDesign.Documents.Item(myDocumentCounter).filePath & "\PAKIET - " & name_bez
	end if
		
	
	name_d = name_bez & "_d.pdf"
	name_small = name_bez & "_d_small.pdf"
	path_name_d = myPackageFolder & "\" & name_d
	path_name_small = myPackageFolder & "\" & name_small
	
	
	set myDocument = myInDesign.Documents.Item(myDocumentCounter)
	
	window.document.title = "pakietowanie [" & myDocumentCounter & "/" & myInDesign.Documents.Count & "]"
	
	show temp_msg & myDocumentCounter & " O O O O O " & name_bez & "<br>"
	
	'TWORZENIE PACZKI
	myDocument.package myPackageFolder, copyingFonts, copyingLinkedGraphics, copyingProfiles, updatingGraphics, includingHiddenLayers, ignorePreflightErrors, creatingReport, versionComments, forceSave
	show temp_msg & myDocumentCounter & " V O O O O " & name_bez & "<br>"
	
	'EXPORT idml
	myDocument.Export idExportFormat.idInDesignMarkup, myPackageFolder & "\" & name_bez & ".idml", False
	show temp_msg & myDocumentCounter & " V V O O O " & name_bez & "<br>"
	
	'EXPORT pdf_small
	myDocument.Export idExportFormat.idPDFType, path_name_small, False, myInDesign.pdfExportPresets.item("111_small")
	show temp_msg & myDocumentCounter & " V V V O O " & name_bez & "<br>"
	
	'EXPORT pdf_d
	myDocument.Export idExportFormat.idPDFType, path_name_d, False, myInDesign.pdfExportPresets.item("111")
	show temp_msg & myDocumentCounter & " V V V V O " & name_bez & "<br>"
	
	'RAROWANIE
	Call rarr(myPackageFolder,myPackageFolder)
	
	temp_msg = temp_msg & myDocumentCounter & " V V V V V " & name_bez & "<br>"
	show temp_msg
Next


For i = 1 To 100
	
	show temp_msg & "<br>" & i
	'WScript.Sleep 1000
	
Next

'save all
For myDocumentCounter = 1 To myInDesign.Documents.Count
	myInDesign.Documents.Item(myDocumentCounter).Save
Next

'close all
For myDocumentCounter = 1 To myInDesign.Documents.Count
	myInDesign.Documents.Item(1).Close
Next


window.close

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTION'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function rarr(co, gdzie)
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oShell = CreateObject("Wscript.Shell")
	path_rar = "C:\Program Files\WinRAR\WinRAR.exe"
	
	'http://acritum.com/software/manuals/winrar/
	oShell.Run """" & path_rar & """ a -r -ep1 -df """ & gdzie & """ """ & co & """"
	
end function

Function back_script(FullPath,Name)
	Set objShell = CreateObject("Wscript.Shell")
'	FullPath = Wscript.ScriptFullName
'	Name = Wscript.ScriptName


	Set filesys = CreateObject("Scripting.FileSystemObject")

	dest = "D:\InD\syf\skrypt\back\" & timeStamp & "___" & Name

	'msgbox(dest)
	filesys.CopyFile FullPath, dest
End Function

Function timeStamp()
    Dim t 
    t = Now
    timeStamp = Year(t) & "-" & _
    Right("0" & Month(t),2)  & "-" & _
    Right("0" & Day(t),2)  & "_" & _  
    Right("0" & Hour(t),2) & _
    Right("0" & Minute(t),2) '    '& _    Right("0" & Second(t),2) 
End Function

Function show(value)
    on error resume next
    window.output.innerhtml = value
    if err then wscript.quit
end Function
 
Function createwindow()
    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    dim signature, shellwnd, proc
    on error resume next
    set createwindow = nothing
    signature = left(createobject("Scriptlet.TypeLib").guid, 38)
    set proc = createobject("WScript.Shell").exec("mshta about:""<script>moveTo(-32000,-32000);</script><hta:application id=app border=dialog minimizebutton=no maximizebutton=no scroll=no showintaskbar=yes contextmenu=no selection=no innerborder=no /><object id='shellwindow' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shellwindow.putproperty('" & signature & "',document.parentWindow);</script>""")
    do
        if proc.status > 0 then exit Function
        for each shellwnd in createobject("Shell.Application").windows
            set createwindow = shellwnd.getproperty(signature)
            if err.number = 0 then exit Function
            err.clear
        next
    loop
end Function