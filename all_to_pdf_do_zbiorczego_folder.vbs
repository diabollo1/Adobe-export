call back_script("C:\Users\Tomek\AppData\Roaming\Adobe\InDesign\Version 9.0\pl_PL\Scripts\Scripts Panel\all_to_pdf.vbs", "all_to_pdf.vbs")

set myInDesign = CreateObject("InDesign.Application")

temp_msg = ""

dim window, i
 
set window = createwindow()
	window.document.write "<html><body bgcolor=buttonface>Processing...<br><span id='output' style='font-size: 8px;'></span></body></html>"
	window.document.title = "do_pdf [" & myInDesign.Documents.Count * 2 & "]"
	window.resizeto 600, 900
	window.moveto 20, 20

temp_msg = temp_msg & "Do przerobienia: " & myInDesign.Documents.Count & " * 2 = " & myInDesign.Documents.Count * 2 & "<br>"

'----------------------------------------------------
'path = :
path = "D:\InD\zzz_pakiety"

	For myDocumentCounter = 1 To myInDesign.Documents.Count

		'MsgBox (myInDesign.Documents.Item(myDocumentCounter).Name)
		
		'path = myInDesign.Documents.Item(myDocumentCounter).filePath

		name = myInDesign.Documents.Item(myDocumentCounter).name
		name_bez = replace(myInDesign.Documents.Item(myDocumentCounter).name, ".indd", "")
		name_d = name_bez & "_d.pdf"
		name_small = name_bez & "_d_small.pdf"

		path_name_d = path & "\" & name_d
		path_name_small = path & "\" & name_small

		
		
		show temp_msg & myDocumentCounter * 2 - 1 & " O " & name_d & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 - 1 & "/" & myInDesign.Documents.Count * 2 & "]"
		
		'Export - 111
		myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_d, False, myInDesign.pdfExportPresets.item("111")
		
		temp_msg = temp_msg & myDocumentCounter * 2 - 1 & " V " & name_d & "<br>"
		show temp_msg & myDocumentCounter * 2 & " O " & name_small & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 & "/" & myInDesign.Documents.Count * 2 & "]"
		
		'Export - 111_small
		myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_small, False, myInDesign.pdfExportPresets.item("111_small")
		
		temp_msg = temp_msg & myDocumentCounter * 2 & " V " & name_small & "<br>"
		show temp_msg
		
	Next
	
'----------------------------------------------------
'path = tam gdzie plik
	For myDocumentCounter = 1 To myInDesign.Documents.Count

		'MsgBox (myInDesign.Documents.Item(myDocumentCounter).Name)
		
		path = myInDesign.Documents.Item(myDocumentCounter).filePath

		name = myInDesign.Documents.Item(myDocumentCounter).name
		name_bez = replace(myInDesign.Documents.Item(myDocumentCounter).name, ".indd", "")
		name_d = name_bez & "_d.pdf"
		name_small = name_bez & "_d_small.pdf"

		path_name_d = path & "\" & name_d
		path_name_small = path & "\" & name_small

		
		
		show temp_msg & myDocumentCounter * 2 - 1 & " O " & name_d & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 - 1 & "/" & myInDesign.Documents.Count * 2 & "]"
		
		'Export - 111
		myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_d, False, myInDesign.pdfExportPresets.item("111")
		
		temp_msg = temp_msg & myDocumentCounter * 2 - 1 & " V " & name_d & "<br>"
		show temp_msg & myDocumentCounter * 2 & " O " & name_small & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 & "/" & myInDesign.Documents.Count * 2 & "]"
		
		'Export - 111_small
		myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_small, False, myInDesign.pdfExportPresets.item("111_small")
		
		temp_msg = temp_msg & myDocumentCounter * 2 & " V " & name_small & "<br>"
		show temp_msg
		
	Next
'----------------------------------------------------



'Set WScript = CreateObject("WScript.Shell")

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


'myInDesign.ActiveDocument.Export idExportFormat.idPDFType, "D:\InD\syf\skrypt\aaa.pdf", False, myInDesign.pdfExportPresets.item("111")
'MsgBox ("done" & "<br>" & temp_msg)
