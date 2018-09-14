call back_script("C:\Users\Tomek\AppData\Roaming\Adobe\InDesign\Version 9.0\pl_PL\Scripts\Scripts Panel\Adobe-export\all_to_pdf.vbs", "all_to_pdf.vbs")

set myInDesign = CreateObject("InDesign.Application")
ile_dokumentow = myInDesign.Documents.Count

temp_msg = ""

dim window, i
 
set window = createwindow()
	window.document.write "<html><body bgcolor=buttonface>Processing...<br><span id='output' style='font-size: 10px;'></span></body></html>"
	window.document.title = "do_pdf [" & ile_dokumentow * 2 & "]"
	window.resizeto 600, 900
	window.moveto 20, 20

temp_msg = temp_msg & "Do przerobienia: " & ile_dokumentow & " * 2 = " & ile_dokumentow * 2 & "<br>"

'MsgBox(ile_dokumentow)

For myDocumentCounter = 1 To ile_dokumentow
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'checkPreflight(myInDesign.Documents.Item(myDocumentCounter))
	' Set profiles = myInDesign.Documents.Item(myDocumentCounter).PreflightProfiles
	' profileCount = profiles.Count
	' str = "Preflight profiles: "
	' For i = 1 To profileCount
		' If i > 1 Then
			' str = str & ", "
		' End If
		' str = str & profiles.Item(i).Name
	' Next
	' MsgBox(str)
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

	'MsgBox (myInDesign.Documents.Item(myDocumentCounter).Name)
	
	path = myInDesign.Documents.Item(myDocumentCounter).filePath
	name = myInDesign.Documents.Item(myDocumentCounter).name
	name_bez = replace(myInDesign.Documents.Item(myDocumentCounter).name, ".indd", "")
	name_d = name_bez & "_d.pdf"
	name_small = name_bez & "_d_small.pdf"

	path_name_d = path & "\" & name_d
	path_name_small = path & "\" & name_small
	
	set pref_111 = myInDesign.pdfExportPresets.item("111")
	set pref_111_small = myInDesign.pdfExportPresets.item("111_small")
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'msgbox myInDesign.pdfExportPreferences.useSecurity
	
	' pref_111.openDocumentPassword = "password"
	
	' For Each objProperty In myInDesign.pdfExportPresets.item("111")
	 
	  ' temp_msg = temp_msg & objProperty.Name & "<br>"
	  ' show temp_msg
	
	' Next
		' ascascasc asad
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'msgbox checkPreflight(myInDesign.Documents.Item(myDocumentCounter))
	if checkPreflight(myInDesign.Documents.Item(myDocumentCounter)) = "brak" then
	
		show temp_msg & myDocumentCounter * 2 - 1 & " O " & name_d & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 - 1 & "/" & ile_dokumentow * 2 & "]"
		myInDesign.pdfExportPreferences.useSecurity = False
		myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_d, False, pref_111
		'export_pdf()
		
		temp_msg = temp_msg & myDocumentCounter * 2 - 1 & " V " & name_d & "<br>"
		show temp_msg & myDocumentCounter * 2 & " O " & name_small & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 & "/" & ile_dokumentow * 2 & "]"
		
		'Export - 111_small
			myInDesign.pdfExportPreferences.useSecurity = True
			myInDesign.pdfExportPreferences.ChangeSecurityPassword = "romdruk.pl"
			myInDesign.pdfExportPreferences.disallowChanging = True 
			myInDesign.pdfExportPreferences.disallowCopying = True 
			myInDesign.pdfExportPreferences.disallowDocumentAssembly = True 
			myInDesign.pdfExportPreferences.disallowExtractionForAccessibility = True 
			myInDesign.pdfExportPreferences.disallowFormFillIn = True 
			myInDesign.pdfExportPreferences.disallowPrinting = True 
			myInDesign.pdfExportPreferences.disallowHiResPrinting = True 
			'myInDesign.pdfExportPreferences.disallowNotes = True 
			
		myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_small, False, pref_111_small
		
		temp_msg = temp_msg & myDocumentCounter * 2 & " V " & name_small & "<br>"
		show temp_msg
	else
		temp_msg = temp_msg & myDocumentCounter * 2 - 1 & " !!!ERROR!!! " & name_d & "<br>"
		temp_msg = temp_msg & myDocumentCounter * 2 & " !!!ERROR!!! " & name_small & "<br>"
		window.document.title = "do_pdf [" & myDocumentCounter * 2 & "/" & ile_dokumentow * 2 & "]"
		show temp_msg
	end if
	
Next

'Set WScript = CreateObject("WScript.Shell")

For i = 1 To 1000
	
	show temp_msg & "<br>" & i
	'WScript.Sleep 1000
	
Next

'save all
For myDocumentCounter = 1 To ile_dokumentow
	myInDesign.Documents.Item(myDocumentCounter).Save
Next

'close all
For myDocumentCounter = ile_dokumentow To 1 Step - 1
	'msgbox("myDocumentCounter:" & myDocumentCounter & "/" & ile_dokumentow)
	if checkPreflight(myInDesign.Documents.Item(myDocumentCounter)) = "brak" then
		myInDesign.Documents.Item(1).Close
	else
		ilosc_bledow = ilosc_bledow + 1
	end if
Next

if ilosc_bledow > 0 then
	temp_msg = temp_msg & "<br><br><h1>!!!ERROR!!!<br>Ilosc plikow z bledami: " & ilosc_bledow & "</h1><br>"
		show temp_msg
else
	window.close
end if



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

Function export_pdf()
	'myInDesign.Documents.Item(myDocumentCounter).Export idExportFormat.idPDFType, path_name_d, False, pref_111
	set myInDesign1 = CreateObject("InDesign.Application")
	myInDesign1.Documents.Item(1).Export idExportFormat.idPDFType, "aassssaaa", False, myInDesign1.pdfExportPresets.item("111")
End Function

function checkPreflight(plik)
	
	Rem Assume there is an document.
	Set myDoc = plik
	Rem Use the second preflight profile
	Set myProfile = myInDesign.PreflightProfiles.Item(2)
	Rem Process the doc with the rule
	Set myProcess = myInDesign.PreflightProcesses.Add(myDoc, myProfile)
	myProcess.WaitForProcess()
	results = myProcess.ProcessResults
		'MsgBox("aaa" & Left(results,4) & "bbb")
	Rem If Errors were found
	If Left(results,4) = "None" Then
		Rem Export the file to PDF. The "true" value selects to open the file after export.
		'myProcess.SaveReport("c:\PreflightResults.pdf")
		'MsgBox("brak")
		checkPreflight = "brak"
	else
		'MsgBox("jest")
		checkPreflight = "jest"
	End If
	Rem Cleanup
	myProcess.Delete()
  
End Function

'myInDesign.ActiveDocument.Export idExportFormat.idPDFType, "D:\InD\syf\skrypt\aaa.pdf", False, myInDesign.pdfExportPresets.item("111")
'MsgBox ("done" & "<br>" & temp_msg)
