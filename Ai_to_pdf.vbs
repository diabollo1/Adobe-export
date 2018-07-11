' exportFileAsPDF(D:\InD\zzz_pakiety)
' Sub exportFileAsPDF (dest)
' Set appRef = CreateObject("Illustrator.Application")
' Set saveOptions = CreateObject("Illustrator.PDFSaveOptions")
' saveOptions.ColorCompression = 6 'aiJPEGHigh
' saveOptions.Compatibility = 5 'aiAcrobat5
' Set frontDocument = appRef.ActiveDocument
' Cali frontDocument.SaveAs (dest, saveOptions)
' End Sub


	'dest = "D:\InD\zzz_pakiety"




set myAi = CreateObject("Illustrator.Application")


For myDocumentCounter = 1 To myAi.Documents.Count
	myAi.Documents.Item(myDocumentCounter).Save
Next


For myDocumentCounter = 1 To myAi.Documents.Count

	set aktualnyPlik = myAi.Documents.Item(myDocumentCounter)
	path = aktualnyPlik.Path
	name = aktualnyPlik.name
	name_bez = replace(aktualnyPlik.name, ".ai", "")
	name_d = name_bez & "_d.pdf"
	name_small = name_bez & "_d_small.pdf"

	path_name_d = path & "\" & name_d
	path_name_small = path & "\" & name_small
	
	Set saveOptions = CreateObject("Illustrator.PDFSaveOptions")
	
	saveOptions.PDFPreset = "111"
		call aktualnyPlik.SaveAs(path_name_d, saveOptions)
	saveOptions.PDFPreset = "111_small"
		call aktualnyPlik.SaveAs(path_name_small, saveOptions)
	
Next

' For myDocumentCounter = 1 To myAi.Documents.Count
	' set aktualnyPlik = myAi.Documents.Item(myDocumentCounter)
	' aktualnyPlik.Save
' Next

For myDocumentCounter = 1 To myAi.Documents.Count
	set aktualnyPlik = myAi.Documents.Item(1)
	aktualnyPlik.Close(1)
Next

