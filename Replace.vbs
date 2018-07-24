Seaspan()

Function Seaspan()
	
	dim repeat 
	dim delrepeat
	dim repeat2
	dim count 
	dim htmlData
	dim DDdata
	dim readBinary
	dim search 
	dim search2 
	dim fso
	dim i 
	dim k
	dim l 
	dim folderspec
	dim j 
	dim strPath
	dim list
	dim last 
	dim first
	dim lastnum
	dim del
	dim found
	dim index
	dim CompArr
	CompArr = Array()
	
	l=1
	k=1
	count = 1
	index = 0
	i = 1
	first = 1
	j = 2
	found = false
	
	repeat = "<div class=""dropdown"">"
	repeat2 = "<div id=""myDropdown1"" class=""dropdown-content"">"
	
	search = "<a href=""#"" class=""scene"" data-id="""
	search2 = """>"
	last = "<div id=""titleBar"">"
	
	Set list = CreateObject("System.Collections.ArrayList")
	set fso = CreateObject("Scripting.FileSystemObject")

	folderspec = fso.GetParentFolderName(WScript.ScriptFullName)
	Set f = fso.GetFolder(folderspec)
	strPath = f & "\index.html"
	
	readBinary = read(strPath,fso)
	i = Instr(readBinary,search)
	j = Instr(i,readBinary,search2)
	first = i
	lastnum = Instr(i,readBinary,last)
	
	if Instr(readBinary,repeat) <> 0 then
		'delrepeat = Mid(readBinary,Instr(readBinary,repeat),(Instr(readBinary,repeat2)+len(repeat2))-Instr(readBinary,repeat))
		'msgbox(delrepeat)
		'readBinary = Replace(readBinary,delrepeat,"")
		
	end	if
	
	while i <> 0
		i = Instr(i+1,readBinary,search)
		if i <> 0 then	
			j = Instr(i,readBinary,search2)
			'msgbox(Mid(readBinary,i+Len(search),j-(i+Len(search))))
			'a(count)=Mid(readBinary,i+Len(search),j-(i+Len(search)))
			count = count+1
		end if			
	Wend

	'wscript.echo list.Count
	dim a()
	redim CompArr(1)
	redim a(count)
	i = 1
	j = 2
	count = 0
	
	while i <> 0
		
		i = Instr(i+1,readBinary,search)
		l=1
		k=1
		
		if i <> 0 then	
			
			j = Instr(i,readBinary,search2)
			a(count)=Mid(readBinary,i+Len(search),j-(i+Len(search)))
			
			while mid(a(count),Instr(Lcase(a(count)),"r")-k,1) = "-"
				k = k + 1
			wend
			
			l=k
			
			while mid(a(count),Instr(Lcase(a(count)),"r")-l,1) <> "-"
				l=l+1
			wend
			
			'msgbox(mid(a(count),Instr(Lcase(a(count)),"r")-l+1,(Instr(Lcase(a(count)),"r")-k) - (Instr(Lcase(a(count)),"r")-l)))
			found = false
			index = 0
			
			while index<UBound(CompArr)
				
				If InStr(CompArr(index),mid(a(count),Instr(Lcase(a(count)),"r")-l+1,(Instr(Lcase(a(count)),"r")-k) - (Instr(Lcase(a(count)),"r")-l))) <> 0 then
					found = true
				end if		
				
				index = index + 1
				
			wend
			
			if found = false then			
				ReDim Preserve CompArr(UBound(CompArr)+1)
				CompArr(UBound(CompArr)-1) = mid(a(count),Instr(Lcase(a(count)),"r")-l+1,(Instr(Lcase(a(count)),"r")-k) - (Instr(Lcase(a(count)),"r")-l)) 
			end if
			
			count = count + 1
			
		end if			
	Wend
	
	For i = LBound(CompArr)+1 To UBound(CompArr)-1
		htmlData = htmlData + "  <div class=""dropdown"">" + chr(13)+chr(10) &_
                "<button onclick=""myFunction(" + chr(39)+ "myDropdown" + CStr(i) + chr(39)+")"" class=""dropbtn"">" + CompArr(i) + " </button>" +chr(13)+chr(10) &_
                "<div id=""myDropdown" + CStr(i) + """ class=""dropdown-content"">" + chr(13)+chr(10)
		DDdata = ""		
		For j = LBound(a) To UBound(a)-1
			if InStr(a(j),CompArr(i)) > 0 then
				DDdata = DDdata + " <a href=""#"" class=""scene"" data-id="""+ a(j) +"""> "+ a(j) +"</a> " +chr(13)+chr(10)
			end if	
				
		Next
		htmlData = htmlData + DDdata
	
		htmlData = htmlData + " </div></div> " 

		'MsgBox(CompArr(i))
	Next
	htmlData = htmlData + "</ul></div>"
    
	del = Mid(readBinary,first,lastnum-first)
	readBinary = Replace(readBinary,del,htmlData)
	msgbox(LBound(a))
	strPath = f & "\index.html"
	call writeBinary(readBinary, strPath)
			

	
End Function 

Function read(strPath,fso)

	Dim oFile: Set oFile = fso.GetFile(strPath)
	'msgbox(strPath)
	
	If IsNull(oFile) Then 
		MsgBox("File not found: " & strPath)
		Exit Function	
	End if

	With oFile.OpenAsTextStream()
		readBinary = .Read(oFile.Size)
		.Close
	End With
	
	read = readBinary
	
End Function

Function writeBinary(strBinary, strPath)

	Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")

    ' below lines pupose: checks that write access is possible!
    Dim oTxtStream

    On Error Resume Next
    Set oTxtStream = oFSO.createTextFile(strPath)

    If Err.number <> 0 Then MsgBox(Err.message) : Exit Function
    On Error GoTo 0

    Set oTxtStream = Nothing
    ' end check of write access

    With oFSO.createTextFile(strPath)
        .Write(strBinary)
        .Close
    End With

End Function