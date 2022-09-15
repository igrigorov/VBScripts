' to execute: cscript //E:vbscript Nomedialize.vbs <folder>

Dim fso, pivotFolder, folder, sFolder
Dim output

' Functions first


' Returns true if the file is of matching extension, according to the mode
' 	mode = "m" = music
'	mode = "p" = picture
Function checkExtension(filename, mode)
	filenameL = LCase(filename)
	ext3 = Right(filenameL, 3)
	ext4 = Right(filenameL, 4)

	checkExtension = False

	If mode = "m" Then
		If ext3 = "mp3" Then 
			checkExtension = true
		End If
	Else
		If ext3 = "jpg" Or ext4 = "jpeg" Or ext3 = "png" Then 
			checkExtension = true
		End If
	End If
End Function


' Recursive function, executing the main course of work:
' - Creation of .nomedia files
' - Moving of picture files, if mixed with music files, along with
' - Creation of /art subfolder
Function recurseFolder (pivotFolder)
	Dim folders
	Dim localOutput
	
	foundPic = false
	foundMp3 = false
	' check if the folder contains file of X type; only switch-on flags, otherwise they get reset wrongly
	For Each file In pivotFolder.Files
		If foundMp3 = false Then
			foundMp3 = checkExtension(file.Name, "m")
		End If
		If foundPic = false Then
			foundPic = checkExtension(file.Name, "p")
		End If
	Next
	
	MsgBox "Folder" & pivotFolder.Name & ": music=" & foundMp3 & ": pics=" & foundPic

	Set folders = pivotFolder.SubFolders
	foundSubfolders = folders.Count > 0

	If foundPic = true Then
		If foundMp3 = true Then
			localOutput = localOutput & createArtNomediaFolder(pivotFolder)
		ElseIf foundSubfolders = true Then
			localOutput = pivotFolder.Path & ": Found subfolders under picture-only folder... WHAT TO DO?" & vbCrLf
		Else
			localOutput = localOutput & createNomediaFile(pivotFolder)
		End If
	End If
	
	For Each folder In folders
		If folder.Name <> "art" Then
			localOutput = localOutput & recurseFolder(folder)
		End If
	Next
	
	recurseFolder = localOutput
End Function


' Creates an /art subfolder, with a .nomedia file in it and puts all picture files in the current folder in it
Function createArtNomediaFolder (pivotFolder)
	On Error Resume Next

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set newFolder = fso.CreateFolder(pivotFolder.path & "\art")
	If Err.Number <> 0 Then
		createArtNomediaFolder = pivotFolder.path & ": art subfolder already exists. Error: " & Err.Description & vbCrLf
		Exit Function
	End If

	Set newFile = fso.CreateTextFile(pivotFolder.path & "\art\.nomedia", True, True)
	newFile.Close

	' move pic files in the art folder
	For Each file In pivotFolder.Files
		If checkExtension(file.Name, "p") = true Then 
			' MsgBox file.Path & ", " & pivotFolder.path & "\art"
			file.Move pivotFolder.path & "\art\"
		End If
	Next

	createArtNomediaFolder = pivotFolder.Path & ": Created art folder and .nomedia file" & vbCrLf
End Function


' Creates an .nomedia file in the current folder
Function createNomediaFile (pivotFolder)
	On Error Resume Next

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set newFile = fso.CreateTextFile(pivotFolder.path & "\.nomedia", True, True)
	If Err.Number <> 0 Then
		createNomediaFile = pivotFolder.path & ": .nomedia already exists. Error: " & Err.Description & vbCrLf
		Exit Function
	Else
		newFile.Close
	End If

	createNomediaFile = pivotFolder.Path & ": Created .nomedia file" & vbCrLf
End Function


' Main code with some logging

sFolder = Wscript.Arguments.Item(0)
If sFolder = "" Then
	Wscript.Echo "No Folder parameter was passed"
	Wscript.Quit
End If

output = ""
localOutput = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Set outputFile = fso.CreateTextFile(sFolder & "\Nomedialize.output.txt", True, True)
Set pivotFolder = fso.GetFolder(sFolder)

output = recurseFolder(pivotFolder)

outputFile.WriteLine(output)
outputFile.Close

MsgBox ".nomedia file(s) created under art folder(s). Check Nomedialize.output.txt for log."
