' to execute: cscript //E:vbscript Nomedialize.vbs <folder>

Dim fso, pivotFolder, folder, sFolder
Dim output

' Functions first

Function recurseFolder (pivotFolder)
	Dim folders
	Dim localOutput
	
	foundPic = false
	foundMp3 = false
	For Each file In pivotFolder.Files
		If Right(file.Name, 3) = "mp3" Then 
			foundMp3 = true
		End If
		If Right(file.Name, 3) = "jpg" Or Right(file.Name, 4) = "jpeg" Or Right(file.Name, 3) = "png" Then 
			foundPic = true
		End If
	Next

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
		If Right(file.Name, 3) = "jpg" Or Right(file.Name, 4) = "jpeg" Or Right(file.Name, 3) = "png" Then 
			' MsgBox file.Path & ", " & pivotFolder.path & "\art"
			file.Move pivotFolder.path & "\art\"
		End If
	Next

	createArtNomediaFolder = pivotFolder.Path & ": Created art folder and .nomedia file" & vbCrLf
End Function


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

' Main code

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
