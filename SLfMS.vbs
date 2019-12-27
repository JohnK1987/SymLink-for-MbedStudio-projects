''This script can manage the MbedOs library for possibility make a new project Offline 
''and also make a project without a additional download of the library approx 1GB for each new project.
''That is achieved thank to use of Symbolic or hard links in the Windows applied to the MbedOs library folder and its mbed-os.lib file.

''Run as Admin
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
End If

''A title for all boxes
Title = "SLforMbedStudio"
''All objects
Dim FSO:set FSO = CreateObject("Scripting.FileSystemObject")
Dim WS: set WS = CreateObject("Wscript.Shell")
Dim SA: set SA = CreateObject("Shell.Application")
''Store current folder of the script
CurrentFolder = FSO.GetParentFolderName(WScript.ScriptFullName)
''Check source
SourceState = SourceCheck(CurrentFolder)

If(SourceState = True) Then
	''Menu
	Do	
		Question = InputBox("What would you like to do ?" &vbcr&vbcr&_
							"1 - New project from soure"&vbcr&_
							"2 - Link source to existing project"&vbcr&_
							"3 - Delete links from a project"&vbcr&_
							"4 - Delete a whole project"&vbcr&_
							"5 - Show all SymLinks",_
							Title,_ 
							1)
		Select Case Question
			Case 1 			Call NewMbedProject()
			Case 2 			Call ReplaceMbedSource()
			Case 3 			Call DeleteLinksFromProject()
			Case 4 			Call DeleteProject()
			case 5 			Call WS.Run("cmd.exe /C dir/AL /S C:\ | find ""SYMLINK""")
			Case vbEmpty	''Exit
			Case Else		MsgBox "Wrong command! Try it again...",0,Title
		End Select
	Loop While(Question  < 1 and Question  > 5)	
Else
	''Wrong set of source MsgBox
	MsgBox 	"A source setting is not correct." & SourceState &vbcr&vbcr&_
			"Please be so kind and make new empty project with a name SOURCE in the MbedStudio. "&vbcr&_
			"Then move the SOURCE project to a new separete folder (for example C:\MbedSource\)."&vbcr&_
			"And this script move to the same directory.",_
			0, Title
End If

Function SourceCheck(CurFolder)
	Dim FolderPath, FilePath
	Folderpath = CurFolder & "\Source\Mbed-os"
	FilePath = FolderPath & "\mbed.h"
	If FSO.FolderExists(FolderPath) Then 
		If FSO.FileExists(FilePath) Then 
			SourceCheck = True
		Else	
			SourceCheck = "Mbed.h not found in " & Folderpath & "."
		End If 
	Else
		SourceCheck = "Source folder not found in " & CurFolder & "."
	End If 
End Function

Sub NewMbedProject()
	Dim MbedWorkspaceFolder, NewProjectName
	''Folder selector/picker
	MbedWorkspaceFolder = SelectFolder("C:\", "Please select a MbedStudio workspace folder where you want to make a new project from Source!")
	If (MbedWorkspaceFolder <> vbNull) Then
		Do
			state = False
			NewProjectName = ""
			NewProjectName = InputBox("Please paste a new project name",_
									Title,_ 
									"NewProject")
			If (Not IsEmpty(NewProjectName)) and (Len(NewProjectName) <> 0) and NewProjectName <> " " and FSO.FolderExists(MbedWorkspaceFolder & "\" & NewProjectName)= false Then
				''Make new Folder for project
				Destination = MbedWorkspaceFolder & "\" & NewProjectName
				FSO.CreateFolder Destination
				WScript.Sleep 100
				If FSO.FolderExists(Destination) Then 
					''Check Files from the Source
					For Each objFolder In FSO.GetFolder(CurrentFolder & "\Source\").Files
						If objFolder.Name = "mbed-os.lib" Then
							''Symlink file
							SA.ShellExecute "cmd.exe", "/C mklink /h " & Chr(34) & Destination & "\mbed-os.lib" & Chr(34) & " " & Chr(34) & objFolder.Path & Chr(34), , "runas", 1
						Else
							''Copy file
							FSO.CopyFile objFolder.Path, Destination & "\"
						End if
					Next
					''Check Folders from the Source
					For Each objFolder In FSO.GetFolder(CurrentFolder & "\Source\").SubFolders
						If objFolder.Name = "mbed-os" Then
							''Symlink folder
							SA.ShellExecute "cmd.exe", "/C mklink /d " & Chr(34) & Destination & "\mbed-os" & Chr(34) & " " & Chr(34) & objFolder.Path & Chr(34), , "runas", 1
						Else
							''Copy folder
							FSO.CopyFolder objFolder.Path, Destination & "\"
						End if
					Next
				Else
					MsgBox "Error the requested folder in destination " & Destination & " was not create!",0,Title
				End if
				State = True
			ElseIf	(Len(NewProjectName) = 0) Then
				Exit Sub
			ElseIf (Not IsEmpty(NewProjectName)) and (Len(NewProjectName) <> 0) and NewProjectName <> " " and FSO.FolderExists(MbedWorkspaceFolder & "\" & NewProjectName)= true then
				MsgBox "This projet name already exist! Try it again...",0,Title
			Else
				MsgBox "Wrong projet name! Try it again...",0,Title
			End if
			
		Loop While (State <> True)
	End if
End sub

Sub ReplaceMbedSource()
	Dim MbedWorkspaceFolder
	MbedWorkspaceFolder = SelectFolder("C:\", "Please select a project folder in the MbedStudio workspace folder, where you want to replace standart library with the linked Source!")
	If (MbedWorkspaceFolder <> vbNull) Then
		Destination = MbedWorkspaceFolder & "\mbed-os"
		If FSO.FolderExists(Destination) Then
			SA.ShellExecute "cmd.exe", "rd " & Chr(34) & Destination & Chr(34), , "runas", 1
		End If 
		SA.ShellExecute "cmd.exe", "/C mklink /d " & Chr(34) & Destination & Chr(34) & " " & Chr(34) & CurrentFolder & "\Source\mbed-os" & Chr(34), , "runas", 1
		
		Destination = MbedWorkspaceFolder & "\mbed-os.lib"
		If FSO.FileExists(Destination) Then
			SA.ShellExecute "cmd.exe", "del " & Chr(34) & Destination & Chr(34), , "runas", 1
		End If 			
		SA.ShellExecute "cmd.exe", "/C mklink /h " & Chr(34) & Destination & Chr(34) & " " & Chr(34) & CurrentFolder & "\Source\mbed-os.lib" & Chr(34), , "runas", 1	
		MsgBox "The Mbed-os library folder and mbed-os.lib file from " & MbedWorkspaceFolder & " were replaced with linked sources.",0,Title
	End if
End sub

Sub DeleteLinksFromProject()
	Dim MbedWorkspaceFolder
	MbedWorkspaceFolder = SelectFolder("C:\", "Please select a project folder in the MbedStudio workspace folder, where you want to delete the linked Source!")
	If (MbedWorkspaceFolder <> vbNull) Then
		Destination = MbedWorkspaceFolder & "\mbed-os"
		If FSO.FolderExists(Destination) Then
			SA.ShellExecute "cmd.exe", "/C rd " & Chr(34) & Destination & Chr(34), , "runas", 1
		End If 
		
		Destination = MbedWorkspaceFolder & "\mbed-os.lib"
		If FSO.FileExists(Destination) Then
			SA.ShellExecute "cmd.exe", "/C del " & Chr(34) & Destination & Chr(34), , "runas", 1
		End If 	
		MsgBox "The linked sources were deleted from the project " & MbedWorkspaceFolder & ".",0,Title
	End if
End sub

Sub DeleteProject()
	Dim MbedWorkspaceFolder
	MbedWorkspaceFolder = SelectFolder("C:\", "Please select a project folder in the MbedStudio workspace folder, what you want to delete!")
	If (MbedWorkspaceFolder <> vbNull) Then
		Destination = MbedWorkspaceFolder
		If FSO.FolderExists(Destination) Then
			SA.ShellExecute "cmd.exe", "/C rd " & Chr(34) & Destination & Chr(34) & "/q /s", , "runas", 1
			MsgBox "The project " & MbedWorkspaceFolder & " was deleted.",0,Title
		Else
			MsgBox "The project folder " & MbedWorkspaceFolder & " not exist.",0,Title
		End If 	
	End if
End sub

Function SelectFolder( myStartFolder, myDialogStr )
	' Modified original code 
	' Written by Rob van der Woude
	' https://www.robvanderwoude.com/vbstech_ui_selectfolder.php
    ' Standard housekeeping
    Dim objFolder, objItem
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull
    ' Create a dialog object
    Set objFolder = SA.BrowseForFolder( 0, myDialogStr, 0, myStartFolder )
    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path
    ' Standard housekeeping
    Set objFolder = Nothing
    On Error Goto 0
End Function

set FSO = Nothing
set WS = Nothing
set SA = Nothing