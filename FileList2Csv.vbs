Option Explicit

Dim fso, csv, debugMode
debugMode = True

Main()

Sub Main()
	Dim currentPath, scannedPath, outputFilePath, scanSubDir, separator
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	currentPath = fso.GetAbsolutePathName(".")

	''' user settings
	scannedPath = "c:\Windows\"
	outputFilePath = "c:\FileList2Csv\" & "FileList.csv"
	scanSubDir = False
	separator = "|"

	On Error Resume Next

	Const overWritten = True, unicodeMode = True
	Set csv = fso.CreateTextFile(outputFilePath, overWritten, unicodeMode)
	
	''' write csv headers
	csv.WriteLine "sep=" & separator
	csv.WriteLine Replace("Path,FileName,Size,Date Modified", ",", separator)
	ScanFiles scannedPath, scanSubDir, separator
	
	csv.Close

	Set csv = Nothing
	Set fso = Nothing
End Sub

Sub ScanFiles(scannedPath, includeSubDir, separatedBy)
	scannedPath = TrimStr(scannedPath)
	separatedBy = TrimStr(separatedBy)

	If Not IsEmptyOrNull(scannedPath) Then
		If fso.FolderExists(scannedPath) Then
			If IsEmptyOrNull(separatedBy) Then separatedBy = "|"

			includeSubDir = ParseBool(includeSubDir)
			
			Dim di, f, d, q, flagErr
			On Error Resume Next
			
			Set di = fso.GetFolder(scannedPath)
			q = Chr(34) ' double-quote
			flagErr = False
			
			Debug "scanning .. " & scannedPath
			
			For Each f In di.Files
				Debug vbTab & vbTab & f.Name
				
				csv.WriteLine q & f.ParentFolder & q & separatedBy & _
					q & f.Name & q & separatedBy & _
					f.Size & separatedBy & _
					CStr(f.DateLastModified) & _
					CStr(f.DateCreated)

				If Err.Number <> 0 Then
					flagErr = True
					Debug Err.Description
					On Error GoTo 0
					Exit For
				End If
			Next
			
			If includeSubDir And Not flagErr Then
				For Each d In di.Subfolders
					ScanFiles d.Path, includeSubDir, separatedBy
				Next
			End If
			
			Set di = Nothing
			
		End If
	End If
End Sub

Sub Debug(text)
	If debugMode Then
		If IsEmptyOrNull(text) Then
			WScript.Echo VbCrLf
		Else
			WScript.Echo text
		End If
	End If
End Sub

Function ParseBool(value) ' as boolean
	Dim ret : ret = False
	If Not IsEmptyOrNull(value) Then
		On Error Resume Next
		ret = CBool(value)
	End If
	ParseBool = ret
End Function

Public Function ParseInt(value) ' as integer
	Dim ret : ret = 0
	If Not IsEmptyOrNull(value) Then
		On Error Resume Next
		ret = CInt(value)
	End If
	ParseInt = ret
End Function

Function TrimStr(text) 'as string
	Dim ret : ret = ""
	If Not IsEmptyOrNull(text) Then
		text = Trim(text)
		If Not IsEmptyOrNull(text) Then ret = text
	End If
	TrimStr = ret
End Function

''' used for Object
Function IsNothing(obj)
	Dim ret : ret = False
	If IsObject(obj) Then
		ret = obj Is Nothing
	Else
		ret = IsEmptyOrNull(obj)
	End If
	IsNothing = ret
End Function

''' used for non-object
Function IsEmptyOrNull(checked_value)
	Dim ret : ret = False
	If IsObject(checked_value) Then
		ret = isNothing(checked_value)
	Else
		If IsEmpty(checked_value) Then
			ret = True 'check initilized or not
		ElseIf IsNull(checked_value) Then
			ret = True 'check no valid data
		ElseIf IsNumeric(checked_value) Then
			'ret = False
		ElseIf checked_value = "" Then
			ret = True 'check blank
		Else
			'ret = False
		End If
	End If
	IsEmptyOrNull = ret
End Function


