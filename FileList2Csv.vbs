Option Explicit

Dim fso, csv, debugMode
Dim currentPath, scannedPath, outputFilePath, scanSubDir, separator

' example: 
' cscript filelist2csv.vbs
' cscript filelist2csv.vbs c:\windows
' cscript filelist2csv.vbs "c:\program files" ".\filelist.csv"
' cscript filelist2csv.vbs c:\windows ".\filelist.csv" false
' cscript filelist2csv.vbs c:\windows ".\filelist.csv" false ","

Main()

Sub Main()
	debugMode = True
	
	ReadArgs()
	Start()
End Sub

Sub ReadArgs()
	Dim args : Set args = WScript.Arguments

	Set fso = CreateObject("Scripting.FileSystemObject")
	currentPath = fso.GetAbsolutePathName(".")
	
	debug args.Count
	
	Select Case args.Count
		Case 0 ' use this case to hardcode the path and call vbs with none argument
			scannedPath = currentPath '"c:\windows" 
			outputFilePath = currentPath & "\FileList.csv"
			scanSubDir = False
			separator = "|"
		Case 1
			scannedPath = TrimStr(args(0))
			outputFilePath = currentPath & "\FileList.csv"
			scanSubDir = False
			separator = "|"
		Case 2
			scannedPath = TrimStr(args(0))
			outputFilePath = TrimStr(args(1))
			scanSubDir = False
			separator = "|"
		Case 3
			scannedPath = TrimStr(args(0))
			outputFilePath = TrimStr(args(1))
			scanSubDir = ParseBool(TrimStr(args(2)))
			separator = "|"
		Case Else
			scannedPath = TrimStr(args(0))
			outputFilePath = TrimStr(args(1))
			scanSubDir = ParseBool(TrimStr(args(2)))
			separator = Trimstr(args(3))
	End Select
	
	WScript.Echo "FileList2Csv script"
	WScript.Echo "-------------------"
	WScript.Echo "Syntax: cscript FileList2Csv.vbs 'path\to\be\scanned' 'path\to\csv\export' include_subdir csv_separator"
	
	Debug vbTab & "- scannedPath = '" & scannedPath & "'"
	Debug vbTab & "- outputFilePath = '" & outputFilePath & "'"
	Debug vbTab & "- scanSubDir = " & CStr(scanSubDir)
	Debug vbTab & "- separator = '" & separator & "'"

	Dim chkPath(1)
	chkPath(0) = CheckPath(scannedPath)
	chkPath(1) = CheckPath(outputFilePath)

	Debug VbCrLf & vbTab & "'" & scannedPath & "' .. " & CStr(chkPath(0))
	Debug vbTab & "'" & outputFilePath & "' .. " & CStr(chkPath(1))

	If chkPath(0) And chkPath(1) Then
		Debug VbCrLf
	Else
		Debug vbTab & "*** error found in arguments ***"
	End If
End Sub

Sub Start()
	On Error Resume Next
	
	Const overWritten = True, unicodeMode = True
	Set csv = fso.CreateTextFile(outputFilePath, overWritten, unicodeMode)

	''' write csv headers
	csv.WriteLine "sep=" & separator
	csv.WriteLine Replace("Path,FileName,Size,Date Modified,Date Crreated", ",", separator)
	
	ScanFiles scannedPath, scanSubDir, separator

	csv.Close
	
	Set csv = Nothing
	Set fso = Nothing
End Sub

Sub ScanFiles(pathToScan, includeSubDir, separatedBy)
	pathToScan = TrimStr(pathToScan)
	separatedBy = TrimStr(separatedBy)
	
	If Not IsEmptyOrNull(pathToScan) Then
		If fso.FolderExists(pathToScan) Then
			If IsEmptyOrNull(separatedBy) Then separatedBy = "|"
			
			includeSubDir = ParseBool(includeSubDir)

			Dim di, f, d, q, chkPath
			On Error Resume Next

			Set di = fso.GetFolder(pathToScan)
			q = Chr(34) ' double-quote
			chkPath = False

			Debug "scanning .. '" & pathToScan & "'"

			For Each f In di.Files
				Debug Space(16) & "+-- " & f.Name '& vbtab & f.Size

				csv.WriteLine q & f.ParentFolder & q & separatedBy & _
					q & f.Name & q & separatedBy & _
					f.Size & separatedBy & _
					CStr(f.DateLastModified) & separatedBy & _
					CStr(f.DateCreated)
				
				If Err.Number <> 0 Then
					chkPath = True
					Debug Err.Description
					On Error GoTo 0
					Exit For
				End If
			Next

			If includeSubDir And Not chkPath Then
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

Function TrimStr(text) 'as string
	Dim ret : ret = ""
	If Not IsEmptyOrNull(text) Then
		text = Trim(text)
		If Not IsEmptyOrNull(text) Then ret = text
	End If
	TrimStr = ret
End Function

Function CheckPath(fullPath) 'as boolean
	Dim ret : ret = False
	fullPath = TrimStr(fullPath)
	If fullPath <> "" Then
		If fso.FolderExists(fullPath) Then
			'Debug "'" & fullPath & "' is folder."
			ret = True
		Else
			If fso.FileExists(fullPath) Then
				'Debug "'" & fullPath & "' is file."
				ret = True
			Else
				Dim p, i, j
				i = InStr(fullPath, "\")
				j = InStrRev(fullPath, "\")
				
				If j > i + 1 Then
					p = Left(fullPath, j)
				Else
					p = fullPath
				End If

				ret = fso.FolderExists(p)
			End If
		End If

	End If
	CheckPath = ret
End Function

Function ParseBool(value) ' as boolean
	Dim ret : ret = False
	value = TrimStr(value)
	If value <> "" Then
		On Error Resume Next
		If IsNumeric(value) Then
			ret = Iif(value = "0", False, True)
		Else
			ret = CBool(value)
		End If
	End If
	ParseBool = ret
End Function

Function IsNothing(obj)
	Dim ret : ret = False
	If IsObject(obj) Then
		ret = obj Is Nothing
	Else
		ret = IsEmptyOrNull(obj)
	End If
	IsNothing = ret
End Function

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