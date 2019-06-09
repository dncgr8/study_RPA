
' External functions used to make sure the excel instances are exited and processes killed
Declare Function EndTask Lib "user32.dll" (ByVal hWnd As IntPtr) As Integer
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
	   (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
Declare Function GetWindowThreadProcessId Lib "user32.dll" _
	   (ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
Declare Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Integer) As IntPtr

' The 'active' instance of excel.
' This is updated on any access of an instance (either adding a new
' instance or getting one)
'
Private CurrentInstance As Object

' Map of instances keyed against the handles which represent them.
'
Private HandleMap As Dictionary(Of Integer, Object)

' Map of handles keyed against the instances they represent.
' Here to ensure we don't assign multiple handles to a single
' instance and bring about a memory leak
'
Private InstanceMap As Dictionary(Of Object, Integer)

' Constructor - this just initialises the collections which map
' the excel instances to handles and vice versa.
Public Sub New()

	Me.HandleMap = New Dictionary(Of Integer, Object)()
	Me.InstanceMap = New Dictionary(Of Object, Integer)()

End Sub

' Gets the handle for a given instance
'
' If the instance is not yet held, then it is added to the 
' 	map and a handle is assigned to it. It is also set as the
' 	'current' instance, accessed with a handle of zero in the
' 	below methods.
'
' Either way, the handle which identifies the instance is returned
'
' @param Instance The instance for which a handle is required
'
' @return The handle of the instance
Protected Function GetHandle(Instance As Object) As Integer

	If Instance Is Nothing Then
		Throw New ArgumentNullException("Tried to add an empty instance")
	End If

	' Check if we already have this instance - if so, return it.
	If InstanceMap.ContainsKey(Instance) Then
		CurrentInstance = Instance
		Return InstanceMap(Instance)
	End If

	Dim key as Integer
	For key = 1 to Integer.MaxValue
		If Not HandleMap.ContainsKey(key)
			HandleMap.Add(key, Instance)
			InstanceMap.Add(Instance, key)
			Me.CurrentInstance = Instance
			Return key
		End If
	Next key

	Return 0

End Function


' Gets the instance corresponding to the given handle, setting
' 	the instance as the 'current' instance for future calls
'
' A value of 0 will provide the 'current' instance, which
' 	is set each time an instance is added or accessed.
'
' This will return Nothing if the given handle does not
' correspond to a registered instance, or if the current
' instance was closed and the reference has not been updated.
'
' @param Handle The handle representing the instance required,
' 		or zero to get the 'current' instance.
Protected Function GetInstance(Handle As Integer) As Object

	Dim Instance As Object = Nothing
	
	If Handle = 0 Then
		If CurrentInstance Is Nothing Then
			' Special case - getting the current instance when the
			' instance is not set, try and get a current open instance.
			' If none there, create a new one and assign a handle as if
			' CreateInstance() had been called
		'	Try
		'		Instance = GetObject(,"Excel.Application")
		'	Catch ex as Exception ' Not running
		'		Instance = Nothing
		'	End Try
		'	If Instance Is Nothing Then
				Create_Instance(Handle)
				' Instance = CreateObject("Excel.Application")
				' Force the instance into the maps.
				' GetHandle(Instance)
				' CurrentInstance should now be set.
				' If it's not, we have far bigger problems
		'	End If
		End If
		Return CurrentInstance
	End If

	Instance = HandleMap(Handle)
	If Not Instance Is Nothing Then
		CurrentInstance = Instance
	End If
	Return Instance

End Function


' Close the instance with the given handle, not saving any work, making
' sure that it is removed from this VBO's collection of instances
'
' @param Handle The handle representing the instance to close
Protected Sub CloseInstance(Handle As Integer)
	CloseInstance(Handle, False)
End Sub

' Close the instance with the given handle, saving the work as specified.
'
' @param Handle The handle representing the instance
'
' @param SaveWorkbooks True to save the workbooks before quitting the instance,
' 		False to discard any changes
Protected Sub CloseInstance(Handle As Integer, SaveWorkbooks As Boolean)
	
	Dim Instance As Object = Nothing
	If Handle = 0 AndAlso CurrentInstance Is Nothing Then
		Throw New NullReferenceException("Tried to close nonexistent current instance")
	ElseIf Handle = 0 ' Current Instance - reset it
		Handle = GetHandle(CurrentInstance) ' We need the handle to remove from HandleMap
		Instance = CurrentInstance
		CurrentInstance = Nothing
	Else
		Instance = GetInstance(Handle)
	End If

	Me.HandleMap.Remove(Handle)
	Me.InstanceMap.Remove(Instance)

	Instance.DisplayAlerts = False ' Hide alerts
	
	' First close all the workbooks and the workbooks collection
	Dim wbs as Object = Instance.Workbooks
	If wbs IsNot Nothing Then
		For Each Workbook As Object In wbs
			Workbook.Close(SaveWorkbooks)
		Next
		wbs.Close()
	End If
	
	' Try quitting - sometimes this is enough
	Instance.Quit()
	
	' Try and force a com object release - this might quit excel for us.
	System.Runtime.InteropServices.Marshal.ReleaseComObject(Instance)

	' Now if the com object has released the RCW, we need to stop
	' We'll know because if we try and get the version and it fails
	' then the COM object has been cleaned up.
	' If so, end the proc now - we have to assume the instance is gone
	
	Dim Ver as Double = 0.0
	Try
		Ver = Val(Instance.Version)
	Catch ex as Exception
		' Not got the version - assuming cleared up
		Return
	Finally
		SetLastError(0) ' If any errors have occurred thus far, clear them
	End Try

	' Now it's the messy stuff to try and find the excel instance and nuke
	' it from orbit. It's the only way to be sure.

	' The window handle for the excel instance
	Dim hwnd As IntPtr = IntPtr.Zero
	' Later versions of excel expose the window handle
	If Val(Ver) >= 10 Then _
		hwnd = New IntPtr(CType(Instance.Parent.Hwnd, Integer))
		
	' If the window handle isn't set, must be an earlier version of excel
	' Use FindWindow to find the window with the GUID that we set in it on creation
	If IntPtr.Equals(hwnd, IntPtr.Zero) Then _
		hwnd = FindWindow(Nothing, Instance.Caption)
	
	' If the window handle is still zero, the instance must already be closed
	If Not IntPtr.Equals(hwnd, IntPtr.Zero) Then
	
        ' Get the process ID for the window we have
		Dim resp, procId as Integer
        resp = GetWindowThreadProcessId(hwnd, procId)
		
        If procId = 0 Then ' can’t get Process ID
            If EndTask(hwnd) = 0 Then ' EndTask returns a bool - 0 = False
				Throw New ApplicationException("Failed to close Excel Instance.")
			End If
        Else ' We have a process ID - use it to kill excel
			Dim proc As Process = Process.GetProcessById(procId)
			' Try clicking the 'X'
			proc.CloseMainWindow()
			proc.Refresh()

			If Not proc.HasExited Then
				proc.Kill()	' Last resort - kill it with fire
			End If	
		End If	
	End If

End Sub

' Creates a new workbook in the instance represented by the given handle
'
' @param Handle The handle of the instance on which the workbook should be held
'
' @return The workbook object that was created.
Protected Function NewWorkbook(Handle as Integer) As Object

	Dim wb as Object = GetInstance(handle).Workbooks.Add()

	' Create a new Worksheet?
	if wb.Worksheets.Count = 0 Then
		wb.Sheets.Add().Activate()
	Else ' Just use the first sheet
		wb.Sheets(1).Activate()
	End If
	
	Return wb
	
End Function

' Gets the workbook in the given instance with the given name.
'
' @param Handle The handle representing the instance which holds the workbook
'
' @param Name The name of the workbook on the instance
'
' @return The object representing the workbook defined
Protected Function GetWorkbook(Handle As Integer, Name as String) As Object

	Dim wb as Object = Nothing
	If String.IsNullOrEmpty(Name) Then
		wb = GetInstance(Handle).ActiveWorkbook
		If wb Is Nothing ' We need to create a deafult workbook
			wb = NewWorkbook(Handle)
		End If
		Return wb
	Else
		Return GetInstance(Handle).Workbooks(Name)
	End If

End Function

' Gets the worksheet specified by the given handle, workbook name and
' worksheet name. If no such sheet is available, this will create a
' new one and return that
'
' @param Handle The handle identifying the instance which should be
' 		acted on
'
' @param WorkbookName The name of the workbook within the instance
'
' @param WorksheetName The name of the worksheet required
'
' @return The sheet object representing the sheet with the given name
Protected Function GetWorksheet(Handle As Integer, _
		WorkbookName As String, _
		WorksheetName As String) As Object

	Return GetWorksheet(Handle,WorkbookName,WorksheetName,True)

End Function

' Gets the worksheet specified by the given handle, workbook name and
' worksheet name. If no such sheet is available, this will create a
' new one or return Nothing, depending on the given flag
'
' @param Handle The handle identifying the instance which should be
' 		acted on
'
' @param WorkbookName The name of the workbook within the instance
'
' @param WorksheetName The name of the worksheet required
'
' @param CreateIfNotExists True to create the worksheet if it doesn't
' 		exist; False to return Nothing if it doesn't exist.
'
' @return The sheet object representing the sheet required or Nothing
'		if no such sheet exists and CreateIfNotExists was False
Protected Function GetWorksheet(Handle As Integer, _
		WorkbookName As String, _
		WorksheetName As String, _
		CreateIfNotExists As Boolean) As Object

	Dim wb As Object = GetWorkbook(Handle, WorkbookName)
	
	If (String.IsNullOrEmpty(WorksheetName)) Then

		Dim ws as Object = wb.ActiveSheet
		If ws Is Nothing Then
			Return wb.Sheets.Add()
		Else
			Return ws
		End If

	Else
		Dim sheets as Object = wb.Sheets
		If sheets IsNot Nothing Then
			For Each sheet as Object in sheets
				If sheet.Name = WorksheetName Then _
					Return sheet
			Next
		End If
		' Didn't find the sheet...
		If CreateIfNotExists Then

			Dim sheet as Object = sheets.Add()
			sheet.Name = WorksheetName
			Return sheet

		End If
		' Nothing else we can do - return nowt
		Return Nothing

	End If

End Function

' Gets the next cell, relative to a given cell in a given direction
' Note that if the cell is at a boundary, then the same cell is returned.
'
' @param cell : The cell to use as a base cell
'
' @param strDir : one of "L", "R", "U", "D" representing a direction from
' 		the given cell to move in.
'
' @return : The cell object representing the 'next cell' 
Protected Function GetNextCell(cell as Object, strDir as String) As Object
	
	Try
		Select Case strDir
			Case "L"
				cell = cell.Offset(0,-1)
			Case "R"
				cell = cell.Offset(0,1)
			Case "U"
				cell = cell.Offset(-1,0)
			Case "D"
				cell = cell.Offset(1,0)
		End Select
	Catch ex As Exception
		' A COM Exception is thrown if the cell is at a boundary and the offset
		' would break that boundary
	End Try
	
	Return cell

End Function

Function ParseDelimSeparatedVariables(data as String, delimStr As String, schema As DataTable, firstRowIsHeader As Boolean) As DataTable

		Const SchemaColumnName As String = "Column Name"
		Const DefaultState As Integer = 0
		Const Instring As Integer = 1
		Const FirstQuote As Integer = 2

		Const Quote As Char = """"c
		If delimStr.Length = 0 Then delimStr = ","
		If delimStr.Length <> 1 Then Throw New Exception("Delimiter must be a single character")
		
		Dim delim As Char = delimStr(0)

		Dim state As Integer = DefaultState
		Dim firstRow As Boolean = True
		If Not firstRowIsHeader Then
			firstRow = False
		End If
		Dim columnValue As New StringBuilder()

		Dim emptySchema As Boolean = schema Is Nothing OrElse schema.Rows.Count = 0
		Dim outputCollection As New DataTable()

		If Not emptySchema Then
			For Each schemaRow As DataRow In schema.Rows
				Dim colName As String = schemaRow(SchemaColumnName).ToString
				outputCollection.Columns.Add(colName, GetType(String))
			Next
		End If

		Dim row As DataRow = Nothing
		Dim colIndex As Integer = 0
		Using sw As New StringReader(data)
			While True
				Dim line As String = sw.ReadLine()
				If line Is Nothing Then Exit While

				' If we're not processing a CRLF in the middle of a string, we want to move
				' onto the next row; if we are, we keep the current row and column since we're
				' still writing to that 'cell'.
				If state <> Instring Then
					row = outputCollection.NewRow
					colIndex = 0
				End If
				For Each ch As Char In line
					Select Case ch
						Case delim ' ie. 'ch' is the specified delimiter - "," or "\t"
							Select Case state
								Case Instring
									columnValue.Append(delim)
								Case Else ' Covers 'default' and 'first quote'.
									If firstRow Then
										If emptySchema Then
											If firstRowIsHeader Then
												Dim colName As String = columnValue.ToString
												outputCollection.Columns.Add(colName, GetType(String))
											End If
										Else
											If firstRowIsHeader Then
												Dim schemaColName As String = schema.Rows(colIndex).Item(SchemaColumnName).ToString
												Dim colName As String = columnValue.ToString
												If colName <> schemaColName Then
													Throw New Exception("Column name mismatch. Column '" & colName & "' dosen't match schema name of '" & schemaColName & "'")
												End If
											End If
										End If
									Else
										If Not firstRowIsHeader Then
											Dim colName As String = "Column " & outputCollection.Columns.Count
											outputCollection.Columns.Add(colName, GetType(String))
										End If
										row.Item(colIndex) = columnValue.ToString
									End If

									columnValue.Length = 0
									state = DefaultState
									colIndex += 1
							End Select
						Case Quote
							Select Case state
								Case FirstQuote
									state = Instring
									columnValue.Append(Quote)
								Case Instring
									state = FirstQuote
								Case Else
									' If we find a quote in the middle of a non-quoted cell, it's
									' a literal quote; otherwise (ie. at the start of a cell), it
									' means the cell value is wrapped - go into 'Instring' state
									If columnValue.Length > 0 Then
										columnValue.Append(Quote)
									Else
										state = Instring
									End If
							End Select
						Case Else
							columnValue.Append(ch)
					End Select
				Next

				If firstRow Then
					If emptySchema Then
						If firstRowIsHeader Then
							Dim colName As String = columnValue.ToString
							outputCollection.Columns.Add(colName, GetType(String))
						End If
					Else
						If firstRowIsHeader Then
							Dim schemaColName As String = schema.Rows(colIndex).Item(SchemaColumnName).ToString
							Dim colName As String = columnValue.ToString
							If colName <> schemaColName Then
								Throw New Exception("Column name mismatch. Column '" & colName & "' dosen't match schema name of '" & schemaColName & "'")
							End If
						End If
					End If
					firstRow = False
					columnValue.Length = 0
					state = DefaultState
				Else
					' If we're still in the middle of the string we want to include the CRLF in the
					' actual value that we're writing and leave the state at 'Instring'
					If state = Instring Then
						columnValue.Append(vbCrLf)
					Else
						If Not firstRowIsHeader Then
							Dim colName As String = "Column " & outputCollection.Columns.Count
							outputCollection.Columns.Add(colName, GetType(String))
						End If
						row.Item(colIndex) = columnValue.ToString
						outputCollection.Rows.Add(row)
						columnValue.Length = 0
						state = DefaultState
					End If
				End If

			End While

		End Using
		
		Return outputCollection

End Function