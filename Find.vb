Function Find (ByVal Handle As Integer, ByVal WorkbookName As String, 
               ByVal Worksheet As String, ByVal RangeReference As String,
               ByVal FindWhat As String, ByVal LookIn As Integer,
               ByVal WholeWords As Boolean, ByVal MatchCase As Boolean,
               ByVal SearchByRow As Boolean, ByVal SearchPrevious As Boolean)As String

'Excel 操作　オブジェクト

Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
Dim xlBooks As Microsoft.Office.Interop.Excel.Workbooks = Nothing
Dim xlBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

'Excel アプリケーション　生成

xlApp = GetInstance(Handle)
xlBooks = xlApp.WorkBooks
xlBook = xlBooks.(WorkbookName)

'シートを選択する

xlSheet = CType(xlBook.Sheets(WorksheetName), Microsoft.Office.Interop.Excel.Worksheet)
xlSheet.Select()


Dim referenceRange As Microsoft.Office.Interop.Excel.Range = Nothing

If "".Equals (RangeReference)Then
    referenceRange = xlSheet.UsedRange

Else 
    referenceRange = xlSheet.Range(RangeReference)

End If
`If referenceRange.Cell.Count = 1 Then
`   Return FindWhat.Equals("" & referenceRange.Cells(1).Value)
`End If

Dim result As Microsoft.Office.Interop.Excel.Range = Nothing
result = referenceRange.Find(FindWhat,LookIn:=LookIn, LookAt:=Microsoft.VisualBasic.IIf(WholeWords,1,2),
                             SearchOrder:=Microsoft.VisualBasic.IIf(SearchByRow,1,2),
                             SearchDirection:=Microsoft.VisualBasic.IIf(SearchPrevious,1,2),MatchCase:=MatchCase)

Dim resultNext As Microsoft.Office.Interop.Excel.Range = result
    
    If result Is Nothing Then 
        Throw New Exception("指定値が見つかりません。")

    End If 

    If SearchPrevious Then 
        Do 
            resultNext = referenceRange.Find(FindWhat,After:=resultNext,
            LookIn:=LookIn, LookAt:=Microsoft.VisualBasic.IIf(WholeWords,1,2),
            SearchOrder:=Microsoft.VisualBasic.IIf(SearchByRow,1,2),
            SearchDirection:=Microsoft.VisualBasic.IIf(SearchPrevious,2,1),
            MatchCase:=MatchCase)

            If resultNext.Address = referenceRange.Cells(referenceRange.Cells.Count).Address Then
                return resultNext.Address
            End If
        
        Loop until result.Address.Equals(resultNext.Address)

    
    Else
        Do
            resultNext = referenceRange.Find(indWhat,After:=resultNext,
            LookIn:=LookIn, LookAt:=Microsoft.VisualBasic.IIf(WholeWords,1,2),
            SearchOrder:=Microsoft.VisualBasic.IIf(SearchByRow,1,2),
            SearchDirection:=Microsoft.VisualBasic.IIf(SearchPrevious,2,1),
            MatchCase:=MatchCase)

            If resultNext.Address = referenceRange.Cells(1).Address Then
                    Return resultNext.Address
            End If 

        Loop until  result.Address.Equals(resultNext.Address)

    End If 

    Return result.Address

End Function 
'insert this code into Blueprism VBO