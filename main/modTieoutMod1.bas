Attribute VB_Name = "modTieoutMod1"
Sub ExcelFile_Get_dart(ByRef filePath As String)
    Dim i As Integer
    Dim NewFileName As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim cellAddress As String
    Dim sht As String
    Dim keyword1 As String, keyword2 As String
    

    ' 파일 경로를 지정합니다.
    'FilePath = "C:\Users\yeyi\Downloads\데이터분석\XBRL완전성확인\SKT 별도감사보고서_2022_0310_공시_최종-DOC.htm"
    
    ' 엑셀 파일 열기
    Workbooks.Open (filePath)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '코드 실행 중 셀 계산 방지
    
    ' 엑셀 파일 형식과 새로운 이름으로 저장
    'NewFileName = ActiveSheet.name + Format(Date, "-mm-dd") + Format(Time, "_hhmmss")
    'NewFileName = ThisWorkbook.Path + "\work\TieOut" + Format(Date, "-mm-dd") + Format(Time, "_hhmmss")
    NewFileName = "C:\xampp\htdocs\work\TieOut" + Format(Date, "-mm-dd") + Format(Time, "_hhmmss")
    ActiveWorkbook.SaveAs Filename:=NewFileName, FileFormat:=xlWorkbookDefault

    ' 작업하려는 워크시트를 지정합니다. 원하는 워크시트의 이름으로 변경하세요.
    Set ws = ActiveWorkbook.Sheets(1)
    ws.name = "DART"
    
    ' 찾을 키워드를 지정합니다.
    keyword1 = "재무상태표"
    keyword2 = "연결재무상태표"
    
    ' 역순으로 워크시트의 각 셀을 확인하고 키워드를 찾습니다.
    For Each cell In ws.UsedRange
        If Replace(cell.Value, " ", "") = keyword1 Or Replace(cell.Value, " ", "") = keyword2 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword1 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword2 Then '반기재무상태표처럼 앞2글자가 붙는 경우 추가고려
            ' 키워드를 찾은 경우 해당 셀의 주소를 기록합니다.
            cellAddress = cell.Address
            Exit For
        End If
    Next cell
    
    ' 키워드를 찾은 경우 해당 셀의 윗줄을 모두 삭제합니다.
    If cellAddress <> "" Then
        ' 키워드를 찾은 행의 한 줄 위의 행까지 삭제합니다.
        ws.Rows("1:" & ws.Range(cellAddress).Row - 1).Delete
        ws.Cells(1, 1).EntireRow.Insert Shift:=xlDown
        
    End If
    
    
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete ' 이미지 Shape 삭제
    Next
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Range(ws.Cells(i, "A"), ws.Cells(i, "Z")).Hyperlinks.Count > 0 And ws.Cells(i, "A").Value <> "주석" Then
           ws.Range("A" & (i) & ":A" & lastRow + 10).EntireRow.Delete
        End If
    Next i
    
    
    Dim rowCount As Integer
    Dim emptyRowCount As Integer

    rowCount = ActiveSheet.UsedRange.Rows.Count
    emptyRowCount = 0
    
    ' 뒤에서부터 검사하여 빈 행 카운트
    For i = rowCount To 1 Step -1
        If WorksheetFunction.CountA(ActiveSheet.Rows(i)) = 0 Then
            emptyRowCount = emptyRowCount + 1
            
            ' 빈 행 삭제
            If emptyRowCount > 1 Then
                ActiveSheet.Rows(i).Delete Shift:=xlUp
            End If
        
        Else
            emptyRowCount = 0
            
        End If
    Next i
        
    sht = "DART"
    SetTableName sht
    SetStyle
    
    ' 열린 엑셀 파일을 저장하고 닫습니다.
    ActiveWorkbook.Save
    'ActiveWorkbook.Close
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExcelFile_Get_dart2(ByRef filePath As String)

    Dim wsSource As Worksheet ' 두 번째 엑셀 파일의 시트
    Dim wsTarget As Worksheet ' 첫 번째 엑셀 파일의 대상 시트 ("XBRL" 시트)
    Dim wbSource As Workbook  ' 두 번째 엑셀 파일
    Dim sht As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '코드 실행 중 셀 계산 방지
    
    ' 두 번째 파일 경로를 지정합니다.
    'Dim FilePath As String
    'FilePath = "C:\Users\yeyi\Downloads\데이터분석\XBRL완전성확인\사업보고서_69_00121941_대상주식회사_0711_검사완료.xls"
    
    ' 첫 번째 엑셀 파일의 "XBRL" 시트를 지정합니다.
    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.name = "XBRL"
    Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    ' 두 번째 엑셀 파일 열기
    Set wbSource = Workbooks.Open(filePath)
    
    ' 두 번째 엑셀 파일의 시트를 순회하며 "기본정보" 시트를 제외하고 데이터를 복사합니다.
    For Each wsSource In wbSource.Sheets
        ' "기본정보" 시트를 제외하고 데이터를 복사합니다.
        wsSource.UsedRange.Copy wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Offset(2, 0)
    Next wsSource
    
    ' 두 번째 엑셀 파일을 닫습니다.
    wbSource.Close SaveChanges:=False
    
    
    ' 찾을 키워드를 지정합니다.
    keyword1 = "재무상태표"
    keyword2 = "연결재무상태표"
    
    ' 역순으로 워크시트의 각 셀을 확인하고 키워드를 찾습니다.
    For Each cell In wsTarget.UsedRange
        If Replace(cell.Value, " ", "") = keyword1 Or Replace(cell.Value, " ", "") = keyword2 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword1 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword2 Then '반기재무상태표처럼 앞2글자가 붙는 경우 추가고려
            ' 키워드를 찾은 경우 해당 셀의 주소를 기록합니다.
            cellAddress = cell.Address
            Exit For
        End If
    Next cell
    
    ' 키워드를 찾은 경우 해당 셀의 윗줄을 모두 삭제합니다.
    If cellAddress <> "" Then
        ' 키워드를 찾은 행의 한 줄 위의 행까지 삭제합니다.
        wsTarget.Rows("1:" & wsTarget.Range(cellAddress).Row - 1).Delete
        wsTarget.Cells(1, 1).EntireRow.Insert Shift:=xlDown
        
    End If
    
    
    For i = wsTarget.Shapes.Count To 1 Step -1
        wsTarget.Shapes(i).Delete ' 이미지 Shape 삭제
    Next
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If wsTarget.Range(wsTarget.Cells(i, "A"), wsTarget.Cells(i, "Z")).Hyperlinks.Count > 0 And wsTarget.Cells(i, "A").Value <> "주석" Then
           wsTarget.Range("A" & (i) & ":A" & lastRow + 10).EntireRow.Delete
        End If
    Next i
    
    Dim rowCount As Integer
    Dim emptyRowCount As Integer
    
    'Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    rowCount = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    emptyRowCount = 0
    
    ' 뒤에서부터 검사하여 빈 행 카운트
    For i = rowCount To 1 Step -1
        If WorksheetFunction.CountA(wsTarget.Rows(i)) = 0 Then
            emptyRowCount = emptyRowCount + 1
    
            ' 빈 행 삭제
            If emptyRowCount > 1 Then
                wsTarget.Rows(i).Delete Shift:=xlUp
            End If
    
        Else
            emptyRowCount = 0
    
        End If
    Next i
        
    wsTarget.Activate
    
    sht = "XBRL"
    SetTableName sht
    SetStyle
    
    wsTarget.name = "DART_2nd"
    
    ActiveWorkbook.Save
    'ActiveWorkbook.Close

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 나머지 코드 작성
End Sub


Sub ExcelFile_Get_xbrl(ByRef filePath As String)

    Dim wsSource As Worksheet ' 두 번째 엑셀 파일의 시트
    Dim wsTarget As Worksheet ' 첫 번째 엑셀 파일의 대상 시트 ("XBRL" 시트)
    Dim wbSource As Workbook  ' 두 번째 엑셀 파일
    Dim sht As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '코드 실행 중 셀 계산 방지
    
    ' 두 번째 파일 경로를 지정합니다.
    'Dim FilePath As String
    'FilePath = "C:\Users\yeyi\Downloads\데이터분석\Tie-Out_XBRL완전성확인\사업보고서_69_00121941_대상주식회사_0711_검사완료.xls"
    
    ' 첫 번째 엑셀 파일의 "XBRL" 시트를 지정합니다.
    ActiveWorkbook.Sheets.Add
    ActiveSheet.name = "XBRL"
    Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    ' 두 번째 엑셀 파일 열기
    Set wbSource = Workbooks.Open(filePath)
    
    note_count = 1
    
    ' 두 번째 엑셀 파일의 시트를 순회하며 "기본정보" 시트를 제외하고 데이터를 복사합니다.
    For Each wsSource In wbSource.Sheets
        ' "기본정보" 시트를 제외하고 데이터를 복사합니다.
        If wsSource.name <> "기본정보" Then
            ' 데이터를 붙여넣을 위치를 결정합니다.
            
            If wsSource.name Like "*주석*" Then
            
                Dim lastRow As Long
                lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
                
                ' 마지막 행의 셀이 병합되었는지 확인합니다.
                Dim mergedRows As Long
                If wsTarget.Cells(lastRow, "A").MergeCells Then
                    ' 병합된 경우, 병합된 셀의 범위를 찾습니다.
                    mergedRows = wsTarget.Cells(lastRow, "A").MergeArea.Rows.Count
                    ' 병합된 셀의 마지막 행 번호를 계산합니다.
                    insertRow = lastRow + mergedRows + 1
                Else
                    ' 병합되지 않은 경우, 마지막 행 번호에서 +2를 합니다.
                    insertRow = lastRow + 2
                End If
                
                ' 시트 이름을 붙여넣을 위치에 입력합니다.
                wsTarget.Cells(insertRow, 1).Value = note_count & ". " & Split(wsSource.name, "주석 - ")(1)
                
                note_count = note_count + 1
            End If
        
            wsSource.UsedRange.Copy wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Offset(2, 0)
        End If
    Next wsSource
    
    ' 두 번째 엑셀 파일을 닫습니다.
    wbSource.Close SaveChanges:=False
    
    
    Dim rowCount As Integer
    Dim emptyRowCount As Integer
    
    'Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    rowCount = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    emptyRowCount = 0
    
    ' 뒤에서부터 검사하여 빈 행 카운트
    For i = rowCount To 1 Step -1
        If WorksheetFunction.CountA(wsTarget.Rows(i)) = 0 Then
            emptyRowCount = emptyRowCount + 1
    
            ' 빈 행 삭제
            If emptyRowCount > 1 Then
                wsTarget.Rows(i).Delete Shift:=xlUp
            End If
    
        Else
            emptyRowCount = 0
    
        End If
    Next i
        
    wsTarget.Activate

    sht = "XBRL"
    SetTableName sht
    SetStyle
    
    ActiveWorkbook.Save
    'ActiveWorkbook.Close

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 나머지 코드 작성
End Sub


Sub SetStyle()
Dim rng As Range
Dim column As Range
Dim originalWidth As Double
Dim newWidth As Double
Dim lastColumn As Long

    ' 셀 서식 지정
    With Cells
        .VerticalAlignment = xlCenter   ' 텍스트 맞춤 세로(가운데)
        .Interior.Pattern = xlNone  ' 배경 제거
        .WrapText = False   ' 줄바꿈 없음
        .RowHeight = 16.5   ' 줄 높이(StandardHeight값)
        .Font.Size = 9     ' 폰트 사이즈 변경
        .IndentLevel = 0    ' 들여쓰기 없음
        .Font.name = "나눔바른고딕 Light"    ' 폰트 지정

        ' 공백 문자로 변경
        .Replace What:=ThisWorkbook.Sheets(1).Range("A1").Value, Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        ' 숫자로 변경 -> 공백과 함께 있는 경우가 있어 공백 제거 후 변경하도록 위치 조정
        '.Replace What:="-", Replacement:="0", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    End With
    
    ActiveSheet.UsedRange.ColumnWidth = 20
    Columns("A:A").ColumnWidth = 30
        
           
    Set rng = ActiveSheet.UsedRange ' 또는 적절한 범위를 지정
    
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.HorizontalAlignment = xlCenter Then
                If InStr(cell.Value, ".") > 0 Then
                    cell.NumberFormat = "#,##0.00;[Red](#,##0.00);-"
                Else
                    cell.NumberFormat = "#,##0;[Red](#,##0);-"
                End If
            Else
                If InStr(cell.Value, ".") > 0 Then
                    cell.NumberFormat = "#,##0.00;[Red](#,##0.00);-    "
                Else
                    cell.NumberFormat = "#,##0;[Red](#,##0);-    "
                End If
            End If
        End If
    Next cell


    '-'표시 숫자'0'으로 수정 및 헤더부분 바탕색 입력
    For Each cell In rng
       If Len(Trim(cell.Value)) = 1 And Trim(cell.Value) = "-" Then
            cell.Value = "0"
        End If
       
    Next cell



    ' 선택 Clear
    Range("A1").Select
    

End Sub

'======================================
' Table이름 지정 -> Ref.Check에서 내용 표시 시 사용(주석)
'======================================

Sub SetTableName(ByRef sht As String)

    Dim tableRange As Range, curRange As Range, preRange As Range, nextRange As Range, nextRange2 As Range
    Dim ws As Worksheet
    Dim noteStart As Boolean
    
    Set ws = ActiveWorkbook.ActiveSheet
        
    rowASheetCount = ActiveSheet.UsedRange.Cells.Rows.Count + 10
    colASheetCount = ActiveSheet.UsedRange.Cells.Columns.Count

    ' Name 번호
    NoteNumber = 1
    TableNumber = 1
    noteStart = False
    
    Dim keywords As Variant
    keywords = Array("재무상태표", "손익계산서", "자본변동표", "현금흐름표")


    ' 테이블 확인
    For Each curRange In ActiveSheet.UsedRange.Cells
        
        ' 주석 번호 확인
        If curRange.column = 1 And CStr(curRange.Value) Like NoteNumber & ".*" Then
            ' 주석 Name 정의
            Dim NoteTitle As String
            NoteTitle = Trim(curRange.Value)
            ws.Names.Add name:="NOTE" & NoteNumber, RefersTo:=curRange
            
            NoteNumber = NoteNumber + 1
            
            noteStart = True
        End If
    
        If noteStart = False Then
            If curRange.Borders(xlEdgeLeft).LineStyle = xlNone And curRange.Borders(xlEdgeRight).LineStyle = xlNone And _
               curRange.Borders(xlEdgeTop).LineStyle = xlNone And curRange.Borders(xlEdgeBottom).LineStyle = xlNone Then
               
                ' 셀의 텍스트에서 모든 공백을 제거합니다.
                Dim textWithoutSpaces As String
                textWithoutSpaces = Replace(curRange.Value, " ", "")
        
                ' 특정 키워드가 포함되어 있는지 검사합니다.
                For Each keyword In keywords
                    If InStr(textWithoutSpaces, keyword) > 0 Then
                        NoteTitle = keyword
                        Exit For ' 키워드를 찾으면 루프를 종료합니다.
                    End If
                Next keyword
                
            End If
        End If
                
        'curRange.Select
        ' 테이블 시작 위치
        If rowTableStart = 0 And curRange.Row > 1 _
            And curRange.Borders(xlEdgeLeft).LineStyle = xlContinuous And curRange.Borders(xlEdgeTop).LineStyle = xlContinuous Then
            ' 이전 구간으로 시작 위치 확정
            Set preRange = Cells(curRange.Row - 1, curRange.column)
            If preRange.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone _
                And preRange.Borders(xlEdgeRight).LineStyle = xlLineStyleNone Then
                rowTableStart = curRange.Row
                colTableStart = curRange.column
            End If
        End If
        
        ' 테이블 종료 위치
        If rowTableStart > 0 And curRange.Borders(xlEdgeRight).LineStyle = xlContinuous _
            And curRange.Borders(xlEdgeBottom).LineStyle = xlContinuous Then
            
            curRange.Value = RTrim(curRange.Value)
            
            ' 다음 구간 조회하여 종료 위치 확정
            Set nextRange = Cells(curRange.Row + 1, curRange.column + 1)
            
            If nextRange.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone _
                And nextRange.Borders(xlEdgeTop).LineStyle = xlLineStyleNone Then
                rowTableEnd = curRange.Row
                colTableEnd = curRange.column
                
                Set nextRange2 = Cells(curRange.Row + 1, curRange.column)
                If nextRange2.Borders(xlEdgeLeft).LineStyle = xlContinuous _
                And nextRange2.Borders(xlEdgeTop).LineStyle = xlContinuous _
                And nextRange2.Borders(xlEdgeRight).LineStyle = xlLineStyleNone _
                And nextRange2.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone _
                And nextRange2.MergeCells = False Then
                Set newRow = Cells(rowTableEnd + 1, 1).EntireRow
                newRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                With Cells(rowTableEnd, 1).EntireRow.Borders(xlEdgeLeft)
                    .LineStyle = xlNone
                End With
                With Cells(rowTableEnd, 1).EntireRow.Borders(xlEdgeRight)
                    .LineStyle = xlNone
                End With
                End If
                
            Else
                Set nextRange2 = Cells(curRange.Row, curRange.column + 1)
                If nextRange2.Borders(xlEdgeLeft).LineStyle = xlContinuous _
                And nextRange2.Borders(xlEdgeBottom).LineStyle = xlContinuous _
                And nextRange2.Borders(xlEdgeRight).LineStyle = xlLineStyleNone _
                And nextRange2.Borders(xlEdgeTop).LineStyle = xlLineStyleNone _
                And nextRange2.MergeCells = False Then
                rowTableEnd = curRange.Row
                colTableEnd = curRange.column
                Set newRow = Cells(rowTableEnd + 1, 1).EntireRow
                newRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                With Cells(rowTableEnd, 1).EntireRow.Borders(xlEdgeLeft)
                    .LineStyle = xlNone
                End With
                With Cells(rowTableEnd, 1).EntireRow.Borders(xlEdgeRight)
                    .LineStyle = xlNone
                End With

                End If
                
            End If
            

        End If
        
        ' 테이블인 경우 처리
        If rowTableStart > 0 And colTableEnd > 0 And rowTableStart = rowTableEnd And colTableStart = colTableEnd Then
        
            rowTableStart = 0
            colTableEnd = 0

        ElseIf rowTableStart > 0 And colTableEnd > 0 Then
    
            ' 테이블 선택
            Set tableRange = ws.Range(Cells(rowTableStart, colTableStart), Cells(rowTableEnd, colTableEnd))
            
            ' 테이블 Name 정의
            Dim nameNote As String
            nameNote = sht & "_TABLE" & TableNumber
            
            'tableRange.name = nameNote '통합문서범위에서 name 저장
            ws.Names.Add name:=nameNote, RefersTo:=tableRange '해당 시트범위에서 name 저장
            ws.Names(nameNote).Comment = NoteTitle
            
            TableNumber = TableNumber + 1

            ' 테이블 상/하단 라인 추가
            If Not IsEmpty(Cells(rowTableEnd + 1, 1).Value) Then
                Cells(rowTableEnd + 1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            End If
            
            ' CheckRange 초기화
            rowTableStart = 0
            colTableEnd = 0
        End If

    Next
    
End Sub


Sub MatchTables() 'control As IRibbonControl)
    Dim wsXBRL As Worksheet
    Dim wsDART As Worksheet
    Dim name As name
    Dim matchTable As name
    Dim bestMatchScore As Double
    Dim uniqueNumbers As Object
    Dim rngXBRL As Range
    Dim rng As Range
    Dim rngB As Range
   
    Set wb = ActiveWorkbook

    ' 모든 이름을 확인하고 깨진 이름을 삭제합니다.
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        On Error Resume Next
        If nm.RefersToRange Is Nothing Then
            ' 깨진 이름인 경우 삭제합니다.
            nm.Delete
        End If
        On Error GoTo 0
    Next i
    
    
    
    ' "XBRL" 시트가 있는지 확인
    On Error Resume Next ' 시트를 찾을 수 없는 경우 에러 처리
    Set wsXBRL = ActiveWorkbook.Sheets("XBRL")
    On Error GoTo 0 ' 에러 처리를 원래대로 복구
    
    ' "XBRL" 시트가 없는 경우 "DART" 시트로 설정
    If wsXBRL Is Nothing Then
        Set wsXBRL = ActiveWorkbook.Sheets("DART_2nd")
    End If
    
    'Set wsXBRL = ActiveWorkbook.Sheets("XBRL")
    Set wsDART = ActiveWorkbook.Sheets("DART")
    
    ' 'XBRL' 시트의 테이블을 순회합니다.
    'For Each name In wsXBRL.Names
    For Each wsXBRL In ActiveWorkbook.Sheets
        If wsXBRL.name <> "DART" Then
            For Each nameXBRL In wsXBRL.Names
                bestMatchScore = 0
                
                Set uniqueNumbers = CreateObject("Scripting.Dictionary")
                Set rng = nameXBRL.RefersToRange
                
                Dim cell As Range
                For Each cell In rng
                    If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then ' And cell.Value <> 0 Then
                        If Not uniqueNumbers.Exists(cell.Value) Then
                            uniqueNumbers(cell.Value) = 1
                        End If
                    End If
                Next cell
                
                If uniqueNumbers.Count <> 0 Then
                
                    For Each name In wsDART.Names
                        If name.name Like "*TABLE*" Then
                            Dim matchScore As Double
                            matchScore = CompareTables(uniqueNumbers, name, wsDART, wsXBRL)
                            
                            Set rng_dart = wsDART.Range(name)
    
                            For Each cell In rng_dart
                                If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
                                    numericCellCount = numericCellCount + 1
                                    ' 조건에 따라 다른 작업 수행
                                    If cell.Interior.Color = RGB(255, 255, 153) Or cell.Interior.Color = RGB(255, 217, 102) Then
                                        coloredCellCount = coloredCellCount + 1
                                    End If
                                End If
                            Next cell
                            
                            ' 일치하는 'DART' 테이블을 선택합니다.
                            If matchScore > bestMatchScore And (coloredCellCount / numericCellCount) < 0.8 Then '노란색셀의 비율이 80% 이상이면 이미 다른 테이블과 매칭이 끝난 데이블로 판단
                                bestMatchScore = matchScore
                                ' 해당 'DART' 테이블을 저장하거나 작업을 수행합니다.
                                Set matchTable = name
                            End If
                        End If
                        
                    Next
                    
                    Set rng = nameXBRL.RefersToRange
                    
                    For Each c In rng
                    
                        c.Interior.ColorIndex = xlNone
                        
                        'If IsNumeric(c.Value) And Not IsEmpty(c.Value) And c.Value <> 0 Then ' Check if the cell value is numeric and not empty or zero
                        If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then  ' Check if the cell value is numeric and not empty or zero
                            Set rngB = matchTable.RefersToRange
                            c.Interior.Color = RGB(255, 0, 0) '빨간색
                            c.Font.Color = RGB(255, 255, 255)
                            
                            For Each d In rngB
                                If IsNumeric(d.Value) And Not IsEmpty(d.Value) And Not d.HasFormula Then
                                   
                                    If c.Value = d.Value Then ' If the values in the cells are the same

                                        d.Formula = "=" & wsXBRL.name & "!" & c.Address(0, 0)
                                        d.Interior.Color = RGB(255, 255, 153) '노란색
                                        'd.Font.Color = RGB(0, 0, 0)
                                        c.Interior.ColorIndex = xlNone
                                        c.Font.Color = RGB(0, 0, 0)
            
                                        Exit For ' Exit the loop since a link has been added to cell d
                                        
                                    ElseIf c.Value * -1 = d.Value Then

                                        d.Formula = "=" & wsXBRL.name & "!" & c.Address(0, 0)
                                        d.Interior.Color = RGB(255, 217, 102) '짙은 노란색
                                        'd.Font.Color = RGB(0, 0, 0)
                                        c.Interior.ColorIndex = xlNone
                                        c.Font.Color = RGB(0, 0, 0)
            
                                        Exit For ' Exit the loop since a link has been added to cell d
                                    
                                    End If
                                    
                                End If
                            Next d
                        End If
                        
                    Next c
                    
                End If
 
            Next
        
        End If
    Next
    
    '여기
    ' 대상 시트 생성 및 설정
    Set lastSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Set tgtWs = ActiveWorkbook.Sheets.Add(After:=lastSheet)
    tgtWs.name = "Tables_Recon"
    
    tgtWs.Cells(1, "A").Value = "Sheet"
    tgtWs.Cells(1, "A").Font.Bold = True ' 글꼴을 굵게 설정
    tgtWs.Cells(1, "A").HorizontalAlignment = xlCenter ' 가운데 정렬
        
    tgtWs.Cells(1, "B").Value = "Note"
    tgtWs.Cells(1, "B").Font.Bold = True ' 글꼴을 굵게 설정
    tgtWs.Cells(1, "B").HorizontalAlignment = xlCenter ' 가운데 정렬
        
    tgtWs.Cells(1, "C").Value = "Missing value"
    tgtWs.Cells(1, "C").Font.Bold = True ' 글꼴을 굵게 설정
    tgtWs.Cells(1, "C").HorizontalAlignment = xlCenter ' 가운데 정렬
    tgtRow = 2
        
    For Each wsXBRL In ActiveWorkbook.Sheets
        If wsXBRL.name <> "DART" Then
            For Each nameXBRL In wsXBRL.Names
                If nameXBRL.name Like "*TABLE*" Then
                    Set rng = nameXBRL.RefersToRange
                    
                    For Each c In rng
                        If c.Interior.Color = RGB(255, 0, 0) Then '빨간색
                            tgtWs.Cells(tgtRow, "A").Value = wsXBRL.name
                            tgtWs.Cells(tgtRow, "B") = nameXBRL.Comment
                            tgtWs.Cells(tgtRow, "C").Formula = "=" & wsXBRL.name & "!" & c.Address
                            tgtRow = tgtRow + 1
                        End If
                    Next c
                End If
            Next nameXBRL
        End If
    Next wsXBRL

    For Each wsDART In ActiveWorkbook.Sheets
        If wsDART.name = "DART" Then
            For Each nameDART In wsDART.Names
                If nameDART.name Like "*TABLE*" Then
                    Set rng = nameDART.RefersToRange
                    
                    For Each c In rng
                        If Not IsEmpty(c.Value) And IsNumeric(c.Value) And c.Interior.ColorIndex = xlNone Then
                            tgtWs.Cells(tgtRow, "A").Value = wsDART.name
                            tgtWs.Cells(tgtRow, "B") = nameDART.Comment
                            tgtWs.Cells(tgtRow, "C").Formula = "=" & wsDART.name & "!" & c.Address
                            tgtRow = tgtRow + 1
                        End If
                    Next c
                End If
            Next nameDART
        End If
    Next wsDART

    SetStyle
    
    ActiveWorkbook.Save
    
End Sub

Function CompareTables(uniqueNumbers As Object, name As name, wsDART As Worksheet, wsXBRL As Worksheet) As Double
    ' 'XBRL' 테이블과 'DART' 시트 간의 숫자값 유사성을 평가하는 함수
    Dim matchScore As Double
    Dim rng As Range
    Dim uniqueNumbers_dart As Object
        
    Set rng = name.RefersToRange
    
    matchScore = 0
    
    ' 'XBRL' 테이블에서 고유한 숫자값을 추출합니다.

    
    ' 'DART' 시트에서 해당 숫자값을 비교합니다.
    Dim numMatches As Double
    numMatches = 0
    
    Set uniqueNumbers_dart = CreateObject("Scripting.Dictionary")
    
    For Each cell In rng    'dart탭 테이블
        'If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) And cell.Value <> 0 And Not uniqueNumbers_dart.Exists(cell.Value) Then
        '    uniqueNumbers_dart(cell.Value) = 1
        '    If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) And cell.Value <> 0 Then
        '        If uniqueNumbers.Exists(cell.Value) Then
        '            numMatches = numMatches + 1
        '        End If
        '    End If
        'End If
        
        If Not cell.HasFormula Then ' 공식이 아닌 경우만 처리
            If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then 'And cell.Value <> 0 Then
                If Not uniqueNumbers_dart.Exists(cell.Value) Then  '중복되는 값은 고려하지 않음.
                    uniqueNumbers_dart(cell.Value) = 1
                    
                    If uniqueNumbers.Exists(cell.Value) Then  'uniqueNumbers 안에도 중복값이 없으므로 중복되는 값은 한번만 count
                       numMatches = numMatches + 1
                    End If
                End If
                

            End If
        End If
    
    Next cell
    
    Test = uniqueNumbers.Count
    
    ' 유사성 점수를 계산합니다.
    If wsXBRL.name <> "XBRL" Then
        If numMatches > 0 Then
            matchScore = numMatches / uniqueNumbers.Count * numMatches / uniqueNumbers_dart.Count
        Else
            ' 분모가 0인 경우에 대한 처리 (예: 나눗셈을 피하기 위해 0으로 설정)
            matchScore = 0
        End If
                
    Else
        matchScore = numMatches / uniqueNumbers.Count
    End If
    
    ' 유사성 점수 반환
    CompareTables = matchScore
End Function


Sub CopyNoBorderCells() 'control As IRibbonControl)
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim srcCell As Range, rng As Range
    Dim tgtRow As Long
    Dim lastRow As Long
    Dim hasBorder As Boolean
    Dim lastSheet As Worksheet
    
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '코드 실행 중 셀 계산 방지
    
    ' 원본 시트 설정 (DART_2nd 또는 XBRL을 선택)
    On Error Resume Next
    Set srcWs = ActiveWorkbook.Sheets("DART_2nd")
    On Error GoTo 0
    
    If srcWs Is Nothing Then
        Set srcWs = ActiveWorkbook.Sheets("XBRL")
    End If
    
    ' 대상 시트 생성 및 설정
    Set lastSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Set tgtWs = ActiveWorkbook.Sheets.Add(After:=lastSheet)
    tgtWs.name = "Sentences_Recon"
    tgtRow = 2
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    
    ' 원본 시트의 A열을 스캔
    For Each srcCell In srcWs.Range("A1:A" & lastRow)
        Dim borderCount As Integer
        borderCount = 0
        
        ' 테두리가 있는지 검사
        If srcCell.Borders(xlEdgeTop).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        If srcCell.Borders(xlEdgeBottom).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        If srcCell.Borders(xlEdgeLeft).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        If srcCell.Borders(xlEdgeRight).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        
        ' 테두리가 2면 미만인 경우 데이터와 링크를 복사
        If borderCount < 2 And srcCell.Value <> "" Then
            tgtWs.Cells(tgtRow, "E").Formula = "=" & srcWs.name & "!" & srcCell.Address & " & " & srcWs.name & "!" & srcCell.Offset(0, 1).Address
            tgtRow = tgtRow + 1
        End If
    Next srcCell
    
    tgtWs.Cells(1, "E").Value = "DART_2nd"
    tgtWs.Cells(1, "E").Font.Bold = True ' 글꼴을 굵게 설정
    tgtWs.Cells(1, "E").HorizontalAlignment = xlCenter ' 가운데 정렬
    
        ' 원본 시트 설정 (DART_2nd 또는 XBRL을 선택)
    Set srcWs = ActiveWorkbook.Sheets("DART")
    
    tgtRow = 2
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    
    ' 원본 시트의 A열을 스캔
    For Each srcCell In srcWs.Range("A1:A" & lastRow)
        borderCount = 0
        
        ' 테두리가 있는지 검사
        If srcCell.Borders(xlEdgeTop).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        If srcCell.Borders(xlEdgeBottom).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        If srcCell.Borders(xlEdgeLeft).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        If srcCell.Borders(xlEdgeRight).LineStyle <> xlNone Then
            borderCount = borderCount + 1
        End If
        
        ' 테두리가 2면 미만인 경우 데이터와 링크를 복사
        If borderCount < 2 And srcCell.Value <> "" Then
            tgtWs.Cells(tgtRow, "A").Formula = "=" & srcWs.name & "!" & srcCell.Address & " & " & srcWs.name & "!" & srcCell.Offset(0, 1).Address
            tgtRow = tgtRow + 1
        End If
    Next srcCell
    
    tgtWs.Cells(1, "A").Value = "DART"
    tgtWs.Cells(1, "A").Font.Bold = True ' 글꼴을 굵게 설정
    tgtWs.Cells(1, "A").HorizontalAlignment = xlCenter ' 가운데 정렬
    
    FindMostSimilarSentences
    
    Comparecells
            ' 셀 서식 지정
    With Cells
        .VerticalAlignment = xlCenter   ' 텍스트 맞춤 세로(가운데)
        .Interior.Pattern = xlNone  ' 배경 제거
        .WrapText = False   ' 줄바꿈 없음
        .RowHeight = 16.5   ' 줄 높이(StandardHeight값)
        .Font.Size = 9     ' 폰트 사이즈 변경
        .IndentLevel = 0    ' 들여쓰기 없음
        .Font.name = "나눔바른고딕 Light"    ' 폰트 지정
        .ColumnWidth = 50
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub FindMostSimilarSentences()
    Dim wsDART As Worksheet
    Dim wsSentence As Worksheet
    Dim lastRowDART As Integer, lastRowSentence As Integer, lastRow As Integer
    Dim i As Integer, j As Integer
    Dim minDistance As Double, currentDistance As Integer
    Dim mostSimilarSentence As Range
    Dim min_row As Integer
        
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '코드 실행 중 셀 계산 방지
    
    ' DART 시트와 문장대사 시트를 설정합니다.
    Set wsSentence = ActiveWorkbook.Sheets("Sentences_Recon")

    ' 각 시트의 마지막 행을 찾습니다.
    lastRowDART = wsSentence.Cells(wsSentence.Rows.Count, "E").End(xlUp).Row
    lastRowSentence = wsSentence.Cells(wsSentence.Rows.Count, "A").End(xlUp).Row

    ' 문장대사 시트의 첫 번째 컬럼에 가장 유사한 문장을 수식으로 가져옵니다.
    For i = 2 To lastRowSentence ' 첫 번째 행은 헤더이므로 2부터 시작합니다.
        Dim sentence As String
        sentence = wsSentence.Cells(i, "A").Value

        ' 문장을 띄어쓰기 기준으로 조각내기
        Dim sentenceFragments() As String
        sentenceFragments = Split(wsSentence.Cells(i, 1).Value, " ")
                
        minDistance = 10000 ' 초기 최소 거리 설정
        
        min_row = WorksheetFunction.Max(i - 10, 2)
        
        For j = min_row To i + 10
            
             ' 비교 대상 문장을 띄어쓰기 기준으로 조각내기
            Dim targetFragments() As String
            targetFragments = Split(wsSentence.Cells(j, "E").Value, " ")
                       
            If (MatchedFragmentRatio(sentenceFragments, targetFragments) >= 0.5) And (wsSentence.Cells(j, 5).Value <> "") Then

                currentDistance = LevenshteinDistance(wsSentence.Cells(j, 5).Value, wsSentence.Cells(i, 1).Value)
                If currentDistance < minDistance Then
                    minDistance = currentDistance
                    Set mostSimilarSentence = wsSentence.Cells(j, "E")
                End If
            End If
        Next j
        
        If Not mostSimilarSentence Is Nothing Then
            ' 가장 유사한 문장을 두 번째 컬럼에 수식으로 설정
            mostSimilarSentence.Cut Destination:=wsSentence.Cells(i, 2)
            
            ' mostSimilarSentence 변수 리셋
            Set mostSimilarSentence = Nothing
            
        End If

    Next i
    
    wsSentence.Cells(1, "E").Cut Destination:=wsSentence.Cells(1, 2)
    
    lastRow = wsSentence.Cells(wsSentence.Rows.Count, "B").End(xlUp).Row
    
    ' E 열의 데이터를 B 열의 가장 끝에 이동 및 복사
    For i = 1 To lastRowDART
        If Not IsEmpty(wsSentence.Cells(i, "E").Value) Then
            wsSentence.Cells(i, "E").Cut Destination:=wsSentence.Cells(lastRow + 1, "B")
            lastRow = lastRow + 1
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Integer  '문장 유사성확인
    Dim i As Integer, j As Integer
    Dim len1 As Integer, len2 As Integer
    Dim dist() As Integer
    Dim cost As Integer

    len1 = Len(s1)
    len2 = Len(s2)
    ReDim dist(len1, len2)

    For i = 0 To len1
        dist(i, 0) = i
    Next i

    For j = 0 To len2
        dist(0, j) = j
    Next j

    For i = 1 To len1
        For j = 1 To len2
            cost = IIf(Mid(s1, i, 1) = Mid(s2, j, 1), 0, 1)
            dist(i, j) = WorksheetFunction.Min(dist(i - 1, j) + 1, _
                                               dist(i, j - 1) + 1, _
                                               dist(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = dist(len1, len2)
End Function


Function MatchedFragmentRatio(sourceFragments() As String, targetFragments() As String) As Double
    ' 두 개의 조각 배열에서 매치되는 조각의 비율 계산
    Dim matchedCount As Integer
    Dim totalSourceFragments As Integer
    totalSourceFragments = UBound(sourceFragments) - LBound(sourceFragments) + 1
    
    Dim sourceFragment As Variant
    Dim fragment As String
    
    For Each sourceFragment In sourceFragments
        fragment = sourceFragment ' Variant를 String으로 변환
        If IsFragmentMatch(fragment, targetFragments) Then
            matchedCount = matchedCount + 1
        End If
    Next sourceFragment
    
    If totalSourceFragments > 0 Then
        MatchedFragmentRatio = matchedCount / totalSourceFragments
    Else
        MatchedFragmentRatio = 0
    End If
End Function
Function IsFragmentMatch(fragment As String, targetFragments() As String) As Boolean
    ' 조각이 대상 배열에 포함되는지 확인
    Dim targetFragment As Variant
    For Each targetFragment In targetFragments
        If fragment = targetFragment Then
            IsFragmentMatch = True ' 일치하는 경우 True 반환
            Exit Function ' 함수를 종료합니다.
        End If
    Next targetFragment
    
    ' 모든 대상 조각을 검사한 후에도 일치하는 경우가 없으면 False 반환
    IsFragmentMatch = False
End Function


Sub Comparecells()
    Dim aStr As String
    Dim bStr As String
    Dim lastRow As Long
    Dim current_row As Integer
    Dim wsSentence As Worksheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '코드 실행 중 셀 계산 방지
    
    Set wsSentence = ActiveWorkbook.Sheets("Sentences_Recon")
    
    lastRow = WorksheetFunction.Max(wsSentence.Cells(Cells.Rows.Count, "A").End(xlUp).Row, wsSentence.Cells(Cells.Rows.Count, "B").End(xlUp).Row)
    

    For current_row = 2 To lastRow
        
        If Range("A" & current_row).Value = "" And Len(Range("B" & current_row).Value) > 0 Then
            Range("C" & current_row).Value = "추가"
            Range("C" & current_row).Characters.Font.ColorIndex = 14
            Range("C" & current_row).Characters.Font.Underline = True
            
        ElseIf Len(Range("A" & current_row).Value) > 0 And Range("B" & current_row).Value = "" Then
                Range("C" & current_row).Value = "삭제"
                Range("C" & current_row).Characters.Font.ColorIndex = 3
                Range("C" & current_row).Characters.Font.Strikethrough = True
                
        Else
            Dim strOne As String
            Dim strTwo As String
            strOne = Trim(Range("A" & current_row).Value)
            strTwo = Trim(Range("B" & current_row).Value)
            
            Call CompareAndDisplay(strOne, strTwo, Range("C" & current_row))
            Range("C" & current_row).Font.Size = 9
        
        End If

'        If strOne = strTwo Then
'            Range("C" & current_row).Value = "일치"
'
'        End If

        If Range("A" & current_row).Value = Range("B" & current_row).Value Then
            Range("C" & current_row).Value = "일치"

        End If

        
        Range("A" & current_row).WrapText = True
        Range("B" & current_row).WrapText = True
        
    Next
    
    wsSentence.Cells(1, "C").Value = "비교"
    wsSentence.Cells(1, "C").Font.Bold = True ' 글꼴을 굵게 설정
    wsSentence.Cells(1, "C").HorizontalAlignment = xlCenter ' 가운데 정렬
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
Sub CompareAndDisplay(strOne As String, strTwo As String, outCell As Range, Optional Delimiter As String = " ")
    Dim strResult As String
    Dim olStart As Variant, olLength As Variant
    Dim nwStart As Variant, nwLength As Variant
    Dim i As Long
    
    
    On Error Resume Next
    strResult = ComparedText(strOne, strTwo, olStart, olLength, nwStart, nwLength)
    
    With outCell.Cells(1, 1)
        .Clear
        .Value = strResult
        For i = LBound(olStart) To UBound(olStart)
            If olStart(i) <> 0 Then
                
                With .Characters(olStart(i), olLength(i)).Font
                    .ColorIndex = 3
                    .Strikethrough = True
                End With
            End If
        Next i
        For i = LBound(nwStart) To UBound(nwStart)
            If nwStart(i) <> 0 Then
                With .Characters(nwStart(i), nwLength(i)).Font
                    .ColorIndex = 14
                    .Underline = True
                End With
            End If
        Next i
    End With
    
    
    
End Sub
Function ComparedText(aString As String, bString As String, _
                                        Optional ByRef oStart As Variant, Optional oLength As Variant, _
                                        Optional ByRef nStart As Variant, Optional ByRef nLength As Variant, _
                                        Optional Delimiter As String = " ") As String
    Dim aWords As Variant, aWord As String
    Dim bWords As Variant
    Dim outWords() As String
    Dim aPoint As Long, bPoint As Long, outPoint As Long
    Dim matchPoint As Variant, outLength As Long
    Dim High As Long
    Dim oldStart() As Long, oldLength() As Long, oldPoint As Long
    Dim newStart() As Long, newLength() As Long, newPoint As Long
    
    Rem remove double delimiters
        'to be done
    
    aWords = Split(aString, Delimiter)
    bWords = Split(bString, Delimiter)
    High = UBound(aWords)
    ReDim outWords(0 To High + UBound(bWords))
    ReDim oldStart(0 To High + UBound(bWords)): ReDim oldLength(0 To High + UBound(bWords))
    ReDim newStart(0 To High + UBound(bWords)): ReDim newLength(0 To High + UBound(bWords))
    oldPoint = -1: newPoint = -1
    outLength = Len(Delimiter)
    aPoint = 0: bPoint = 0: outPoint = LBound(outWords) - 1
    
    Do
        aWord = aWords(aPoint)
        If LCase(aWord) = LCase(bWords(bPoint)) Then
            Rem is word in common
            outPoint = outPoint + 1
            outWords(outPoint) = aWord
            outLength = outLength + Len(aWord) + Len(Delimiter)
            bWords(bPoint) = vbNullString
            aPoint = aPoint + 1
            bPoint = bPoint + 1
        Else
            Rem word divergence
            matchPoint = Application.Match(aWord, bWords, 0)
            If IsError(matchPoint) Then
                Rem old word is not in new string
                outPoint = outPoint + 1
                outWords(outPoint) = aWord
                
                oldPoint = oldPoint + 1
                oldStart(oldPoint) = outLength: oldLength(oldPoint) = Len(aWord)
                
                outLength = outLength + Len(aWord) + Len(Delimiter)
                aPoint = aPoint + 1
            Else
                Rem old word is in new string, i.e. is a common word
                
                Do Until LCase(bWords(bPoint)) = LCase(aWord)
                    outPoint = outPoint + 1
                    outWords(outPoint) = bWords(bPoint)
                    
                    newPoint = newPoint + 1
                    newStart(newPoint) = outLength: newLength(newPoint) = Len(bWords(bPoint))
        
                    outLength = outLength + Len(outWords(outPoint)) + Len(Delimiter)
                    bWords(bPoint) = vbNullString
                    bPoint = bPoint + 1
                Loop
            End If
        End If
    Loop Until (High < aPoint) Or (UBound(bWords) < bPoint)
    
    Rem last new/different string
    Do Until UBound(bWords) < bPoint
        outPoint = outPoint + 1
        outWords(outPoint) = bWords(bPoint)
        
        newPoint = newPoint + 1
        newStart(newPoint) = outLength: newLength(newPoint) = Len(bWords(bPoint))
        
        outLength = outLength + Len(outWords(outPoint)) + Len(Delimiter)
        bPoint = bPoint + 1
    Loop
    
    Rem final common string
    Do Until High < aPoint
        outPoint = outPoint + 1
        outWords(outPoint) = aWords(aPoint)
        
        oldPoint = oldPoint + 1
        oldStart(oldPoint) = outLength: oldLength(oldPoint) = Len(aWords(aPoint))
        
        outLength = outLength + Len(outWords(outPoint)) + Len(Delimiter)
        aPoint = aPoint + 1
    Loop
    
    ReDim Preserve outWords(0 To outPoint)
    oStart = oldStart: oLength = oldLength
    nStart = newStart: nLength = newLength
    ComparedText = Join(outWords, Delimiter)
End Function

Sub ImportWordDocument()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim wordPara As Object
    Dim i As Integer
    Dim filePath As String

    ' Word 파일 경로 설정
    filePath = "C:\Users\yeyi\Downloads\데이터분석\Tie-Out_XBRL완전성확인\KEPCO_Consolidated_FY22_4Q_Final.docx"

    ' Word 애플리케이션 객체 생성
    Set wordApp = CreateObject("Word.Application")

    ' Word 문서 열기
    Set wordDoc = wordApp.Documents.Open(filePath)

    ' Excel 시트 초기화
    i = 1

    ' Word 문서의 각 단락을 읽어서 Excel 시트에 쓰기
    For Each wordPara In wordDoc.Paragraphs
        ThisWorkbook.Sheets("Sheet1").Cells(i, 1).Value = wordPara.Range.Text
        i = i + 1
    Next wordPara

    ' Word 문서 닫기 및 Word 애플리케이션 종료
    wordDoc.Close
    wordApp.Quit

    ' 객체 해제
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub


