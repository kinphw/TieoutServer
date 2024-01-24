Attribute VB_Name = "modTieoutMod1"
Sub ExcelFile_Get_dart(ByRef filePath As String)
    Dim i As Integer
    Dim NewFileName As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim cellAddress As String
    Dim sht As String
    Dim keyword1 As String, keyword2 As String
    

    ' ���� ��θ� �����մϴ�.
    'FilePath = "C:\Users\yeyi\Downloads\�����ͺм�\XBRL������Ȯ��\SKT �������纸��_2022_0310_����_����-DOC.htm"
    
    ' ���� ���� ����
    Workbooks.Open (filePath)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '�ڵ� ���� �� �� ��� ����
    
    ' ���� ���� ���İ� ���ο� �̸����� ����
    'NewFileName = ActiveSheet.name + Format(Date, "-mm-dd") + Format(Time, "_hhmmss")
    'NewFileName = ThisWorkbook.Path + "\work\TieOut" + Format(Date, "-mm-dd") + Format(Time, "_hhmmss")
    NewFileName = "C:\xampp\htdocs\work\TieOut" + Format(Date, "-mm-dd") + Format(Time, "_hhmmss")
    ActiveWorkbook.SaveAs Filename:=NewFileName, FileFormat:=xlWorkbookDefault

    ' �۾��Ϸ��� ��ũ��Ʈ�� �����մϴ�. ���ϴ� ��ũ��Ʈ�� �̸����� �����ϼ���.
    Set ws = ActiveWorkbook.Sheets(1)
    ws.name = "DART"
    
    ' ã�� Ű���带 �����մϴ�.
    keyword1 = "�繫����ǥ"
    keyword2 = "�����繫����ǥ"
    
    ' �������� ��ũ��Ʈ�� �� ���� Ȯ���ϰ� Ű���带 ã���ϴ�.
    For Each cell In ws.UsedRange
        If Replace(cell.Value, " ", "") = keyword1 Or Replace(cell.Value, " ", "") = keyword2 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword1 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword2 Then '�ݱ��繫����ǥó�� ��2���ڰ� �ٴ� ��� �߰����
            ' Ű���带 ã�� ��� �ش� ���� �ּҸ� ����մϴ�.
            cellAddress = cell.Address
            Exit For
        End If
    Next cell
    
    ' Ű���带 ã�� ��� �ش� ���� ������ ��� �����մϴ�.
    If cellAddress <> "" Then
        ' Ű���带 ã�� ���� �� �� ���� ����� �����մϴ�.
        ws.Rows("1:" & ws.Range(cellAddress).Row - 1).Delete
        ws.Cells(1, 1).EntireRow.Insert Shift:=xlDown
        
    End If
    
    
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete ' �̹��� Shape ����
    Next
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Range(ws.Cells(i, "A"), ws.Cells(i, "Z")).Hyperlinks.Count > 0 And ws.Cells(i, "A").Value <> "�ּ�" Then
           ws.Range("A" & (i) & ":A" & lastRow + 10).EntireRow.Delete
        End If
    Next i
    
    
    Dim rowCount As Integer
    Dim emptyRowCount As Integer

    rowCount = ActiveSheet.UsedRange.Rows.Count
    emptyRowCount = 0
    
    ' �ڿ������� �˻��Ͽ� �� �� ī��Ʈ
    For i = rowCount To 1 Step -1
        If WorksheetFunction.CountA(ActiveSheet.Rows(i)) = 0 Then
            emptyRowCount = emptyRowCount + 1
            
            ' �� �� ����
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
    
    ' ���� ���� ������ �����ϰ� �ݽ��ϴ�.
    ActiveWorkbook.Save
    'ActiveWorkbook.Close
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExcelFile_Get_dart2(ByRef filePath As String)

    Dim wsSource As Worksheet ' �� ��° ���� ������ ��Ʈ
    Dim wsTarget As Worksheet ' ù ��° ���� ������ ��� ��Ʈ ("XBRL" ��Ʈ)
    Dim wbSource As Workbook  ' �� ��° ���� ����
    Dim sht As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '�ڵ� ���� �� �� ��� ����
    
    ' �� ��° ���� ��θ� �����մϴ�.
    'Dim FilePath As String
    'FilePath = "C:\Users\yeyi\Downloads\�����ͺм�\XBRL������Ȯ��\�������_69_00121941_����ֽ�ȸ��_0711_�˻�Ϸ�.xls"
    
    ' ù ��° ���� ������ "XBRL" ��Ʈ�� �����մϴ�.
    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.name = "XBRL"
    Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    ' �� ��° ���� ���� ����
    Set wbSource = Workbooks.Open(filePath)
    
    ' �� ��° ���� ������ ��Ʈ�� ��ȸ�ϸ� "�⺻����" ��Ʈ�� �����ϰ� �����͸� �����մϴ�.
    For Each wsSource In wbSource.Sheets
        ' "�⺻����" ��Ʈ�� �����ϰ� �����͸� �����մϴ�.
        wsSource.UsedRange.Copy wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Offset(2, 0)
    Next wsSource
    
    ' �� ��° ���� ������ �ݽ��ϴ�.
    wbSource.Close SaveChanges:=False
    
    
    ' ã�� Ű���带 �����մϴ�.
    keyword1 = "�繫����ǥ"
    keyword2 = "�����繫����ǥ"
    
    ' �������� ��ũ��Ʈ�� �� ���� Ȯ���ϰ� Ű���带 ã���ϴ�.
    For Each cell In wsTarget.UsedRange
        If Replace(cell.Value, " ", "") = keyword1 Or Replace(cell.Value, " ", "") = keyword2 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword1 Or Mid(Replace(cell.Value, " ", ""), 3) = keyword2 Then '�ݱ��繫����ǥó�� ��2���ڰ� �ٴ� ��� �߰����
            ' Ű���带 ã�� ��� �ش� ���� �ּҸ� ����մϴ�.
            cellAddress = cell.Address
            Exit For
        End If
    Next cell
    
    ' Ű���带 ã�� ��� �ش� ���� ������ ��� �����մϴ�.
    If cellAddress <> "" Then
        ' Ű���带 ã�� ���� �� �� ���� ����� �����մϴ�.
        wsTarget.Rows("1:" & wsTarget.Range(cellAddress).Row - 1).Delete
        wsTarget.Cells(1, 1).EntireRow.Insert Shift:=xlDown
        
    End If
    
    
    For i = wsTarget.Shapes.Count To 1 Step -1
        wsTarget.Shapes(i).Delete ' �̹��� Shape ����
    Next
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If wsTarget.Range(wsTarget.Cells(i, "A"), wsTarget.Cells(i, "Z")).Hyperlinks.Count > 0 And wsTarget.Cells(i, "A").Value <> "�ּ�" Then
           wsTarget.Range("A" & (i) & ":A" & lastRow + 10).EntireRow.Delete
        End If
    Next i
    
    Dim rowCount As Integer
    Dim emptyRowCount As Integer
    
    'Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    rowCount = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    emptyRowCount = 0
    
    ' �ڿ������� �˻��Ͽ� �� �� ī��Ʈ
    For i = rowCount To 1 Step -1
        If WorksheetFunction.CountA(wsTarget.Rows(i)) = 0 Then
            emptyRowCount = emptyRowCount + 1
    
            ' �� �� ����
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
    
    ' ������ �ڵ� �ۼ�
End Sub


Sub ExcelFile_Get_xbrl(ByRef filePath As String)

    Dim wsSource As Worksheet ' �� ��° ���� ������ ��Ʈ
    Dim wsTarget As Worksheet ' ù ��° ���� ������ ��� ��Ʈ ("XBRL" ��Ʈ)
    Dim wbSource As Workbook  ' �� ��° ���� ����
    Dim sht As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '�ڵ� ���� �� �� ��� ����
    
    ' �� ��° ���� ��θ� �����մϴ�.
    'Dim FilePath As String
    'FilePath = "C:\Users\yeyi\Downloads\�����ͺм�\Tie-Out_XBRL������Ȯ��\�������_69_00121941_����ֽ�ȸ��_0711_�˻�Ϸ�.xls"
    
    ' ù ��° ���� ������ "XBRL" ��Ʈ�� �����մϴ�.
    ActiveWorkbook.Sheets.Add
    ActiveSheet.name = "XBRL"
    Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    ' �� ��° ���� ���� ����
    Set wbSource = Workbooks.Open(filePath)
    
    note_count = 1
    
    ' �� ��° ���� ������ ��Ʈ�� ��ȸ�ϸ� "�⺻����" ��Ʈ�� �����ϰ� �����͸� �����մϴ�.
    For Each wsSource In wbSource.Sheets
        ' "�⺻����" ��Ʈ�� �����ϰ� �����͸� �����մϴ�.
        If wsSource.name <> "�⺻����" Then
            ' �����͸� �ٿ����� ��ġ�� �����մϴ�.
            
            If wsSource.name Like "*�ּ�*" Then
            
                Dim lastRow As Long
                lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
                
                ' ������ ���� ���� ���յǾ����� Ȯ���մϴ�.
                Dim mergedRows As Long
                If wsTarget.Cells(lastRow, "A").MergeCells Then
                    ' ���յ� ���, ���յ� ���� ������ ã���ϴ�.
                    mergedRows = wsTarget.Cells(lastRow, "A").MergeArea.Rows.Count
                    ' ���յ� ���� ������ �� ��ȣ�� ����մϴ�.
                    insertRow = lastRow + mergedRows + 1
                Else
                    ' ���յ��� ���� ���, ������ �� ��ȣ���� +2�� �մϴ�.
                    insertRow = lastRow + 2
                End If
                
                ' ��Ʈ �̸��� �ٿ����� ��ġ�� �Է��մϴ�.
                wsTarget.Cells(insertRow, 1).Value = note_count & ". " & Split(wsSource.name, "�ּ� - ")(1)
                
                note_count = note_count + 1
            End If
        
            wsSource.UsedRange.Copy wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Offset(2, 0)
        End If
    Next wsSource
    
    ' �� ��° ���� ������ �ݽ��ϴ�.
    wbSource.Close SaveChanges:=False
    
    
    Dim rowCount As Integer
    Dim emptyRowCount As Integer
    
    'Set wsTarget = ActiveWorkbook.Sheets("XBRL")
    
    rowCount = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    emptyRowCount = 0
    
    ' �ڿ������� �˻��Ͽ� �� �� ī��Ʈ
    For i = rowCount To 1 Step -1
        If WorksheetFunction.CountA(wsTarget.Rows(i)) = 0 Then
            emptyRowCount = emptyRowCount + 1
    
            ' �� �� ����
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
    
    ' ������ �ڵ� �ۼ�
End Sub


Sub SetStyle()
Dim rng As Range
Dim column As Range
Dim originalWidth As Double
Dim newWidth As Double
Dim lastColumn As Long

    ' �� ���� ����
    With Cells
        .VerticalAlignment = xlCenter   ' �ؽ�Ʈ ���� ����(���)
        .Interior.Pattern = xlNone  ' ��� ����
        .WrapText = False   ' �ٹٲ� ����
        .RowHeight = 16.5   ' �� ����(StandardHeight��)
        .Font.Size = 9     ' ��Ʈ ������ ����
        .IndentLevel = 0    ' �鿩���� ����
        .Font.name = "�����ٸ���� Light"    ' ��Ʈ ����

        ' ���� ���ڷ� ����
        .Replace What:=ThisWorkbook.Sheets(1).Range("A1").Value, Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        ' ���ڷ� ���� -> ����� �Բ� �ִ� ��찡 �־� ���� ���� �� �����ϵ��� ��ġ ����
        '.Replace What:="-", Replacement:="0", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    End With
    
    ActiveSheet.UsedRange.ColumnWidth = 20
    Columns("A:A").ColumnWidth = 30
        
           
    Set rng = ActiveSheet.UsedRange ' �Ǵ� ������ ������ ����
    
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


    '-'ǥ�� ����'0'���� ���� �� ����κ� ������ �Է�
    For Each cell In rng
       If Len(Trim(cell.Value)) = 1 And Trim(cell.Value) = "-" Then
            cell.Value = "0"
        End If
       
    Next cell



    ' ���� Clear
    Range("A1").Select
    

End Sub

'======================================
' Table�̸� ���� -> Ref.Check���� ���� ǥ�� �� ���(�ּ�)
'======================================

Sub SetTableName(ByRef sht As String)

    Dim tableRange As Range, curRange As Range, preRange As Range, nextRange As Range, nextRange2 As Range
    Dim ws As Worksheet
    Dim noteStart As Boolean
    
    Set ws = ActiveWorkbook.ActiveSheet
        
    rowASheetCount = ActiveSheet.UsedRange.Cells.Rows.Count + 10
    colASheetCount = ActiveSheet.UsedRange.Cells.Columns.Count

    ' Name ��ȣ
    NoteNumber = 1
    TableNumber = 1
    noteStart = False
    
    Dim keywords As Variant
    keywords = Array("�繫����ǥ", "���Ͱ�꼭", "�ں�����ǥ", "�����帧ǥ")


    ' ���̺� Ȯ��
    For Each curRange In ActiveSheet.UsedRange.Cells
        
        ' �ּ� ��ȣ Ȯ��
        If curRange.column = 1 And CStr(curRange.Value) Like NoteNumber & ".*" Then
            ' �ּ� Name ����
            Dim NoteTitle As String
            NoteTitle = Trim(curRange.Value)
            ws.Names.Add name:="NOTE" & NoteNumber, RefersTo:=curRange
            
            NoteNumber = NoteNumber + 1
            
            noteStart = True
        End If
    
        If noteStart = False Then
            If curRange.Borders(xlEdgeLeft).LineStyle = xlNone And curRange.Borders(xlEdgeRight).LineStyle = xlNone And _
               curRange.Borders(xlEdgeTop).LineStyle = xlNone And curRange.Borders(xlEdgeBottom).LineStyle = xlNone Then
               
                ' ���� �ؽ�Ʈ���� ��� ������ �����մϴ�.
                Dim textWithoutSpaces As String
                textWithoutSpaces = Replace(curRange.Value, " ", "")
        
                ' Ư�� Ű���尡 ���ԵǾ� �ִ��� �˻��մϴ�.
                For Each keyword In keywords
                    If InStr(textWithoutSpaces, keyword) > 0 Then
                        NoteTitle = keyword
                        Exit For ' Ű���带 ã���� ������ �����մϴ�.
                    End If
                Next keyword
                
            End If
        End If
                
        'curRange.Select
        ' ���̺� ���� ��ġ
        If rowTableStart = 0 And curRange.Row > 1 _
            And curRange.Borders(xlEdgeLeft).LineStyle = xlContinuous And curRange.Borders(xlEdgeTop).LineStyle = xlContinuous Then
            ' ���� �������� ���� ��ġ Ȯ��
            Set preRange = Cells(curRange.Row - 1, curRange.column)
            If preRange.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone _
                And preRange.Borders(xlEdgeRight).LineStyle = xlLineStyleNone Then
                rowTableStart = curRange.Row
                colTableStart = curRange.column
            End If
        End If
        
        ' ���̺� ���� ��ġ
        If rowTableStart > 0 And curRange.Borders(xlEdgeRight).LineStyle = xlContinuous _
            And curRange.Borders(xlEdgeBottom).LineStyle = xlContinuous Then
            
            curRange.Value = RTrim(curRange.Value)
            
            ' ���� ���� ��ȸ�Ͽ� ���� ��ġ Ȯ��
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
        
        ' ���̺��� ��� ó��
        If rowTableStart > 0 And colTableEnd > 0 And rowTableStart = rowTableEnd And colTableStart = colTableEnd Then
        
            rowTableStart = 0
            colTableEnd = 0

        ElseIf rowTableStart > 0 And colTableEnd > 0 Then
    
            ' ���̺� ����
            Set tableRange = ws.Range(Cells(rowTableStart, colTableStart), Cells(rowTableEnd, colTableEnd))
            
            ' ���̺� Name ����
            Dim nameNote As String
            nameNote = sht & "_TABLE" & TableNumber
            
            'tableRange.name = nameNote '���չ����������� name ����
            ws.Names.Add name:=nameNote, RefersTo:=tableRange '�ش� ��Ʈ�������� name ����
            ws.Names(nameNote).Comment = NoteTitle
            
            TableNumber = TableNumber + 1

            ' ���̺� ��/�ϴ� ���� �߰�
            If Not IsEmpty(Cells(rowTableEnd + 1, 1).Value) Then
                Cells(rowTableEnd + 1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            End If
            
            ' CheckRange �ʱ�ȭ
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

    ' ��� �̸��� Ȯ���ϰ� ���� �̸��� �����մϴ�.
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        On Error Resume Next
        If nm.RefersToRange Is Nothing Then
            ' ���� �̸��� ��� �����մϴ�.
            nm.Delete
        End If
        On Error GoTo 0
    Next i
    
    
    
    ' "XBRL" ��Ʈ�� �ִ��� Ȯ��
    On Error Resume Next ' ��Ʈ�� ã�� �� ���� ��� ���� ó��
    Set wsXBRL = ActiveWorkbook.Sheets("XBRL")
    On Error GoTo 0 ' ���� ó���� ������� ����
    
    ' "XBRL" ��Ʈ�� ���� ��� "DART" ��Ʈ�� ����
    If wsXBRL Is Nothing Then
        Set wsXBRL = ActiveWorkbook.Sheets("DART_2nd")
    End If
    
    'Set wsXBRL = ActiveWorkbook.Sheets("XBRL")
    Set wsDART = ActiveWorkbook.Sheets("DART")
    
    ' 'XBRL' ��Ʈ�� ���̺��� ��ȸ�մϴ�.
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
                                    ' ���ǿ� ���� �ٸ� �۾� ����
                                    If cell.Interior.Color = RGB(255, 255, 153) Or cell.Interior.Color = RGB(255, 217, 102) Then
                                        coloredCellCount = coloredCellCount + 1
                                    End If
                                End If
                            Next cell
                            
                            ' ��ġ�ϴ� 'DART' ���̺��� �����մϴ�.
                            If matchScore > bestMatchScore And (coloredCellCount / numericCellCount) < 0.8 Then '��������� ������ 80% �̻��̸� �̹� �ٸ� ���̺�� ��Ī�� ���� ���̺�� �Ǵ�
                                bestMatchScore = matchScore
                                ' �ش� 'DART' ���̺��� �����ϰų� �۾��� �����մϴ�.
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
                            c.Interior.Color = RGB(255, 0, 0) '������
                            c.Font.Color = RGB(255, 255, 255)
                            
                            For Each d In rngB
                                If IsNumeric(d.Value) And Not IsEmpty(d.Value) And Not d.HasFormula Then
                                   
                                    If c.Value = d.Value Then ' If the values in the cells are the same

                                        d.Formula = "=" & wsXBRL.name & "!" & c.Address(0, 0)
                                        d.Interior.Color = RGB(255, 255, 153) '�����
                                        'd.Font.Color = RGB(0, 0, 0)
                                        c.Interior.ColorIndex = xlNone
                                        c.Font.Color = RGB(0, 0, 0)
            
                                        Exit For ' Exit the loop since a link has been added to cell d
                                        
                                    ElseIf c.Value * -1 = d.Value Then

                                        d.Formula = "=" & wsXBRL.name & "!" & c.Address(0, 0)
                                        d.Interior.Color = RGB(255, 217, 102) '£�� �����
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
    
    '����
    ' ��� ��Ʈ ���� �� ����
    Set lastSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Set tgtWs = ActiveWorkbook.Sheets.Add(After:=lastSheet)
    tgtWs.name = "Tables_Recon"
    
    tgtWs.Cells(1, "A").Value = "Sheet"
    tgtWs.Cells(1, "A").Font.Bold = True ' �۲��� ���� ����
    tgtWs.Cells(1, "A").HorizontalAlignment = xlCenter ' ��� ����
        
    tgtWs.Cells(1, "B").Value = "Note"
    tgtWs.Cells(1, "B").Font.Bold = True ' �۲��� ���� ����
    tgtWs.Cells(1, "B").HorizontalAlignment = xlCenter ' ��� ����
        
    tgtWs.Cells(1, "C").Value = "Missing value"
    tgtWs.Cells(1, "C").Font.Bold = True ' �۲��� ���� ����
    tgtWs.Cells(1, "C").HorizontalAlignment = xlCenter ' ��� ����
    tgtRow = 2
        
    For Each wsXBRL In ActiveWorkbook.Sheets
        If wsXBRL.name <> "DART" Then
            For Each nameXBRL In wsXBRL.Names
                If nameXBRL.name Like "*TABLE*" Then
                    Set rng = nameXBRL.RefersToRange
                    
                    For Each c In rng
                        If c.Interior.Color = RGB(255, 0, 0) Then '������
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
    ' 'XBRL' ���̺�� 'DART' ��Ʈ ���� ���ڰ� ���缺�� ���ϴ� �Լ�
    Dim matchScore As Double
    Dim rng As Range
    Dim uniqueNumbers_dart As Object
        
    Set rng = name.RefersToRange
    
    matchScore = 0
    
    ' 'XBRL' ���̺��� ������ ���ڰ��� �����մϴ�.

    
    ' 'DART' ��Ʈ���� �ش� ���ڰ��� ���մϴ�.
    Dim numMatches As Double
    numMatches = 0
    
    Set uniqueNumbers_dart = CreateObject("Scripting.Dictionary")
    
    For Each cell In rng    'dart�� ���̺�
        'If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) And cell.Value <> 0 And Not uniqueNumbers_dart.Exists(cell.Value) Then
        '    uniqueNumbers_dart(cell.Value) = 1
        '    If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) And cell.Value <> 0 Then
        '        If uniqueNumbers.Exists(cell.Value) Then
        '            numMatches = numMatches + 1
        '        End If
        '    End If
        'End If
        
        If Not cell.HasFormula Then ' ������ �ƴ� ��츸 ó��
            If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then 'And cell.Value <> 0 Then
                If Not uniqueNumbers_dart.Exists(cell.Value) Then  '�ߺ��Ǵ� ���� ������� ����.
                    uniqueNumbers_dart(cell.Value) = 1
                    
                    If uniqueNumbers.Exists(cell.Value) Then  'uniqueNumbers �ȿ��� �ߺ����� �����Ƿ� �ߺ��Ǵ� ���� �ѹ��� count
                       numMatches = numMatches + 1
                    End If
                End If
                

            End If
        End If
    
    Next cell
    
    Test = uniqueNumbers.Count
    
    ' ���缺 ������ ����մϴ�.
    If wsXBRL.name <> "XBRL" Then
        If numMatches > 0 Then
            matchScore = numMatches / uniqueNumbers.Count * numMatches / uniqueNumbers_dart.Count
        Else
            ' �и� 0�� ��쿡 ���� ó�� (��: �������� ���ϱ� ���� 0���� ����)
            matchScore = 0
        End If
                
    Else
        matchScore = numMatches / uniqueNumbers.Count
    End If
    
    ' ���缺 ���� ��ȯ
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
    Application.Calculation = xlCalculationManual   '�ڵ� ���� �� �� ��� ����
    
    ' ���� ��Ʈ ���� (DART_2nd �Ǵ� XBRL�� ����)
    On Error Resume Next
    Set srcWs = ActiveWorkbook.Sheets("DART_2nd")
    On Error GoTo 0
    
    If srcWs Is Nothing Then
        Set srcWs = ActiveWorkbook.Sheets("XBRL")
    End If
    
    ' ��� ��Ʈ ���� �� ����
    Set lastSheet = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Set tgtWs = ActiveWorkbook.Sheets.Add(After:=lastSheet)
    tgtWs.name = "Sentences_Recon"
    tgtRow = 2
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    
    ' ���� ��Ʈ�� A���� ��ĵ
    For Each srcCell In srcWs.Range("A1:A" & lastRow)
        Dim borderCount As Integer
        borderCount = 0
        
        ' �׵θ��� �ִ��� �˻�
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
        
        ' �׵θ��� 2�� �̸��� ��� �����Ϳ� ��ũ�� ����
        If borderCount < 2 And srcCell.Value <> "" Then
            tgtWs.Cells(tgtRow, "E").Formula = "=" & srcWs.name & "!" & srcCell.Address & " & " & srcWs.name & "!" & srcCell.Offset(0, 1).Address
            tgtRow = tgtRow + 1
        End If
    Next srcCell
    
    tgtWs.Cells(1, "E").Value = "DART_2nd"
    tgtWs.Cells(1, "E").Font.Bold = True ' �۲��� ���� ����
    tgtWs.Cells(1, "E").HorizontalAlignment = xlCenter ' ��� ����
    
        ' ���� ��Ʈ ���� (DART_2nd �Ǵ� XBRL�� ����)
    Set srcWs = ActiveWorkbook.Sheets("DART")
    
    tgtRow = 2
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    
    ' ���� ��Ʈ�� A���� ��ĵ
    For Each srcCell In srcWs.Range("A1:A" & lastRow)
        borderCount = 0
        
        ' �׵θ��� �ִ��� �˻�
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
        
        ' �׵θ��� 2�� �̸��� ��� �����Ϳ� ��ũ�� ����
        If borderCount < 2 And srcCell.Value <> "" Then
            tgtWs.Cells(tgtRow, "A").Formula = "=" & srcWs.name & "!" & srcCell.Address & " & " & srcWs.name & "!" & srcCell.Offset(0, 1).Address
            tgtRow = tgtRow + 1
        End If
    Next srcCell
    
    tgtWs.Cells(1, "A").Value = "DART"
    tgtWs.Cells(1, "A").Font.Bold = True ' �۲��� ���� ����
    tgtWs.Cells(1, "A").HorizontalAlignment = xlCenter ' ��� ����
    
    FindMostSimilarSentences
    
    Comparecells
            ' �� ���� ����
    With Cells
        .VerticalAlignment = xlCenter   ' �ؽ�Ʈ ���� ����(���)
        .Interior.Pattern = xlNone  ' ��� ����
        .WrapText = False   ' �ٹٲ� ����
        .RowHeight = 16.5   ' �� ����(StandardHeight��)
        .Font.Size = 9     ' ��Ʈ ������ ����
        .IndentLevel = 0    ' �鿩���� ����
        .Font.name = "�����ٸ���� Light"    ' ��Ʈ ����
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
    Application.Calculation = xlCalculationManual   '�ڵ� ���� �� �� ��� ����
    
    ' DART ��Ʈ�� ������ ��Ʈ�� �����մϴ�.
    Set wsSentence = ActiveWorkbook.Sheets("Sentences_Recon")

    ' �� ��Ʈ�� ������ ���� ã���ϴ�.
    lastRowDART = wsSentence.Cells(wsSentence.Rows.Count, "E").End(xlUp).Row
    lastRowSentence = wsSentence.Cells(wsSentence.Rows.Count, "A").End(xlUp).Row

    ' ������ ��Ʈ�� ù ��° �÷��� ���� ������ ������ �������� �����ɴϴ�.
    For i = 2 To lastRowSentence ' ù ��° ���� ����̹Ƿ� 2���� �����մϴ�.
        Dim sentence As String
        sentence = wsSentence.Cells(i, "A").Value

        ' ������ ���� �������� ��������
        Dim sentenceFragments() As String
        sentenceFragments = Split(wsSentence.Cells(i, 1).Value, " ")
                
        minDistance = 10000 ' �ʱ� �ּ� �Ÿ� ����
        
        min_row = WorksheetFunction.Max(i - 10, 2)
        
        For j = min_row To i + 10
            
             ' �� ��� ������ ���� �������� ��������
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
            ' ���� ������ ������ �� ��° �÷��� �������� ����
            mostSimilarSentence.Cut Destination:=wsSentence.Cells(i, 2)
            
            ' mostSimilarSentence ���� ����
            Set mostSimilarSentence = Nothing
            
        End If

    Next i
    
    wsSentence.Cells(1, "E").Cut Destination:=wsSentence.Cells(1, 2)
    
    lastRow = wsSentence.Cells(wsSentence.Rows.Count, "B").End(xlUp).Row
    
    ' E ���� �����͸� B ���� ���� ���� �̵� �� ����
    For i = 1 To lastRowDART
        If Not IsEmpty(wsSentence.Cells(i, "E").Value) Then
            wsSentence.Cells(i, "E").Cut Destination:=wsSentence.Cells(lastRow + 1, "B")
            lastRow = lastRow + 1
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Integer  '���� ���缺Ȯ��
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
    ' �� ���� ���� �迭���� ��ġ�Ǵ� ������ ���� ���
    Dim matchedCount As Integer
    Dim totalSourceFragments As Integer
    totalSourceFragments = UBound(sourceFragments) - LBound(sourceFragments) + 1
    
    Dim sourceFragment As Variant
    Dim fragment As String
    
    For Each sourceFragment In sourceFragments
        fragment = sourceFragment ' Variant�� String���� ��ȯ
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
    ' ������ ��� �迭�� ���ԵǴ��� Ȯ��
    Dim targetFragment As Variant
    For Each targetFragment In targetFragments
        If fragment = targetFragment Then
            IsFragmentMatch = True ' ��ġ�ϴ� ��� True ��ȯ
            Exit Function ' �Լ��� �����մϴ�.
        End If
    Next targetFragment
    
    ' ��� ��� ������ �˻��� �Ŀ��� ��ġ�ϴ� ��찡 ������ False ��ȯ
    IsFragmentMatch = False
End Function


Sub Comparecells()
    Dim aStr As String
    Dim bStr As String
    Dim lastRow As Long
    Dim current_row As Integer
    Dim wsSentence As Worksheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   '�ڵ� ���� �� �� ��� ����
    
    Set wsSentence = ActiveWorkbook.Sheets("Sentences_Recon")
    
    lastRow = WorksheetFunction.Max(wsSentence.Cells(Cells.Rows.Count, "A").End(xlUp).Row, wsSentence.Cells(Cells.Rows.Count, "B").End(xlUp).Row)
    

    For current_row = 2 To lastRow
        
        If Range("A" & current_row).Value = "" And Len(Range("B" & current_row).Value) > 0 Then
            Range("C" & current_row).Value = "�߰�"
            Range("C" & current_row).Characters.Font.ColorIndex = 14
            Range("C" & current_row).Characters.Font.Underline = True
            
        ElseIf Len(Range("A" & current_row).Value) > 0 And Range("B" & current_row).Value = "" Then
                Range("C" & current_row).Value = "����"
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
'            Range("C" & current_row).Value = "��ġ"
'
'        End If

        If Range("A" & current_row).Value = Range("B" & current_row).Value Then
            Range("C" & current_row).Value = "��ġ"

        End If

        
        Range("A" & current_row).WrapText = True
        Range("B" & current_row).WrapText = True
        
    Next
    
    wsSentence.Cells(1, "C").Value = "��"
    wsSentence.Cells(1, "C").Font.Bold = True ' �۲��� ���� ����
    wsSentence.Cells(1, "C").HorizontalAlignment = xlCenter ' ��� ����
    
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

    ' Word ���� ��� ����
    filePath = "C:\Users\yeyi\Downloads\�����ͺм�\Tie-Out_XBRL������Ȯ��\KEPCO_Consolidated_FY22_4Q_Final.docx"

    ' Word ���ø����̼� ��ü ����
    Set wordApp = CreateObject("Word.Application")

    ' Word ���� ����
    Set wordDoc = wordApp.Documents.Open(filePath)

    ' Excel ��Ʈ �ʱ�ȭ
    i = 1

    ' Word ������ �� �ܶ��� �о Excel ��Ʈ�� ����
    For Each wordPara In wordDoc.Paragraphs
        ThisWorkbook.Sheets("Sheet1").Cells(i, 1).Value = wordPara.Range.Text
        i = i + 1
    Next wordPara

    ' Word ���� �ݱ� �� Word ���ø����̼� ����
    wordDoc.Close
    wordApp.Quit

    ' ��ü ����
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub


