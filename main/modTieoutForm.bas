Attribute VB_Name = "modTieoutForm"
Sub Begin()

    Dim x As Integer
    Dim maComp As String
    Dim obj_fso As Object
    
    '���� ��ȿ�� üũ
    
    Dim TextBox_dart_Path As String
    Dim TextBox_dart2_Path As String
    
    'TextBox_dart_Path = ThisWorkbook.Path & "\work\a.htm"
    'TextBox_dart2_Path = ThisWorkbook.Path & "\work\b.htm"
    
    TextBox_dart_Path = "C:\xampp\htdocs\work\a.htm"
    TextBox_dart2_Path = "C:\xampp\htdocs\work\b.htm"
    
    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    If obj_fso.fileExists(TextBox_dart_Path) = False Then
        MsgBox "���ϰ�ΰ� �߸��Ǿ����ϴ�. ��θ� �ٽ� Ȯ���Ͽ� �ֽʽÿ�.", vbCritical
        Exit Sub
    End If
    If obj_fso.fileExists(TextBox_dart2_Path) = False Then
        MsgBox "���ϰ�ΰ� �߸��Ǿ����ϴ�. ��θ� �ٽ� Ȯ���Ͽ� �ֽʽÿ�.", vbCritical
        Exit Sub
    End If
    
    Call ExcelFile_Get_dart(TextBox_dart_Path) '.Text)
    Call ExcelFile_Get_dart2(TextBox_dart2_Path) '.Text)
    
    'Unload Me
    
End Sub

