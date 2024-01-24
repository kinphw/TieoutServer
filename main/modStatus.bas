Attribute VB_Name = "modStatus"
Option Explicit

Dim strFile As String

Function ReadStatus()
    
    'Dim strFile As String
    'strFile = ThisWorkbook.Path & "\status.txt"
    strFile = "C:\xampp\htdocs\status.txt"

    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fileRead As Object
    Set fileRead = FSO.OpenTextFile(strFile, ForReading) '텍스트 파일의 경로를 입력합니다
    
    Dim strMsg As String
    strMsg = fileRead.ReadAll
    
    fileRead.Close
    
    Debug.Print strMsg
    ReadStatus = strMsg

End Function

Sub WriteStatus(strMsg As String)

    strFile = "C:\xampp\htdocs\status.txt"
    
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fileWrite As Object
    Set fileWrite = FSO.OpenTextFile(strFile, ForWriting)

    fileWrite.Write strMsg
    fileWrite.Close

End Sub
