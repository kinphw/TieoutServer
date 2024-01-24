Attribute VB_Name = "modHandler"
Sub Main()

Debug.Print "Begin"
Call modStatus.WriteStatus("Tie-out : Begin")

Application.ScreenUpdating = False

Debug.Print "File Loading"
Call modStatus.WriteStatus("Tie-out : File Loading")
Call modTieoutForm.Begin

Debug.Print "Matching Tables"
Call modStatus.WriteStatus("Tie-out : Matching Tables")
Call modTieoutMod1.MatchTables

Debug.Print "checking Strings"
Call modStatus.WriteStatus("Tie-out : Checking Strings")
Call modTieoutMod1.CopyNoBorderCells

Application.ScreenUpdating = True

Debug.Print "All Done"
Call modStatus.WriteStatus("Tie-out : All Done")

ActiveWorkbook.Save
Call ActiveWorkbook.Close(SaveChanges:=True)
Application.Quit

End Sub


Sub TestStatus()

Call modStatus.WriteStatus("Begin")
Call modStatus.ReadStatus

End Sub
