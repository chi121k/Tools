Attribute VB_Name = "Module1"
Option Compare Database

Sub 全てのコードウインドウを閉じる()
    Dim c As CodePane
    For Each c In Application.VBE.CodePanes
        c.Window.Close
    Next
End Sub
