Attribute VB_Name = "Module2"
Sub dump(bestand As String, txt As String)
    Open "C:\Users\bas.kerkhof\Desktop\" & bestand For Output As #1
    Print #1, txt
    Close #1
End Sub


