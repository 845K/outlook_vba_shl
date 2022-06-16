Attribute VB_Name = "Module3"
Sub CapsLockVergeten()

    Dim tekst As String
    Dim GetSelectedTextOnActiveWindow As String
    
    Select Case TypeName(Application.ActiveWindow)
    Case "Explorer"
        SendKeys "^c", True
        tekst = Clipboard()
        Clipboard (inverseCaps(tekst))
        SendKeys "^v", True
        
    Case "Inspector"
        GetSelectedTextOnActiveWindow = ActiveInspector.HTMLEditor.Selection.createRange.text
    Case Else
        MsgBox "Wat doe je?!"
    End Select
    

End Sub


Function Clipboard(Optional StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim x As Variant

'Store as variant for 64-bit VBA support
  x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With

End Function

Function swapCase(l As String) As String

    If UCase(l) = l Then
        swapCase = LCase(l)
        Exit Function
    End If
    
    If LCase(l) = l Then
        swapCase = UCase(l)
    Else
        swapCase = l
    End If
    
    
End Function

Function inverseCaps(txt As String) As String

    For i = 1 To Len(txt)
        inverseCaps = inverseCaps + swapCase(Mid(txt, i, 1))
    Next i

End Function
