Function StyleExists(styleName As String) As Boolean
    On Error Resume Next
    Dim s As Style
    Set s = ActiveDocument.Styles(styleName)
    StyleExists = Not s Is Nothing
    On Error GoTo 0
End Function

Sub ApplyStylesBasedOnText()
    Dim para As Paragraph
    Dim prefixText As String
    Dim textAfterPrefix As String
    Dim styleNum As Integer
    Dim paragraphsToStyle As New Collection

    ' First, identify paragraphs and their intended styles
    For Each para In ActiveDocument.Paragraphs
        prefixText = Left(Trim(para.Range.Text), 5)
        If IsNumeric(prefixText) And Len(prefixText) = 5 And (prefixText Like String(5, Left(prefixText, 1))) Then
            styleNum = Val(Left(prefixText, 1)) ' Extract the style number
            textAfterPrefix = Trim(Mid(para.Range.Text, 6))
            paragraphsToStyle.Add Array(para, "Style" & styleNum, textAfterPrefix)
        End If
    Next para

    ' Then, apply styles and text modifications
    Dim item As Variant
    For Each item In paragraphsToStyle
        Set para = item(0)
        Dim styleName As String
        styleName = CStr(item(1)) ' Ensure the style name is treated as a string
        If StyleExists(styleName) Then
            para.Range.Text = item(2)
            para.Range.Style = ActiveDocument.Styles(styleName)
        Else
            MsgBox styleName & " does not exist. Check style names and numbers."
        End If
    Next item
End Sub


====================
11111 - Style 1
Para
22222
Para - style2
