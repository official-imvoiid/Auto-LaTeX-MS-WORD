Sub ConvertLatexToReadable()
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    Dim count As Integer
    count = 0
    Dim startPos As Long
    startPos = 0

    Do
        ' Find opening $
        Dim sr As Range
        Set sr = oDoc.Range(startPos, oDoc.Content.End)
        With sr.Find
            .ClearFormatting
            .text = "$"
            .Forward = True
            .Wrap = wdFindStop
            If Not .Execute Then Exit Do
        End With
        Dim p1 As Long
        p1 = sr.Start

        ' Find closing $
        Dim er As Range
        Set er = oDoc.Range(p1 + 1, oDoc.Content.End)
        With er.Find
            .ClearFormatting
            .text = "$"
            .Forward = True
            .Wrap = wdFindStop
            If Not .Execute Then Exit Do
        End With
        Dim p2 As Long
        p2 = er.End

        ' Get math text between the $ signs
        Dim mathText As String
        mathText = oDoc.Range(p1 + 1, p2 - 1).text

        ' Replace LaTeX commands with real characters
        mathText = Replace(mathText, "\phi", ChrW(966))
        mathText = Replace(mathText, "\Phi", ChrW(934))
        mathText = Replace(mathText, "\lambda", ChrW(955))
        mathText = Replace(mathText, "\Lambda", ChrW(923))
        mathText = Replace(mathText, "\alpha", ChrW(945))
        mathText = Replace(mathText, "\beta", ChrW(946))
        mathText = Replace(mathText, "\gamma", ChrW(947))
        mathText = Replace(mathText, "\delta", ChrW(948))
        mathText = Replace(mathText, "\theta", ChrW(952))
        mathText = Replace(mathText, "\pi", ChrW(960))
        mathText = Replace(mathText, "\sigma", ChrW(963))
        mathText = Replace(mathText, "\omega", ChrW(969))
        mathText = Replace(mathText, "\mu", ChrW(956))
        mathText = Replace(mathText, "\times", ChrW(215))
        mathText = Replace(mathText, "\cdot", ChrW(183))
        mathText = Replace(mathText, "\mod", " mod ")
        mathText = Replace(mathText, "\sqrt", ChrW(8730))
        mathText = Replace(mathText, "\infty", ChrW(8734))
        mathText = Replace(mathText, "\neq", ChrW(8800))
        mathText = Replace(mathText, "\leq", ChrW(8804))
        mathText = Replace(mathText, "\geq", ChrW(8805))
        mathText = Replace(mathText, "{", "")
        mathText = Replace(mathText, "}", "")

        ' Delete the full $...$ block
        oDoc.Range(p1, p2).Delete

        ' Insert the processed text with super/subscript formatting
        Call InsertFormatted(oDoc, p1, mathText)

        startPos = p1 + Len(mathText)
        If startPos >= oDoc.Content.End Then Exit Do
        count = count + 1
    Loop

    MsgBox "Done! " & count & " formulas converted."
End Sub

Sub InsertFormatted(oDoc As Document, insertAt As Long, text As String)
    Dim i As Integer
    Dim c As String
    Dim doSuper As Boolean
    Dim doSub As Boolean
    Dim curPos As Long

    doSuper = False
    doSub = False
    curPos = insertAt
    i = 1

    Do While i <= Len(text)
        c = Mid(text, i, 1)

        If c = "^" Then
            doSuper = True
            doSub = False
            i = i + 1

        ElseIf c = "_" Then
            doSub = True
            doSuper = False
            i = i + 1

        Else
            ' Insert the character
            Dim ins As Range
            Set ins = oDoc.Range(curPos, curPos)
            ins.InsertAfter c

            ' Apply super or subscript if needed
            If doSuper Or doSub Then
                Dim fmt As Range
                Set fmt = oDoc.Range(curPos, curPos + 1)
                If doSuper Then fmt.Font.Superscript = True
                If doSub Then fmt.Font.Subscript = True
                doSuper = False
                doSub = False
            End If

            curPos = curPos + 1
            i = i + 1
        End If
    Loop
End Sub

