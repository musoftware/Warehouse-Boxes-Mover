Attribute VB_Name = "ModText"
Option Explicit
Public NumberNow As Integer
Public memoryzz As String
Const paranthesisX = 1
Const punctuation_markX = 2
Const exclamation_pointX = 3
Const All_MarksX = 4

Enum Type_Char
paranthesis = 1
punctuation_mark = 2
exclamation_point = 3
All_Marks = 4
End Enum
Enum Kind_Function
ReplaceType = 1
AddType = 2
RemoveType = 3
ReplaceBetweenType = 4
DeleteLineType = 5
RemoveCharType = 6
AddLeftRightType = 7
ReplaceAfterCharType = 8
ReplaceLineType = 9
AddLineType = 10
End Enum
Dim Split_TextT() As String
Dim Text, HHV
Public ArabSLang As Boolean

Public Function WordCount(strString As String)
    Dim LastWord As Long
    
    strString = Trim(strString)
    
    LastWord = 1
    Do While LastWord <> 0
        LastWord = InStr(LastWord + 1, strString, " ")
        WordCount = WordCount + 1
    Loop
    
End Function

Public Function ReplaceBinary(Expression, FindText, ReplaceText)

ReplaceBinary = Replace(Expression, FindText, ReplaceText, , , vbBinaryCompare)

End Function
Public Function ReplaceBetween(SubText, F1, F2, r)

Dim GT, RD, Td
GT = 1
RD = ""
ReplaceBetween = SubText
Do While InStr(GT, UCase(SubText), UCase(F1))
If InStr(GT, UCase(SubText), UCase(F1)) Then
If FindText(RD, "{&}" & Td & "{&}") = False Then
Td = Mid(SubText, InStr(GT, UCase(SubText), UCase(F1)), InStr(InStr(GT, UCase(SubText), UCase(F1)) + 1, UCase(SubText), UCase(F2)) - InStr(GT, UCase(SubText), UCase(F1)) + 1)
ReplaceBetween = Replace(ReplaceBetween, Td, F1 & r & F2)
RD = RD & "{&}" & Td & "{&}"
End If
End If
GT = InStr(GT, UCase(SubText), UCase(F1)) + 1 * 1
Loop

End Function

Public Function ReplaceTextIf(Expression, FindTextIf, FindText, ReplaceText, Optional NotFind As Boolean = True)

If NotFind = True Then
If InStr(1, Expression, FindTextIf, vbTextCompare) = 0 Then
ReplaceTextIf = ReplaceText(Expression, FindText, ReplaceText)
End If
Else
If InStr(1, Expression, FindTextIf, vbTextCompare) <> 0 Then
ReplaceTextIf = ReplaceText(Expression, FindText, ReplaceText)
End If
End If

End Function

Public Function AddTextIf(Expression, FindTextIf, AddText, Optional NotFind As Boolean = True)

AddTextIf = Expression
If NotFind = True Then
If InStr(1, Expression, FindTextIf, vbTextCompare) = 0 Then
AddTextIf = AddTextWithNewLine(Expression, AddText)
End If
Else
If InStr(1, Expression, FindTextIf, vbTextCompare) <> 0 Then
AddTextIf = AddTextWithNewLine(Expression, AddText)
End If
End If

End Function

Public Function XAddTextIf(Expression, FindTextIf, AddText, NotFind)

XAddTextIf = Expression
If NotFind = True Then
If InStr(1, Expression, FindTextIf, vbTextCompare) = 0 Then
XAddTextIf = AddTextWithNewLine(Expression, AddText)
End If
Else
If InStr(1, Expression, FindTextIf, vbTextCompare) <> 0 Then
XAddTextIf = AddTextWithNewLine(Expression, AddText)
End If
End If

End Function

Public Function RemoveCharX(ExpressionX, TypeCharX)

Dim numberGo, AllCharIs
AllCharIs = "',-.()!_$*<>/\?;:=+"
RemoveCharX = ExpressionX

If TypeCharX = All_Marks Then
For numberGo = 1 To Len(AllCharIs)
RemoveCharX = ReplaceText(RemoveCharX, Mid(AllCharIs, numberGo, 1), Empty)
Next
ElseIf TypeCharX = exclamation_pointX Then
RemoveCharX = ReplaceText(RemoveCharX, "!", Empty)
ElseIf TypeCharX = paranthesisX Then
RemoveCharX = ReplaceText(RemoveCharX, "(", Empty)
RemoveCharX = ReplaceText(RemoveCharX, ")", Empty)
RemoveCharX = ReplaceText(RemoveCharX, "{", Empty)
RemoveCharX = ReplaceText(RemoveCharX, "}", Empty)
ElseIf TypeCharX = punctuation_markX Then
RemoveCharX = ReplaceText(RemoveCharX, ",", Empty)
End If

End Function

Public Function ReplaceCharX(ExpressionX, TypeCharX As Type_Char, WithText)

Dim numberGo, AllCharIs
AllCharIs = "',-.()!_$*<>/\?;:=+"
ReplaceCharX = ExpressionX

If TypeCharX = All_Marks Then
For numberGo = 1 To Len(AllCharIs)
ReplaceCharX = ReplaceText(ReplaceCharX, Mid(AllCharIs, numberGo, 1), WithText)
Next
ElseIf TypeCharX = exclamation_pointX Then
ReplaceCharX = ReplaceText(ReplaceCharX, "!", WithText)
ElseIf TypeCharX = paranthesisX Then
ReplaceCharX = ReplaceText(ReplaceCharX, "(", WithText)
ReplaceCharX = ReplaceText(ReplaceCharX, ")", WithText)
ReplaceCharX = ReplaceText(ReplaceCharX, "{", WithText)
ReplaceCharX = ReplaceText(ReplaceCharX, "}", WithText)
ElseIf TypeCharX = punctuation_markX Then
ReplaceCharX = ReplaceText(ReplaceCharX, ",", WithText)
End If

End Function

Public Function RemoveChar(Expression, TypeChar)
'
Dim numberGo, AllCharIs
AllCharIs = "',-.()!_$*<>/\?;:=+[]{}'"
RemoveChar = Expression

If TypeChar = All_Marks Then
For numberGo = 1 To Len(AllCharIs)
RemoveChar = ReplaceText(RemoveChar, Mid(AllCharIs, numberGo, 1), Empty)
Next
ElseIf TypeChar = exclamation_point Then
RemoveChar = ReplaceText(RemoveChar, "!", Empty)
ElseIf TypeChar = paranthesis Then
RemoveChar = ReplaceText(RemoveChar, "(", Empty)
RemoveChar = ReplaceText(RemoveChar, ")", Empty)
RemoveChar = ReplaceText(RemoveChar, "{", Empty)
RemoveChar = ReplaceText(RemoveChar, "}", Empty)
ElseIf TypeChar = punctuation_mark Then
RemoveChar = ReplaceText(RemoveChar, ",", Empty)
End If
'
End Function


Public Function RemoveLineIfLine(Expression, TT)

Dim Split_Text() As String
Split_Text = Split(Expression, vbNewLine)
Dim yyy
For yyy = 0 To UBound(Split_Text)
If Split_Text(yyy) = TT Then Split_Text(yyy) = "{ThIS lINE mUST wILl bE dLETEDd}"
Next
RemoveLineIfLine = Join(Split_Text, vbNewLine)
RemoveLineIfLine = ReplaceBinary(RemoveLineIfLine, "{ThIS lINE mUST wILl bE dLETEDd}" & vbNewLine, Empty)
RemoveLineIfLine = ReplaceBinary(RemoveLineIfLine, vbNewLine & "{ThIS lINE mUST wILl bE dLETEDd}", Empty)
RemoveLineIfLine = ReplaceBinary(RemoveLineIfLine, vbNewLine & "{ThIS lINE mUST wILl bE dLETEDd}" & vbNewLine, Empty)
RemoveLineIfLine = ReplaceBinary(RemoveLineIfLine, "{ThIS lINE mUST wILl bE dLETEDd}", Empty)

End Function




Public Function AddTextAfternum(Expression, numofstr, Text)

On Error Resume Next
AddTextAfternum = Mid(Expression, 1, numofstr) & Text & Mid(Expression, numofstr + 1, Len(Expression))

End Function


Public Function AddTextAfternumLINES(Expression, numofstr, Text)
Dim Split_Text() As String, N
Split_Text = Split(Expression, vbNewLine)
For N = 0 To UBound(Split_Text)
Split_Text(N) = AddTextAfternum(Split_Text(N), numofstr, Text)
Next
AddTextAfternumLINES = Join(Split_Text, vbNewLine)

End Function

Public Function RemoveAfterChar(Expression, CharText)

Dim Split_Text() As String, N
Split_Text = Split(Expression, vbNewLine)
For N = 0 To UBound(Split_Text)
If Split_Text(N) <> "" Then
If FindText(Split_Text(N), CharText) Then
Split_Text(N) = Mid(Split_Text(N), 1, InStr(1, Split_Text(N), CharText, vbTextCompare) - 1)
End If
End If
Next
RemoveAfterChar = Join(Split_Text, vbNewLine)

End Function


Public Function RemoveAfterCharNotWith(Expression, CharText)

Dim Split_Text() As String, N
Split_Text = Split(Expression, vbNewLine)
For N = 0 To UBound(Split_Text)
If Split_Text(N) <> "" Then
If FindText(Split_Text(N), CharText) Then
Split_Text(N) = Mid(Split_Text(N), 1, InStr(1, Split_Text(N), CharText, vbTextCompare))
End If
End If
Next
RemoveAfterCharNotWith = Join(Split_Text, vbNewLine)

End Function

Public Function GetAfterChar(Expression, CharText) As String

GetAfterChar = Mid(Expression, InStr(1, Expression, CharText, vbTextCompare) + 1, Len(Expression))


End Function

Public Function RemoveBeforeChar(Expression, CharText)

Dim gg, LL
If Expression <> "" Then
gg = RemoveAfterChar(Expression, CharText)
LL = Len(Expression) - Len(gg)
RemoveBeforeChar = Right(Expression, LL)
Else
RemoveBeforeChar = Expression
End If
End Function



'
Public Function ReplaceAfterChar(Expression, CharText, ReplaceTextX)

Dim Split_Text() As String, N
Split_Text = Split(Expression, vbNewLine)
For N = 0 To UBound(Split_Text)
If FindText(Split_Text(N), CharText) Then
Split_Text(N) = ReplaceText(Split_Text(N), Mid(Split_Text(N), InStr(1, Split_Text(N), CharText, vbTextCompare) + 1, Len(Split_Text(N)) - InStr(1, Split_Text(N), CharText, vbTextCompare)), ReplaceTextX)
End If
Next
ReplaceAfterChar = Join(Split_Text, vbNewLine)

End Function
Public Function ReplaceTextinLine(Expression, FindTextX, ReplaceTextX, Line_Number)

Dim Split_Text() As String
Split_Text = Split(Expression, vbNewLine)
Split_Text(Line_Number - 1) = ReplaceText(Split_Text(Line_Number - 1), FindTextX, ReplaceTextX)
ReplaceTextinLine = Join(Split_Text, vbNewLine)

End Function

Public Function ReplaceTextinLineWhichConsist(Expression, InstrText, FindTextX, ReplaceTextX)

Dim Split_Text() As String, N
Split_Text = Split(Expression, vbNewLine)
For N = 0 To UBound(Split_Text)
If FindText(Split_Text(N), InstrText) Then
Split_Text(N) = ReplaceText(Split_Text(N), FindTextX, ReplaceTextX)
End If
Next
ReplaceTextinLineWhichConsist = Join(Split_Text, vbNewLine)

End Function

Public Function RemoveWord(Expression, Text)

RemoveWord = ReplaceText(Expression, Text, Empty)

End Function

Public Function AddTextWithNewLine(Expression, Text)

If Expression <> "" Then
If Right(Expression, 2) = vbNewLine Then
AddTextWithNewLine = Expression & Text
Else
AddTextWithNewLine = Expression & vbNewLine & Text
End If
Else
AddTextWithNewLine = Text
End If

End Function

Public Function FindText(Expression, Text, Optional start = 1) As Boolean
FindText = InStr(start, Expression, Text, vbTextCompare)
End Function

Public Function FindTextINLINE(Expression, Text, start As Boolean, Optional LeftOfLine = "", Optional RightOfLine = "")

Dim Split_Text() As String
Split_Text = Split(Expression, vbNewLine)
FindTextINLINE = ""
Dim FileSize, N
For N = 0 To UBound(Split_Text)
If Left(UCase(Split_Text(N)), Len(LeftOfLine)) = UCase(LeftOfLine) And Right(UCase(Split_Text(N)), Len(RightOfLine)) = UCase(RightOfLine) Then
    If start = True Then
        If Left(UCase(Split_Text(N)), Len(Text)) = UCase(Text) Then
            If FindText(Split_Text(N), Text) Then
                FindTextINLINE = FindTextINLINE & Split_Text(N) & vbNewLine
            End If
        End If
    Else
        If FindText(Split_Text(N), Text) Then
            FindTextINLINE = FindTextINLINE & Split_Text(N) & vbNewLine
        End If
    End If
    
End If
Next
FindTextINLINE = Left(FindTextINLINE, Len(FindTextINLINE) - 2)

End Function
Public Function XDoFunctionFind(TheExpression, F1Text, F2Text, start, KindOfFunction, String1, String2, String3, LeftOfLine, RightOfLine)

On Error Resume Next
Dim Split_TextT() As String, N
Split_TextT = Split(TheExpression, vbNewLine)
XDoFunctionFind = ""
For N = 0 To UBound(Split_TextT)
If Left(UCase(Split_TextT(N)), Len(LeftOfLine)) = UCase(LeftOfLine) And Right(UCase(Split_TextT(N)), Len(RightOfLine)) = UCase(RightOfLine) Then
    If start = True Then
        If Left(UCase(Split_TextT(N)), Len(Text)) = UCase(Text) Then
            If FindText(Split_TextT(N), F1Text) And FindText(Split_TextT(N), F2Text) Then
                If KindOfFunction = AddType Then
                Split_TextT(N) = Split_TextT(N) & String1
                End If
                If KindOfFunction = AddLeftRightType Then
                Split_TextT(N) = String1 & Split_TextT(N) & String2
                End If
                If KindOfFunction = RemoveType Then
                Split_TextT(N) = RemoveWord(Split_TextT(N), String1)
                End If
                If KindOfFunction = ReplaceType Then
                Split_TextT(N) = ReplaceText(Split_TextT(N), String1, String2)
                End If
                If KindOfFunction = ReplaceBetweenType Then
                Split_TextT(N) = ReplaceBetween(Split_TextT(N), String1, String2, String3)
                End If
                If KindOfFunction = DeleteLineType Then
                Split_TextT(N) = "{ThIS lINE mUST wILl bE dLETEDd}"
                End If
                If KindOfFunction = RemoveCharType Then
                Split_TextT(N) = RemoveCharX(Split_TextT(N), String1)
                End If
                If KindOfFunction = ReplaceAfterCharType Then
                Split_TextT(N) = ReplaceAfterChar(Split_TextT(N), String1, String2)
                End If
                If KindOfFunction = ReplaceLineType Then
                Split_TextT(N) = String1
                End If
                If KindOfFunction = AddLineType Then
                Split_TextT(N) = String1 & vbNewLine & Split_TextT(N)
                End If
            End If
        End If
    Else
        If FindText(Split_TextT(N), F1Text) And FindText(Split_TextT(N), F2Text) Then
                If KindOfFunction = AddType Then
                Split_TextT(N) = Split_TextT(N) & String1
                End If
                If KindOfFunction = AddLeftRightType Then
                Split_TextT(N) = String1 & Split_TextT(N) & String2
                End If
                If KindOfFunction = RemoveType Then
                Split_TextT(N) = RemoveWord(Split_TextT(N), String1)
                End If
                If KindOfFunction = ReplaceType Then
                Split_TextT(N) = ReplaceText(Split_TextT(N), String1, String2)
                End If
                If KindOfFunction = ReplaceBetweenType Then
                Split_TextT(N) = ReplaceBetween(Split_TextT(N), String1, String2, String3)
                End If
                If KindOfFunction = DeleteLineType Then
                Split_TextT(N) = "{ThIS lINE mUST wILl bE dLETEDd}"
                End If
                If KindOfFunction = RemoveCharType Then
                Split_TextT(N) = RemoveCharX(Split_TextT(N), String1)
                End If
                If KindOfFunction = ReplaceAfterCharType Then
                Split_TextT(N) = ReplaceAfterChar(Split_TextT(N), String1, String2)
                End If
                If KindOfFunction = ReplaceLineType Then
                Split_TextT(N) = String1
                End If
                If KindOfFunction = AddLineType Then
                Split_TextT(N) = String1 & vbNewLine & Split_TextT(N)
                End If
        End If
    End If
End If
    
Next
XDoFunctionFind = Join(Split_TextT, vbNewLine)
XDoFunctionFind = ReplaceBinary(XDoFunctionFind, vbNewLine & "{ThIS lINE mUST wILl bE dLETEDd}" & vbNewLine, Empty)
XDoFunctionFind = ReplaceBinary(XDoFunctionFind, "{ThIS lINE mUST wILl bE dLETEDd}" & vbNewLine, Empty)
XDoFunctionFind = ReplaceBinary(XDoFunctionFind, vbNewLine & "{ThIS lINE mUST wILl bE dLETEDd}", Empty)
XDoFunctionFind = ReplaceBinary(XDoFunctionFind, "{ThIS lINE mUST wILl bE dLETEDd}", Empty)

End Function



Public Function ReplaceText(ExpressionT, FindTextT, ReplaceTextT)

ReplaceText = Replace(ExpressionT, FindTextT, ReplaceTextT, , , vbTextCompare)
End Function


Public Function ReplaceFromTo(TextHere, TextFunc, TFrom, TTo)

Dim BEGIND As Boolean, ENDD As Boolean, TEXTG As String, ngG

BEGIND = False
ENDD = False
Split_TextT = Split(TextHere, vbNewLine)
Dim FileSize, N
For N = 0 To UBound(Split_TextT)
Split_TextT = Split(TextHere, vbNewLine)
If Mid(Split_TextT(N), 1, Len(TTo)) = TTo Then
ENDD = True
End If
If Mid(Split_TextT(N), 1, Len(TFrom)) = TFrom Then
BEGIND = True
End If
If ENDD = True Then TEXTG = TEXTG & Split_TextT(N): ReplaceFromTo = ModText.ReplaceText(TextHere, TEXTG, TextFunc): TEXTG = "": BEGIND = False: ENDD = False
If BEGIND = True Then TEXTG = TEXTG & Split_TextT(N) & vbNewLine
Next

End Function


Public Function AddinEveryLineinLeft(E, T)

Dim N, Split_TextT() As String
Split_TextT = Split(E, vbNewLine)

For N = 0 To UBound(Split_TextT)
Split_TextT(N) = T & Split_TextT(N)
Next
AddinEveryLineinLeft = Join(Split_TextT, vbNewLine)

End Function

Public Function AddinEveryLineinRight(E, T)

Dim N
Split_TextT = Split(E, vbNewLine)

For N = 0 To UBound(Split_TextT)
Split_TextT(N) = Split_TextT(N) & T
Next
AddinEveryLineinRight = Join(Split_TextT, vbNewLine)

End Function

Public Function AddinEveryLineinRightandNotFind(Ek, Fk As String, TextinFind As String, TextinnotFind As String) As String

Dim NNHN, Split_TextTa() As String
Split_TextTa = Split(Ek, vbNewLine)

For NNHN = 0 To UBound(Split_TextTa)
If InStr(1, Split_TextTa(NNHN), Fk, vbTextCompare) <> 0 Then
Split_TextTa(NNHN) = Split_TextTa(NNHN) & TextinFind
Else
Split_TextTa(NNHN) = Split_TextTa(NNHN) & TextinnotFind
End If
Next
AddinEveryLineinRightandNotFind = Join(Split_TextTa, vbNewLine)

End Function

Public Function HexO(TT)

Dim GGV() As Byte, N

GGV = TT
HexO = Empty
For N = 0 To UBound(GGV) Step 2
HexO = HexO & Zero(Hex(GGV(N)), 2) & " "
Next

'Dim NHG
'HexO = ""
'For NHG = 1 To Len(tt)
'HexO = HexO & Format(Hex(Asc(Mid(tt, NHG, 1))), "@@") & " "
'If NHG Mod 16 = 0 Then HexO = HexO & vbNewLine
'Next
'On Error GoTo HEE:
'Dim NHG, HHG() As String
'HexO = TT
'For NHG = 2 To 255 Step 1
'HexO = Replace(HexO, Chr(NHG), Chr(NHG) & Chr(1))
'Next
'
'HexO = Left(HexO, Len(HexO) - 1)
'HHG = Split(HexO, Chr(1))
'
'
'For NHG = 0 To UBound(HHG)
'If Len(HEX(Asc(HHG(NHG)))) = 1 Then
'HHG(NHG) = "0" & HEX(Asc(HHG(NHG)))
'Else
'HHG(NHG) = HEX(Asc(HHG(NHG)))
'End If
'Next
'HexO = Join(HHG, " ")
'
'HexO = HexO & " "
'Exit Function
'HEE:
'HexO = ""

End Function

Public Function UnHexO(TT)


Dim NHG
UnHexO = Replace(TT, vbNewLine, "")
For NHG = 255 To 1 Step -1
If Len(Hex(NHG)) = 1 Then
UnHexO = Replace(UnHexO, "0" & Hex(NHG) & " ", Chr(NHG))
Else
UnHexO = Replace(UnHexO, Hex(NHG) & " ", Chr(NHG))
End If

'UnHexO = Replace(UnHexO, Format(Hex(NHG), "@@") & " ", Chr(NHG))
Next

End Function

Public Function TEXTAddLine(SubText, Txt1)
Dim Spilt_Text() As String
Dim n1, n2
Spilt_Text = Split(SubText, vbNewLine)
n1 = LBound(Spilt_Text)
n2 = UBound(Spilt_Text)
ReDim Preserve Spilt_Text(n1 To n2 * 1 + 1 * 1)
Spilt_Text(n2 * 1 + 1 * 1) = Txt1
TEXTAddLine = Join(Spilt_Text, vbNewLine)
ReDim Spilt_Text(0 To 0)

End Function

Public Function TEXTChangeLine(SubText, Txt1, num1)
Dim Spilt_Text() As String
Spilt_Text = Split(SubText, vbNewLine)
Spilt_Text(num1) = Txt1
TEXTChangeLine = Join(Spilt_Text, vbNewLine)
ReDim Spilt_Text(0 To 0)
End Function

Public Function TEXTAddLineWithNum(SubText, Txt1, num1)
Dim Spilt_Text() As String
Spilt_Text = Split(SubText, vbNewLine)
Spilt_Text(num1) = Txt1 & vbNewLine & Spilt_Text(num1)

TEXTAddLineWithNum = Join(Spilt_Text, vbNewLine)
ReDim Spilt_Text(0 To 0)
End Function

Public Function TEXTRemoveLine(Expression, Line_Number)
Dim Split_Text() As String
Split_Text = Split(Expression, vbNewLine)
Split_Text(Line_Number - 1) = "cc>{-STOP-}<cc"

TEXTRemoveLine = Join(Split_Text, vbNewLine)
TEXTRemoveLine = ReplaceBinary(TEXTRemoveLine, "cc>{-STOP-}<cc" & vbNewLine, Empty)
TEXTRemoveLine = ReplaceBinary(TEXTRemoveLine, vbNewLine & "cc>{-STOP-}<cc", Empty)
TEXTRemoveLine = ReplaceBinary(TEXTRemoveLine, vbNewLine & "cc>{-STOP-}<cc" & vbNewLine, Empty)
ReDim Spilt_Text(0 To 0)
End Function

Public Function DelspaceinLeftineveryline(EE)

Dim Split_TextT() As String, N
Split_TextT = Split(EE, vbNewLine)
For N = 0 To UBound(Split_TextT)
If Split_TextT(N) <> "" Then _
Split_TextT(N) = LTrim(Split_TextT(N))
Next
DelspaceinLeftineveryline = Join(Split_TextT, vbNewLine)

End Function
Public Function DelspaceinRightineveryline(EE)
Dim Split_TextT() As String, N
Split_TextT = Split(EE, vbNewLine)
For N = 0 To UBound(Split_TextT)
If Split_TextT(N) <> "" Then _
Split_TextT(N) = RTrim(Split_TextT(N))
Next
DelspaceinRightineveryline = Join(Split_TextT, vbNewLine)
End Function

Public Function DelspaceinTwoineveryline(EE)

Dim Split_TextT() As String, N
Split_TextT = Split(EE, vbNewLine)
For N = 0 To UBound(Split_TextT)
If Split_TextT(N) <> "" Then _
Split_TextT(N) = Trim(Split_TextT(N))

Next
DelspaceinTwoineveryline = Join(Split_TextT, vbNewLine)

End Function

Public Function ReplacebeforeChar(Expression, CharText, ReplaceTextX)

Dim Split_Text() As String, N
Split_Text = Split(Expression, vbNewLine)
For N = 0 To UBound(Split_Text)
If FindText(Split_Text(N), CharText) Then
Split_Text(N) = ReplaceText(Split_Text(N), Mid(Split_Text(N), 1, InStr(1, Split_Text(N), CharText, vbTextCompare) - 1), ReplaceTextX)
End If
Next
ReplacebeforeChar = Join(Split_Text, vbNewLine)

End Function

Public Function SSSLangText(Text, Lane)
SSSLangText = Text
If Lane = 1 Then
SSSLangText = Replace(SSSLangText, "Node", "Note")
End If
End Function

Public Function Removetext(Text, Lane)
Removetext = Replace(Text, Lane, "")
End Function

Public Function RemoveCharfromLeft(Text, Number)
If Text <> "" Then
RemoveCharfromLeft = Right(Text, Len(Text) - Number)
Else
RemoveCharfromLeft = Text
End If
End Function
Public Function RemoveCharfromRight(Text, Number)
RemoveCharfromRight = Left(Text, Len(Text) - Number)
End Function

Public Function RemoveCharfromLeftLine(Text, Number)
Dim Split_TextT() As String, N
Split_TextT = Split(Text, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = RemoveCharfromLeft(Split_TextT(N), Number)
Next
RemoveCharfromLeftLine = Join(Split_TextT, vbNewLine)
End Function
Public Function RemoveCharfromRightLine(Text, Number)
Dim Split_TextT() As String, N
Split_TextT = Split(Text, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = RemoveCharfromRight(Split_TextT(N), Number)
Next
RemoveCharfromRightLine = Join(Split_TextT, vbNewLine)
End Function


Public Function ReplaceBetweenL(Text, Y, Z, r)
Dim Split_TextT() As String, N
Split_TextT = Split(Text, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = ReplaceBetween(Split_TextT(N), Y, Z, r)
Next
ReplaceBetweenL = Join(Split_TextT, vbNewLine)
End Function

Public Function RemoveCharRight(Text, NH1)
RemoveCharRight = Text

If InStr(1, Text, NH1, vbTextCompare) Then
RemoveCharRight = Left(Text, Len(Text) - InStr(1, StrReverse(Text), NH1, vbTextCompare))
End If

End Function
Public Function RemoveCharLeftOnly(Text, NH)
RemoveCharLeftOnly = Text
If FindText(Text, NH) Then
RemoveCharLeftOnly = Right(RemoveCharLeftOnly, Len(RemoveCharLeftOnly) - (InStr(1, RemoveCharLeftOnly, NH, vbTextCompare) - 1))
End If
End Function

Public Function RemoveCharRightOnly(Text, NH)
RemoveCharRightOnly = Text
If FindText(Text, NH) Then

RemoveCharRightOnly = Left(RemoveCharRightOnly, Len(RemoveCharRightOnly) - (InStr(1, StrReverse(RemoveCharRightOnly), NH, vbTextCompare) - 1))
End If

End Function
Public Function RemoveCharLeft(Text, NH)
RemoveCharLeft = Text
If FindText(Text, NH) Then
RemoveCharLeft = Right(RemoveCharLeft, Len(RemoveCharLeft) - InStr(1, RemoveCharLeft, NH, vbTextCompare))
End If
End Function

Public Function RemoveCharLeftAll(Text, NH)
RemoveCharLeftAll = Text
If FindText(Text, NH) Then
Do While FindText(RemoveCharLeftAll, NH)
RemoveCharLeftAll = RemoveCharLeft(RemoveCharLeftAll, NH)
Loop
End If
End Function

Public Function RemoveCharRightAll(Text, NH)
RemoveCharRightAll = Text
If FindText(Text, NH) Then
Do While FindText(RemoveCharRightAll, NH)
RemoveCharRightAll = RemoveCharRight(RemoveCharRightAll, NH)
Loop
End If
End Function



Public Function RemoveCharRightLine(ExpressionA, NH, Line_Number)

Dim Split_Text() As String

Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
Else
Split_Text(Line_Number - 1) = RemoveCharRight(Split_Text(Line_Number - 1), NH)
RemoveCharRightLine = Join(Split_Text, vbNewLine)
End If
End Function

Public Function RemoveCharLeftLine(ExpressionA, NH, Line_Number)

Dim Split_Text() As String

Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
Else
Split_Text(Line_Number - 1) = RemoveCharLeft(Split_Text(Line_Number - 1), NH)
RemoveCharLeftLine = Join(Split_Text, vbNewLine)
End If
End Function

Public Function EditTEXT(ExpressionA, Line_Number, RightT, LeftT)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
Else
If Split_TextT(Line_Number - 1) <> "" Then _
Split_Text(Line_Number - 1) = LeftT & Split_Text(Line_Number - 1) & RightT
EditTEXT = Join(Split_Text, vbNewLine)
End If
End Function

Public Function RemoveCharfromLeftf(ExpressionA, num, Line_Number)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
Else
Split_Text(Line_Number - 1) = RemoveCharfromLeft(Split_Text(Line_Number - 1), num)
RemoveCharfromLeftf = Join(Split_Text, vbNewLine)
End If
End Function

Public Function RemoveCharfromRightf(ExpressionA, num, Line_Number)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
Else
Split_Text(Line_Number - 1) = RemoveCharfromRight(Split_Text(Line_Number - 1), num)
RemoveCharfromRightf = Join(Split_Text, vbNewLine)
End If
End Function


Public Function RepeatLine(ExpressionA, Line_Number, num1)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
Else
Split_Text(Line_Number - 1) = RepeatT(Split_Text(Line_Number - 1), num1)
RepeatLine = Join(Split_Text, vbNewLine)
End If
End Function

Public Function RepeatT(ExpressionA, num)
RepeatT = ExpressionA
Dim gg
For gg = 1 To num
RepeatT = AddTextWithNewLine(RepeatT, ExpressionA)
Next
End Function

Public Function AddTextFirstLine(ExpressionA, Line_Subject, Line_Number)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
AddTextFirstLine = AddTextFirstLine & Line_Subject & vbNewLine
Else
Split_Text(Line_Number - 1) = Line_Subject & Split_Text(Line_Number - 1)
AddTextFirstLine = Join(Split_Text, vbNewLine)
End If
End Function

Public Function AddTextLastLine(ExpressionA, Line_Subject, Line_Number)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
AddTextLastLine = AddTextLastLine & Line_Subject & vbNewLine
Else
Split_Text(Line_Number - 1) = Split_Text(Line_Number - 1) & Line_Subject
AddTextLastLine = Join(Split_Text, vbNewLine)
End If
End Function

Public Function AddTextWithoutNewLine(Expression, Text)
AddTextWithoutNewLine = Expression & Text
End Function

Public Function AddInLeftlines(ExpressionA, Line_Subject)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
AddInLeftlines = AddInLeftlines & Line_Subject & vbNewLine
Else
Dim NNN
For NNN = 0 To UBound(Split_Text)
Split_Text(NNN) = Line_Subject & Split_Text(NNN)
AddInLeftlines = Join(Split_Text, vbNewLine)
Next
End If
End Function
Public Function AddInRightlines(ExpressionA, Line_Subject)
Dim Split_Text() As String
Split_Text = Split(ExpressionA, vbNewLine)
If UBound(Split_Text) = -1 Then
AddInRightlines = AddInRightlines & Line_Subject & vbNewLine
Else
Dim NNN
For NNN = 0 To UBound(Split_Text)
Split_Text(NNN) = Split_Text(NNN) & Line_Subject
AddInRightlines = Join(Split_Text, vbNewLine)
Next
End If
End Function



Public Function AddInLeftChar(SubText, Char, Text)
AddInLeftChar = Replace(SubText, Char, Text & Char)
End Function

Public Function AddInRightChar(SubText, Char, Text)
AddInRightChar = Replace(SubText, Char, Char & Text)
End Function

Public Function AddInLeftCharLines(SubText, Char, Text)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = AddInLeftChar(Split_TextT(N), Char, Text)
Next
AddInLeftCharLines = Join(Split_TextT, vbNewLine)
End Function

Public Function AddInRightCharLines(SubText, Char, Text)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = AddInRightChar(Split_TextT(N), Char, Text)
Next
AddInRightCharLines = Join(Split_TextT, vbNewLine)
End Function

Public Function RemoveCharfromLeftBF(SubText, num, Char)
Dim GBB() As String
If SubText <> "" Then
GBB = Split(SubText, Char)

GBB(0) = RemoveCharfromLeft(GBB(0), num)

RemoveCharfromLeftBF = Join(GBB, Char)
Else
RemoveCharfromLeftBF = SubText
End If
End Function

Public Function RemoveCharfromRightBF(SubText, num, Char)
Dim GBB() As String
GBB = Split(SubText, Char)

GBB(0) = RemoveCharfromRight(GBB(0), num)

RemoveCharfromRightBF = Join(GBB, Char)
End Function

Public Function RemoveCharfromLeftAF(SubText, num, Char)
Dim GBB() As String
If SubText <> "" Then
GBB = Split(SubText, Char)
GBB(1) = RemoveCharfromLeft(GBB(1), num)
RemoveCharfromLeftAF = Join(GBB, Char)
Else
RemoveCharfromLeftAF = ""
End If
End Function

Public Function RemoveCharfromRightAF(SubText, num, Char)
Dim GBB() As String
GBB = Split(SubText, Char)

GBB(1) = RemoveCharfromRight(GBB(1), num)

RemoveCharfromRightAF = Join(GBB, Char)
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function RemoveCharfromLeftBFLines(SubText, num, Char)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = RemoveCharfromLeftBF(Split_TextT(N), Char, Text)
Next
RemoveCharfromLeftBFLines = Join(Split_TextT, vbNewLine)
End Function

Public Function RemoveCharfromRightBFLines(SubText, num, Char)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = RemoveCharfromRightBF(Split_TextT(N), Char, Text)
Next
RemoveCharfromRightBFLines = Join(Split_TextT, vbNewLine)
End Function

Public Function RemoveCharfromLeftAFLines(SubText, num, Char)
Dim Split_TextT() As String, N
If SubText <> "" Then
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = RemoveCharfromLeftAF(Split_TextT(N), Char, Text)
Next
RemoveCharfromLeftAFLines = Join(Split_TextT, vbNewLine)
Else
RemoveCharfromLeftAFLines = ""
End If
End Function

Public Function RemoveCharfromRightAFLines(SubText, num, Char)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = RemoveCharfromRightAF(Split_TextT(N), Char, Text)
Next
RemoveCharfromRightAFLines = Join(Split_TextT, vbNewLine)
End Function


Public Function NumberingLeftLines(SubText, start)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = (N * 1 + start * 1) & Split_TextT(N)
Next
NumberingLeftLines = Join(Split_TextT, vbNewLine)
End Function

Public Function NumberingRightLines(SubText, start)
Dim Split_TextT() As String, N
Split_TextT = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_TextT)
Split_TextT(N) = Split_TextT(N) & (N * 1 + start * 1)
Next
NumberingRightLines = Join(Split_TextT, vbNewLine)
End Function



Public Function NumberingRight(SubText)
NumberingRight = SubText & NumberNow
NumberNow = NumberNow * 1 + 1 * 1
End Function

Public Function NumberingLeft(SubText)
NumberingLeft = NumberNow & SubText
NumberNow = NumberNow * 1 + 1 * 1
End Function

Public Sub SetNum(start)
NumberNow = start
End Sub

Public Function RemoveCharLeftLines(SubText, hN)
Dim Split_Texto() As String, N
Split_Texto = Split(SubText, vbNewLine)
For N = 0 To UBound(Split_Texto)
If FindText(Split_Texto(N), hN) Then _
Split_Texto(N) = RemoveCharLeft(Split_Texto(N), hN)
Next
RemoveCharLeftLines = Join(Split_Texto, vbNewLine)
End Function

Public Function RemoveCharRightLines(SubText, hN)
Dim Split_Texto() As String, Nv
Split_Texto = Split(SubText, vbNewLine)
For Nv = 0 To UBound(Split_Texto)
If FindText(Split_Texto(Nv), hN) Then _
Split_Texto(Nv) = RemoveCharRight(Split_Texto(Nv), hN)
Next
RemoveCharRightLines = Join(Split_Texto, vbNewLine)
End Function

Public Function RemoveCharRightLinesAll(SubText, hN)
Dim Split_Texto() As String, Nv
Split_Texto = Split(SubText, vbNewLine)
For Nv = 0 To UBound(Split_Texto)
If FindText(Split_Texto(Nv), hN) Then _
Split_Texto(Nv) = RemoveCharRightAll(Split_Texto(Nv), hN)
Next
RemoveCharRightLinesAll = Join(Split_Texto, vbNewLine)
End Function

Public Function RemoveCharLeftLinesAll(SubText, hN)
Dim Split_Texto() As String, Nv
Split_Texto = Split(SubText, vbNewLine)
For Nv = 0 To UBound(Split_Texto)
If FindText(Split_Texto(Nv), hN) Then _
Split_Texto(Nv) = RemoveCharLeftAll(Split_Texto(Nv), hN)
Next
RemoveCharLeftLinesAll = Join(Split_Texto, vbNewLine)
End Function


Public Function StringzeroNum(Text, num, StringX)
If num > Len(Text) Then
StringzeroNum = String(num - Len(Text), StringX) & Text
ElseIf num = Len(Text) Then
StringzeroNum = Text
ElseIf num < Len(Text) Then
StringzeroNum = Text
End If
End Function


Public Function Zero(vv, num)
On Error Resume Next
Zero = vv
Zero = String(num - Len(vv), "0") & vv

End Function

Public Function Numbba(f, T, s)
Dim NNN
Numbba = ""
For NNN = Trim(f) To Trim(T) Step s
Numbba = Numbba & NNN & vbNewLine
Next
Numbba = Left(Numbba, Len(Numbba) - 2)
End Function

Public Function DefragMeent(Subb, MM, Chha)
On Error Resume Next
Dim Spp() As String, Spp2() As String, N
Spp = Split(Subb, vbNewLine)
Spp2 = Split(MM, vbNewLine)
DefragMeent = ""
If UBound(Spp) > UBound(Spp2) Then
For N = 0 To UBound(Spp)
DefragMeent = DefragMeent & Spp(N) & Chha & Spp2(N) & vbNewLine
Next
Else
For N = 0 To UBound(Spp2)
DefragMeent = DefragMeent & Spp(N) & Chha & Spp2(N) & vbNewLine
Next
End If
If Right(DefragMeent, 2) = vbNewLine Then
DefragMeent = Left(DefragMeent, Len(DefragMeent) - 2)
End If

End Function
