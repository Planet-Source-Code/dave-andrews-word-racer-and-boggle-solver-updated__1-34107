Attribute VB_Name = "modAnagram"
Option Explicit
Public DLoaded As Boolean
Public WChecked As Boolean
Public InCheck As Boolean
Dim Letters As String
Type DictList
    Word() As String
End Type
Dim Dict(25) As DictList
Public Numbers() As String
Public Grams() As String
Public AllWords() As String
Public XCount As Long
Dim nCount As Long
Function adv(Root As Integer, Val As Integer, MAX As Integer, start As Integer) As Integer
adv = Root + Val
If adv > MAX Then adv = start
End Function



Sub CheckAllWords()
ReDim Grams(0) As String
Dim i As Long
Dim j As Long
For j = 0 To UBound(AllWords)
    If IsWord(AllWords(j)) Then
        ReDim Preserve Grams(i) As String
        Grams(i) = AllWords(j)
        i = i + 1
    End If
Next j

End Sub

Sub CheckGram(MyText As String)
GetLetters MyText
MakeWords Letters
End Sub
Function Factorial(RootNum As Long) As Long
Dim i As Long
Dim Answer As Long
Answer = RootNum
For i = RootNum To 2 Step -1
    Answer = Answer * (i - 1)
Next i
Factorial = Answer
End Function

Sub GenNumList(MyText As String)
ReDim Numbers(Len(MyText) ^ (Len(MyText))) As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim M As Long
Dim P As Integer
Dim T As Integer
Dim temp As String
i = 1
j = 1

Numbers(0) = 1
nCount = 1
For i = 2 To Len(MyText)
    Numbers(j) = "1" & i
    j = j + 1
    nCount = nCount + 1
Next i
PlaceNum j, 1, 1 * (Len(MyText) - 1), Len(MyText) - 2, 1
M = nCount
k = M - 1
For i = 1 To Len(MyText) - 1
    For j = 0 To k
        temp = ""
        For P = 1 To Len(Numbers(j))
            T = Mid(Numbers(j), P, 1) + i
            If T > Len(MyText) Then T = T - Len(MyText)
            temp = temp & T
        Next P
        Numbers(M) = temp
        M = M + 1
    Next j
Next i
ReDim Preserve Numbers(M - 1) As String
End Sub
Sub GetLetters(MyText As String)
Letters = ""
Dim i As Integer
Dim MyChar As String * 1
For i = 1 To Len(MyText)
    MyChar = UCase(Mid(MyText, i, 1))
    If Asc(MyChar) >= 65 And Asc(MyChar) <= 90 Then
        Letters = Letters & MyChar
    End If
Next i
End Sub

Function initMSWord()
'Set objMSWord = New Word.Application
End Function

Function InList(MyList() As String, Check As String) As Boolean
Dim i As Long
For i = 0 To UBound(MyList)
    If MyList(i) = Check Then
        InList = True
        Exit Function
    End If
Next i
InList = False
End Function




Function IsWord(ByVal MyText As String) As Boolean
Dim dNum As Integer
Dim i As Integer
Dim UB As Long
Dim LB As Long
Dim M As Long
MyText = UCase(MyText)
If MyText = "" Then Exit Function
dNum = Asc(Mid(MyText, 1, 1)) - 65
LB = LBound(Dict(dNum).Word)
UB = UBound(Dict(dNum).Word)
M = (UB - (LB - 1)) \ 2 + LB
'Binary search
Do While UB <> LB
    If Dict(dNum).Word(M) = MyText Then IsWord = True: Exit Function
    If Dict(dNum).Word(M) > MyText Then
        UB = M
    ElseIf Dict(dNum).Word(M) < MyText Then
        LB = M
    End If
    M = (UB - (LB - 1)) \ 2 + LB
    If LB = M Or UB = 0 Then IsWord = False: Exit Function
    If UB - LB = 1 Then Exit Do
Loop
If Dict(dNum).Word(LB) = MyText Then IsWord = True: Exit Function
If Dict(dNum).Word(UB) = MyText Then IsWord = True: Exit Function
'For i = 0 To UBound(Dict(dNum).Word)
'    If Dict(dNum).Word(i) = MyText Then
'        IsWord = True
'        Exit Function
'    End If
'Next i
'IsWord = False
End Function



Sub LoadDict()
Dim dName As String
Dim i As Long
Dim lNum As Integer
Dim ff As Integer
ff = FreeFile
For lNum = 0 To 25
    i = 0
    dName = App.Path & "\Dictionary\" & Chr(lNum + 65) & ".dic"
    ReDim Dict(lNum).Word(i)
    Open dName For Input As #ff
    Do While Not EOF(ff)
        ReDim Preserve Dict(lNum).Word(i)
        Input #ff, Dict(lNum).Word(i)
        i = i + 1
    Loop
    Close #ff
Next lNum
DLoaded = True
InCheck = False
End Sub
Sub MakeSuggestions(MyWord As String, ByRef sList As ListBox)
On Error GoTo dTrap
WChecked = False
InCheck = True
sList.Clear
'aggrivation
'if at least half of the letters are in the proper order,
'then we should add the word to the list of suggestions
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim tWord As String
Dim qWord As String
If IsWord(MyWord) Then sList.AddItem MyWord
For i = 0 To 25
    For j = 0 To UBound(Dict(i).Word)
        If WChecked Then GoTo dTrap
        tWord = LCase(Dict(i).Word(j))
        If Len(tWord) - Len(MyWord) > (Len(MyWord) / 2) Then GoTo SkipWord
        If UCase(tWord) = UCase(MyWord) Then GoTo SkipWord
        For k = 1 To Len(MyWord)
            qWord = ""
            If k > 1 Then qWord = Mid(MyWord, 1, k - 1)
            qWord = qWord & "*" & Mid(MyWord, k, 1) & "*" & Mid(MyWord, k + 1, Len(MyWord))
            'qWord = MyWord
            'Mid(qWord, k, 1) = "*"
            If UCase(tWord) Like UCase(qWord) Then
                If Asc(Left(MyWord, 1)) > 90 Then
                    tWord = LCase(tWord)
                ElseIf Asc(Right(MyWord, 1)) <= 90 Then
                    tWord = UCase(tWord)
                ElseIf Asc(Left(MyWord, 1)) <= 90 Then
                    Mid(tWord, 1, 1) = UCase(Left(tWord, 1))
                End If
                sList.AddItem tWord
            End If
        Next k
SkipWord:
    DoEvents
    Next j
Next i
dTrap:
    InCheck = False
End Sub

Sub MakeWords(MyText As String)
GenNumList MyText
Dim i As Long
Dim j As Integer
Dim k As Long
ReDim AllWords(0) As String
Dim temp As String
For i = 0 To UBound(Numbers)
    temp = ""
    For j = 1 To Len(Numbers(i))
        temp = temp & Mid(MyText, CInt(Mid(Numbers(i), j, 1)), 1)
    Next j
    ReDim Preserve AllWords(k) As String
    AllWords(k) = temp
    k = k + 1
Next i
End Sub

Sub PlaceNum(VarNum As Long, SecStart, SecCount As Integer, GroupCount As Integer, LastSec As Integer)
Dim i As Long
Dim j As Long
Dim k As Long
Dim M As Integer
Dim P As Integer
Dim interval As Integer
Dim rCount As Integer
Dim TruStart As Long
k = VarNum
If Numbers(k - 1) = "154" Then
    DoEvents
End If
For i = SecStart To (SecStart + SecCount - 1)
    TruStart = SecStart + interval
    M = i + 1
    For j = 1 To GroupCount
        If SecStart > 1 Then
            P = adv(P, 1, (GroupCount), 0)
        End If
        If M > (SecCount / LastSec) Then
            M = TruStart
        End If
        Numbers(k) = Numbers(i) & Right(Numbers(M + P), 1)
        k = k + 1
        M = M + 1
        nCount = nCount + 1
    Next j
    rCount = rCount + 1
    If rCount = (GroupCount + 1) Then
        interval = interval + (GroupCount + 1)
        rCount = 0
    End If
    If SecStart > 1 Then P = rCount
Next i
If GroupCount - 1 > 0 Then
    PlaceNum k, SecStart + SecCount, GroupCount * SecCount, GroupCount - 1, SecCount
End If
End Sub







Sub UnLoadDict()
Erase Dict()
DLoaded = False
End Sub


