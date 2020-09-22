Attribute VB_Name = "modWR"
Option Explicit
Public MyWords() As String
Public Lookup() As String
Public XGrid() As String
Public x1 As Integer
Public y1 As Integer
Public x2 As Integer
Public y2 As Integer
Public AllStop As Boolean
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public ff As Integer
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public MAX As Integer
Public MIN As Integer

Sub AddNextLetter(a As Integer, b As Integer, StartDir As Integer, RootWord As String)
If AllStop Then Exit Sub
Dim c As Integer
Dim d As Integer
Dim k As Integer
Dim TempWord As String
Dim PathFound As Boolean
Dim MyChar As String
PathFound = False
TempWord = RootWord
For k = StartDir To 8
    c = a
    d = b
    MyChar = GetChar(k, a, b, TempWord)
    If MyChar <> "" Then
        PathFound = True
        AddNextLetter c, d, k + 1, RootWord
        TempWord = TempWord & MyChar
        StartDir = 1
        Exit For
    End If
Next k
k = Len(TempWord)
If k >= MIN And k <= MAX Then
    AddWord TempWord
    CheckIt TempWord, k
End If
If k < MAX And PathFound Then
    AddNextLetter a, b, 1, TempWord
End If
End Sub
Sub CheckIt(MyWord As String, ListNum As Integer)
Dim TMP As String
TMP = RealWord(MyWord)
If IsWord(TMP) Then
    If Not InListbox(TMP, frmSolver.lstWords(ListNum)) Then
        frmSolver.lstWords(ListNum).AddItem TMP
        frmSolver.lstWords(ListNum).Selected(frmSolver.lstWords(ListNum).ListCount - 1) = True
    End If
End If
End Sub
Function InListbox(sStringToFind As String, lstListBox As ListBox) As Boolean
    InListbox = SendMessageByString(lstListBox.hwnd, LB_FINDSTRING, -1, sStringToFind) >= 0
End Function
Function GetChar(d As Integer, ByRef x As Integer, ByRef y As Integer, Word As String) As String
On Local Error GoTo eTrap
Dim Char As String
Dim i As Integer
Dim a As Integer
Dim b As Integer
Select Case d
    Case 1 'upleft
        a = x - 1
        b = y - 1
    Case 2 'up
        a = x
        b = y - 1
    Case 3 'upright
        a = x + 1
        b = y - 1
    Case 4 'left
        a = x - 1
        b = y
    Case 5 'right
        a = x + 1
        b = y
    Case 6 'downleft
        a = x - 1
        b = y + 1
    Case 7 'down
        a = x
        b = y + 1
    Case 8 'downright
        a = x + 1
        b = y + 1
End Select
Char = XGrid(a, b)
If Char = "" Then Exit Function
For i = 1 To Len(Word)
    If Mid(Word, i, 1) = Char Then Exit Function
Next i
x = a
y = b
GetChar = Char
Exit Function
eTrap:
GetChar = ""
End Function
Function RealWord(MyWord As String) As String
Dim i As Integer
Dim a As Integer
Dim b As Integer
For i = 1 To Len(MyWord)
    RealWord = RealWord & Lookup(Asc(Mid(MyWord, i, 1)))
Next i
End Function


Function Reverse(MyWord As String) As String
Dim i As Integer
For i = 1 To Len(MyWord)
    Reverse = Mid(MyWord, i, 1) & Reverse
Next i
End Function

Public Sub SearchWord()
On Local Error GoTo eTrap
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim TempWord As String
Dim PathFound As Boolean
Dim MyChar As String
For i = 0 To x2
    For j = 0 To y2
        a = i
        b = j
        TempWord = XGrid(a, b)
        AddNextLetter a, b, 1, TempWord
        If AllStop Then Exit Sub
    Next j
Next i
'j = UBound(MyWords)
'For i = 0 To j
'    frmSolver.Caption = "Length = (" & K & ") Checking " & i & " / " & j
'    CheckIt MyWords(i), K
'    DoEvents
'    If AllStop Then Exit Sub
'Next i
Erase MyWords()
Exit Sub
eTrap:
    j = -1
    Resume Next
End Sub


Sub AddWord(MyWord As String)
On Local Error GoTo eTrap
Dim i As Long
i = UBound(MyWords) + 1
ReDim Preserve MyWords(i)
MyWords(i) = MyWord
frmSolver.Caption = "Adding Word: " & i
DoEvents
Exit Sub
eTrap:
    i = 0
    Resume Next
End Sub






