Attribute VB_Name = "Functions"
Public start As Boolean
Public score As Integer
Public previous As Integer, current As Integer
Public back(65) As Integer
Public status(9, 9) As Boolean
Public row As Integer, col As Integer
Public valid(9) As Boolean
Public hintrow(9) As Integer, hintcol(9) As Integer
Public i As Integer, j As Integer
Public tottime As Integer
Public validmove As Boolean

Public Function showmoves(position As Integer) As Boolean
showmoves = False
Call index_to_pos(position)   ' calculate current position from index
' calculate 8 possible moves

hintrow(1) = row - 2
hintcol(1) = col + 1

hintrow(2) = row - 2
hintcol(2) = col - 1

hintrow(3) = row + 2
hintcol(3) = col + 1

hintrow(4) = row + 2
hintcol(4) = col - 1

hintrow(5) = row + 1
hintcol(5) = col - 2

hintrow(6) = row - 1
hintcol(6) = col - 2

hintrow(7) = row - 1
hintcol(7) = col + 2

hintrow(8) = row + 1
hintcol(8) = col + 2

' check which moves from the above calculated are valid moves
For i = 1 To 8
    valid(i) = True
    If hintrow(i) > 8 Or hintrow(i) < 1 Or hintcol(i) > 8 Or hintcol(i) < 1 Then
        valid(i) = False
    Else
        If status(hintrow(i), hintcol(i)) = False Then valid(i) = False
    End If
Next i

For i = 1 To 8
    If valid(i) = True Then
        showmoves = True
        Exit For
    End If
Next i

End Function

Public Function index_to_pos(pic As Integer)
row = (pic \ 8) + 1
col = (pic - (pic \ 8) * 8) + 1
End Function

Public Function pos_to_index(trow As Integer, tcol As Integer) As Integer
pos_to_index = ((trow - 1) * 8) + tcol - 1
End Function

Public Function updboard()

If score < 64 Then mainfrm.cmdback.Enabled = True   ' enable BACK button

Call index_to_pos(previous)
status(row, col) = False       ' mark previous square as covered

score = score - 1              'update score
mainfrm.lblscore.Caption = score & " squares left"

back(64 - score) = current   ' store move

For i = 1 To 8
    For j = 1 To 8
        If status(i, j) = False Then
            mainfrm.board(pos_to_index(i, j)).Picture = mainfrm.fillpict.Picture
        Else
            mainfrm.board(pos_to_index(i, j)).Picture = mainfrm.emptypict.Picture
        End If
    Next j
Next i
mainfrm.board(current).Picture = mainfrm.mainpict.Picture
End Function
Public Function again()
start = False
score = 64
tottime = 0
current = 100
previous = 100
For i = 1 To 8
    For j = 1 To 8
        status(i, j) = True
    Next j
Next i
For i = 0 To 63
    mainfrm.board(i).Picture = mainfrm.emptypict.Picture
Next i
mainfrm.cmdback.Enabled = False
mainfrm.cmdhint.Enabled = False
mainfrm.Timer1.Enabled = False
mainfrm.lblscore.Caption = "64 squares left"
mainfrm.lbltime.Caption = "0 Seconds"
mainfrm.lbldir.Caption = "Please select a square on the board to start."
mainfrm.cmdpause.Enabled = False
mainfrm.Caption = "KNIGHT'S TOUR"
End Function
