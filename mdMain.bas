Attribute VB_Name = "mdMain"
'*************************************
' purpose: runs the entire application
'*************************************

Option Explicit

Public Sub StartKnightsTour()

On Error GoTo ERR_HANDLER:

    Dim rngSquare       As Range
    Dim rngSquareNext   As Range
    Dim oKnight         As New clsKnightTour
    
    oKnight.BoardArea = Range("Board")
    Set rngSquare = oKnight.GetRandomSquareFromBoard
    oKnight.PieceAscCode = 140
    oKnight.PieceFontName = "Chess Alpha"
    oKnight.DisplayPiece rngSquare
    With oKnight
        Do
            Set rngSquareNext = .GetNextMove(rngSquare)
            If rngSquareNext Is Nothing Then
                Exit Do 'no more possible moves
            Else
                .MovePiece rngSquare, rngSquareNext
            End If
            Set rngSquare = rngSquareNext
        Loop
    End With
EXIT_HERE:
    Set rngSquare = Nothing
    Set rngSquareNext = Nothing
    Set oKnight = Nothing
    Exit Sub
ERR_HANDLER:
    Debug.Print Err.Description
    GoTo EXIT_HERE
End Sub




