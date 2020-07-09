VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKnightTour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
' class: Implements the Warndorf's rule for solving the Knights tour problem
'***************************************************************************


Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)            'For 32 Bit Systems
#End If

Private Type udtLocation
    RowValue    As Long
    ColumnValue As Integer
End Type

Private mudtPosition(1 To 8)    As udtLocation
Private mrngArrVisited()        As Range
Private mrngBoard               As Range
Private mintVisitedCnt          As Integer
Private mintSequence            As Integer
Private mintPieceAscCode        As Integer
Private mlngDelay               As Long
Private mstrPieceFontName       As String

Private Const CINT_VISITED_COLORINDEX   As Integer = 1
Private Const CINT_CURRENT_COLORINDEX   As Integer = 45
Private Const CINT_CURRENT_SYMBOL_SIZE  As Integer = 28
Private Const CINT_VISITED_SYMBOL_SIZE  As Integer = 8
Private Const CINT_DEFAULT_ASCII_CHAR   As Integer = 88
Private Const CLNG_DEFAULT_DELAY        As Long = 250
Private Const CSTR_DEFAULT_FONT         As String = "Arial"

Private Sub Class_Initialize()
    mlngDelay = CLNG_DEFAULT_DELAY
    mudtPosition(1).RowValue = -2: mudtPosition(1).ColumnValue = 1
    mudtPosition(2).RowValue = 1: mudtPosition(2).ColumnValue = 2
    mudtPosition(3).RowValue = 2: mudtPosition(3).ColumnValue = 1
    mudtPosition(4).RowValue = -1: mudtPosition(4).ColumnValue = 2
    mudtPosition(5).RowValue = -2: mudtPosition(5).ColumnValue = -1
    mudtPosition(6).RowValue = -1: mudtPosition(6).ColumnValue = -2
    mudtPosition(7).RowValue = 1: mudtPosition(7).ColumnValue = -2
    mudtPosition(8).RowValue = 2: mudtPosition(8).ColumnValue = -1
    mintSequence = 1
    ReDim mrngArrVisited(0 To 0)
    mintVisitedCnt = 0
    mstrPieceFontName = CSTR_DEFAULT_FONT
    mintPieceAscCode = CINT_DEFAULT_ASCII_CHAR
End Sub

Public Property Let PieceFontName(ByVal Value As String)
    mstrPieceFontName = Value
End Property

Public Property Let PieceAscCode(ByVal Value As Integer)
    mintPieceAscCode = Value
End Property

Public Property Let BoardArea(ByVal Value As Range)
    Set mrngBoard = Value
    mrngBoard.ClearContents
End Property

Private Function IsValidPosition(ByVal rngCell As Range) As Boolean
    If Intersect(rngCell, mrngBoard) Is Nothing Then
        IsValidPosition = False
    Else
        IsValidPosition = True
    End If
End Function

Public Function GetRandomSquareFromBoard() As Range
    Dim intRndCell          As Integer
    Dim intBoardCellsCnt    As Integer

    If Not mrngBoard Is Nothing Then
        intBoardCellsCnt = mrngBoard.Cells.Count
        intRndCell = GetRandomNumber(intBoardCellsCnt, 1)
        Set GetRandomSquareFromBoard = mrngBoard.Cells(intRndCell)
    End If
End Function

Public Sub RemovePiece(ByVal rngSquare As Range)
    With rngSquare
        .Font.Size = CINT_VISITED_SYMBOL_SIZE
        .Font.Bold = True
        .Font.ColorIndex = CINT_VISITED_COLORINDEX
        .Font.Name = CSTR_DEFAULT_FONT
    End With
End Sub

Public Sub DisplayPiece(ByVal rngSquare As Range)
    With rngSquare
        .Font.Size = CINT_CURRENT_SYMBOL_SIZE
        .Font.Bold = True
        .Font.ColorIndex = CINT_CURRENT_COLORINDEX
        .Value = Chr(mintPieceAscCode)
        .Font.Name = mstrPieceFontName
    End With
End Sub

Public Sub MovePiece(ByVal rngFrom As Range, ByVal rngTo As Range)
    Application.ScreenUpdating = False
    Sleep mlngDelay
    RemovePiece rngFrom
    rngFrom.Value = mintSequence
    mintSequence = mintSequence + 1
    DisplayPiece rngTo
    Application.ScreenUpdating = True
End Sub

Public Function GetNextMove(ByVal rngCell As Range) As Range
    Dim intCnt                      As Integer
    Dim intMaxMoves                 As Integer
    Dim intMoves                    As Integer
    Dim intArrCnt                   As Integer
    Dim intRnd                      As Integer
    Dim rngNewLocation              As Range
    Dim arrListOfSquaresToMoveTo()  As Range
    
    intArrCnt = 0
    intMoves = 0
    intMaxMoves = UBound(mudtPosition)
    
    ReDim Preserve mrngArrVisited(0 To mintVisitedCnt)
    Set mrngArrVisited(mintVisitedCnt) = rngCell
    mintVisitedCnt = mintVisitedCnt + 1

    For intCnt = LBound(mudtPosition) To UBound(mudtPosition)
        DoEvents
        Set rngNewLocation = rngCell.Offset(mudtPosition(intCnt).RowValue, mudtPosition(intCnt).ColumnValue)
        If IsValidPosition(rngNewLocation) Then
            If Not IsVisitedLocation(rngNewLocation) Then
                intMoves = CountPossibleMovesFromLocation(rngNewLocation)
                Select Case intMoves
                Case Is < intMaxMoves
                     intMaxMoves = intMoves
                     intArrCnt = 0
                     ReDim arrListOfSquaresToMoveTo(0 To intArrCnt)
                     Set arrListOfSquaresToMoveTo(intArrCnt) = rngNewLocation
                     intArrCnt = intArrCnt + 1
                Case intMaxMoves
                     ReDim Preserve arrListOfSquaresToMoveTo(0 To intArrCnt)
                     Set arrListOfSquaresToMoveTo(intArrCnt) = rngNewLocation
                     intArrCnt = intArrCnt + 1
                End Select
            End If
        End If
    Next

    If intArrCnt > 0 Then
        intRnd = GetRandomNumber(UBound(arrListOfSquaresToMoveTo), LBound(arrListOfSquaresToMoveTo))
        Set GetNextMove = arrListOfSquaresToMoveTo(intRnd)
    End If
    
    Set rngNewLocation = Nothing
End Function

Private Function CountPossibleMovesFromLocation(ByVal rngCell As Range) As Integer
    Dim intPos          As Integer
    Dim intCnt          As Integer
    Dim rngNewLocation  As Range
    
    intCnt = 0
    For intPos = LBound(mudtPosition) To UBound(mudtPosition)
        DoEvents
        Set rngNewLocation = rngCell.Offset(mudtPosition(intPos).RowValue, mudtPosition(intPos).ColumnValue)
        If IsValidPosition(rngNewLocation) Then
            If Not IsVisitedLocation(rngNewLocation) Then
                intCnt = intCnt + 1
            End If
        End If
    Next
    Set rngNewLocation = Nothing
    CountPossibleMovesFromLocation = intCnt
End Function

Private Function IsVisitedLocation(ByVal rngCell As Range) As Boolean
    Dim intCnt As Integer

    IsVisitedLocation = False
    If mintVisitedCnt = 0 Then
        Exit Function
    End If
    
    For intCnt = LBound(mrngArrVisited) To UBound(mrngArrVisited)
        DoEvents
        If mrngArrVisited(intCnt).Address = rngCell.Address Then
            IsVisitedLocation = True
            Exit Function
        End If
    Next
End Function

Private Function GetRandomNumber(ByVal lngMaxValue As Long, Optional ByVal lngMinValue As Long = 0)
    Randomize
    GetRandomNumber = Int((lngMaxValue - lngMinValue + 1) * Rnd) + lngMinValue
End Function


