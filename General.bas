Attribute VB_Name = "General"
Option Explicit
Dim ForcedScore As Long

Public Function CheckSquare(NewPieces() As PieceAttr, Optional X As Long, Optional Y As Long, Optional Index As Long) As SelectedSquare
Dim PieceN As Long

If X = 0 Then XYConvert Index, X, Y Else Index = IConvert(X, Y)

If Index < 0 Or Index > 99 Or X < 1 Or X > 10 Or Y < 1 Or Y > 10 Then
  CheckSquare.X = 0: CheckSquare.Y = 0: CheckSquare.Index = OFF_BOARD: Exit Function
End If

With CheckSquare
  .Index = Index
  .IsSquare = True
  .X = X
  .Y = Y
End With

For PieceN = 1 To 40
  If NewPieces(PieceN).X = X And NewPieces(PieceN).Y = Y And NewPieces(PieceN).Index <> OFF_BOARD Then
    With CheckSquare
      .IsPiece = True
      If PieceN <= 20 Then .Player = 1 Else .Player = 2
      .Piece = PieceN
      .Double = NewPieces(PieceN).Double
    End With
    Exit Function
  End If
Next

End Function

Public Function StateCheckSquare(NewPieces() As StatePieceAttr, Index As Long) As StateSelectedSquare
Dim PieceN As Long

If Index < 0 Or Index > 99 Then
  StateCheckSquare.Index = OFF_BOARD
  Exit Function
End If

StateCheckSquare.Index = Index

For PieceN = 1 To 40
  If NewPieces(PieceN).Index = Index Then
    With StateCheckSquare
      If PieceN < 21 Then .Player = 1 Else .Player = 2
      .Piece = PieceN
      .Double = NewPieces(PieceN).Double
    End With
    Exit Function
  End If
Next

StateCheckSquare.Index = NOT_PIECE

End Function

Public Sub Deletepiece(Piece As Long, ByRef Pieces() As PieceAttr, NotReal As Boolean)
  If NotReal = False Then frmMain.Shape1(IndexTranslation(Pieces(Piece).Index)).Picture = Nothing
  Pieces(Piece).X = 0: Pieces(Piece).Y = 0: Pieces(Piece).Index = OFF_BOARD
End Sub

Public Function MovePiece(From As Long, Too As Long, ArrayNum As Long, MultiMove As Boolean, Optional Highlight As Boolean) As Boolean
'On Error GoTo Error
Dim FromSquare As SelectedSquare, ToSquare As SelectedSquare, TempSquare As SelectedSquare
Dim Lng1 As Long, Moves() As MMoves, HasTaken As Boolean, Direction As Long, High As Long, Low As Long
Dim HasToTake As Long, Lng2 As Long, TempPieces(1 To 40) As PieceAttr

FromSquare = CheckSquare(Currentpieces, , , Val(From)) 'Record info about original position
ToSquare = CheckSquare(Currentpieces, , , Val(Too))    'Record info about new position

If (ToSquare.X + ToSquare.Y) Mod 2 = 0 Then Exit Function 'Checks if going to a black square

If Turn = 1 Then High = 20: Low = 1 Else High = 40: Low = 21

'-----------Copy Currentpieces into temppieces (for theoretical stuff for force move rule)--

For Lng1 = 1 To 40
  TempPieces(Lng1) = Currentpieces(Lng1)
Next

If CheatSwitch = 1 Then GoTo SkipChecks

'--------Gets the direction in which the piece is traveling in--------

If ArrayNum = 0 Or SMatrix(ArrayNum).SubParentId = 0 Then
  If (ToSquare.Index - FromSquare.Index) < 0 Then
    If (ToSquare.Index - FromSquare.Index) Mod 9 = 0 Then
      Direction = 1
    ElseIf (ToSquare.Index - FromSquare.Index) Mod 11 = 0 Then
      Direction = 4
    End If
  Else
    If (ToSquare.Index - FromSquare.Index) Mod 9 = 0 Then
      Direction = 3
    ElseIf (ToSquare.Index - FromSquare.Index) Mod 11 = 0 Then
      Direction = 2
    End If
  End If
End If

If Direction = 0 And ArrayNum = 0 Then Exit Function

'--------------used for computer player to take more than 1 piece at a time (ignore it it's very complicated)---

If SMatrix(ArrayNum).SubParentId <> 0 Then
  ReDim Moves(0 To 0)
  Moves(0).ParentID = ArrayNum
  For Lng1 = 21 To 40
    If MoveMatrix(SMatrix(Moves(UBound(Moves)).ParentID).ArrayNum).Pieces(Lng1).Index <> MoveMatrix(SMatrix(SMatrix(Moves(UBound(Moves)).ParentID).SubParentId).ArrayNum).Pieces(Lng1).Index Then Exit For
  Next
  Moves(UBound(Moves)).From = MoveMatrix(SMatrix(SMatrix(Moves(UBound(Moves)).ParentID).SubParentId).ArrayNum).Pieces(Lng1).Index
  Moves(UBound(Moves)).Too = MoveMatrix(SMatrix(Moves(UBound(Moves)).ParentID).ArrayNum).Pieces(Lng1).Index
  
  Do While SMatrix(Moves(UBound(Moves)).ParentID).SubParentId <> 0
    ReDim Preserve Moves(0 To UBound(Moves) + 1)
    Moves(UBound(Moves)).ParentID = SMatrix(Moves(UBound(Moves) - 1).ParentID).SubParentId
    For Lng1 = 21 To 40
      If MoveMatrix(SMatrix(Moves(UBound(Moves)).ParentID).ArrayNum).Pieces(Lng1).Index <> MoveMatrix(SMatrix(SMatrix(Moves(UBound(Moves)).ParentID).SubParentId).ArrayNum).Pieces(Lng1).Index Then Exit For
    Next
    Moves(UBound(Moves)).From = MoveMatrix(SMatrix(SMatrix(Moves(UBound(Moves)).ParentID).SubParentId).ArrayNum).Pieces(Lng1).Index
    Moves(UBound(Moves)).Too = MoveMatrix(SMatrix(Moves(UBound(Moves)).ParentID).ArrayNum).Pieces(Lng1).Index
  Loop
  
  For Lng1 = UBound(Moves) To 0 Step -1
    Call MovePiece(Moves(Lng1).From, Moves(Lng1).Too, 0, True)
    Sleep MoveSpeed
  Next
  
  For Lng1 = 1 To 40
    TempPieces(Lng1) = Currentpieces(Lng1)
  Next
  
  GoTo SkipChecks

End If

'----------Checks if forcing taking a piece is neccessary-----------

If ForceMove = True Then
  ForcedScore = 0
  For Lng1 = Low To High
    If CanTake(Lng1, TempPieces) Then
      ReDim SMatrix(0 To 0)
      ReDim MoveMatrix(0 To 0)
      HasToTake = True
      CopyArray TempPieces, MoveMatrix(0).Pieces
      GenerateStates 0, True, Turn
      For Lng2 = 1 To UpperMove
        If (StateCountpieces(MoveMatrix(Lng2).Pieces) < ForcedScore Or ForcedScore = 0) And SMatrix(Lng2).SubParentId = 0 Then ForcedScore = StateCountpieces(MoveMatrix(Lng2).Pieces)
      Next Lng2
      Exit For
    End If
  Next Lng1
End If

'----------Checks for valid move (also not recomended)------------

If FromSquare.Double = True Then
  If ToSquare.IsPiece = True Then
    If ToSquare.Player = Turn Then Exit Function
    Select Case Direction
      Case 1
        If ToSquare.X = 10 Or ToSquare.Y = 1 Then Exit Function
      Case 2
        If ToSquare.X = 10 Or ToSquare.Y = 10 Then Exit Function
      Case 3
        If ToSquare.X = 1 Or ToSquare.Y = 10 Then Exit Function
      Case 4
        If ToSquare.X = 1 Or ToSquare.Y = 1 Then Exit Function
    End Select
    TempSquare = CheckSquare(TempPieces, , , ToSquare.Index + IndexMoves(Direction, 1))
    If TempSquare.IsPiece Then Exit Function
    For Lng1 = 1 To (ToSquare.X - FromSquare.X - 1)
      If CheckSquare(TempPieces, , , FromSquare.Index + IndexMoves(Direction, Lng1)).IsPiece Then Exit Function
    Next
    Call Deletepiece(ToSquare.Piece, TempPieces, Highlight): HasTaken = True
    Too = TempSquare.Index
  ElseIf ToSquare.IsPiece = False And CheckSquare(TempPieces, , , ToSquare.Index - IndexMoves(Direction, 1)).IsPiece = True And Abs(FromSquare.Index - ToSquare.Index) > 11 Then
    TempSquare = CheckSquare(TempPieces, , , ToSquare.Index - IndexMoves(Direction, 1))
    If TempSquare.Player = Turn Then Exit Function
    For Lng1 = 1 To (ToSquare.X - FromSquare.X - 2)
      If CheckSquare(TempPieces, , , FromSquare.Index + IndexMoves(Direction, Lng1)).IsPiece Then Exit Function
    Next
    Call Deletepiece(TempSquare.Piece, TempPieces, Highlight): HasTaken = True
  Else
    For Lng1 = 1 To (ToSquare.X - FromSquare.X)
      If CheckSquare(TempPieces, , , FromSquare.Index + IndexMoves(Direction, Lng1)).IsPiece Then Exit Function
    Next
    If HasToTake = True Then Exit Function
  End If
  
Else

  If ToSquare.X > FromSquare.X + 2 Or ToSquare.X < FromSquare.X - 2 Or ToSquare.Y > FromSquare.Y + 2 Or ToSquare.Y < FromSquare.Y - 2 Then Exit Function
  
  If ToSquare.IsPiece And (ToSquare.X = FromSquare.X + 1 Or ToSquare.X = FromSquare.X - 1 Or ToSquare.Y = FromSquare.Y + 1 Or ToSquare.Y = FromSquare.Y - 1) Then
    If ToSquare.Player = Turn Then Exit Function
    TempSquare = CheckSquare(TempPieces, , , ToSquare.Index + IndexMoves(Direction, 1))
    Select Case Direction
      Case 1
        If ToSquare.X = 10 Or ToSquare.Y = 1 Then Exit Function
      Case 2
        If ToSquare.X = 10 Or ToSquare.Y = 10 Then Exit Function
      Case 3
        If ToSquare.X = 1 Or ToSquare.Y = 10 Then Exit Function
      Case 4
        If ToSquare.X = 1 Or ToSquare.Y = 1 Then Exit Function
    End Select
    If TempSquare.IsPiece Then Exit Function
    Call Deletepiece(ToSquare.Piece, TempPieces, Highlight): HasTaken = True
    Too = TempSquare.Index
  ElseIf ToSquare.X = FromSquare.X + 2 Or ToSquare.X = FromSquare.X - 2 Or ToSquare.Y = FromSquare.Y + 2 Or ToSquare.Y = FromSquare.Y - 2 Then
    If ToSquare.IsPiece Then Exit Function
    TempSquare = CheckSquare(TempPieces, , , ToSquare.Index - IndexMoves(Direction, 1))
    If TempSquare.IsPiece = False Or TempSquare.Player = Turn Then Exit Function
    Call Deletepiece(TempSquare.Piece, TempPieces, Highlight): HasTaken = True
  Else
    If HasToTake = True Then Exit Function
    Select Case FromSquare.Player
    Case 1
      If ToSquare.Y > FromSquare.Y And FromSquare.Double = False Then Exit Function
    Case 2
      If ToSquare.Y < FromSquare.Y And FromSquare.Double = False Then Exit Function
    End Select
  End If

End If

If HasTaken = False And P1MultiMode = True Then Exit Function

SkipChecks:

'---Starts to actually move the piece---

For Lng1 = 1 To 40 'For the back button which doesn't work very well!
  Lastpieces(Lng1) = TempPieces(Lng1)
Next

If CheatSwitch = 1 And Highlight = False Then
  TempSquare = CheckSquare(TempPieces, , , Val(Too))
  If TempSquare.IsPiece Then Deletepiece TempSquare.Piece, TempPieces, False
End If

'------Moves crap and changes pictures (and stuff like that)-----

TempPieces(FromSquare.Piece).Index = Too
XYConvert Too, TempPieces(FromSquare.Piece).X, TempPieces(FromSquare.Piece).Y

If HasTaken And (frmMain.Option2 Or Turn = 1) And CanTake(CheckSquare(TempPieces, , , IndexTranslation(Val(Too))).Piece, TempPieces, True) And AutoDebug = False And Highlight = False Then
  frmMain.Shape1(Too).Picture = frmMain.ImageList1.ListImages(5).Picture
  P1MultiMode = True
Else
  P1MultiMode = False
End If

Select Case FromSquare.Player
  Case 1
    If TempPieces(FromSquare.Piece).Y = 1 And MultiMove = False And P1MultiMode = False Then TempPieces(FromSquare.Piece).Double = True
  Case 2
    If TempPieces(FromSquare.Piece).Y = 10 And MultiMove = False And P1MultiMode = False Then TempPieces(FromSquare.Piece).Double = True
End Select

If Countpieces(TempPieces) = ForcedScore Or HasToTake = False Or P1MultiMode = True Then
  If Highlight = True Then MovePiece = True: Exit Function
  For Lng1 = 1 To 40
    Currentpieces(Lng1) = TempPieces(Lng1)
  Next
End If

If Highlight Then Exit Function

frmMain.Shape1(IndexTranslation(Val(From))).Picture = Nothing
Too = IndexTranslation(Val(Too))

If HasTaken And (frmMain.Option2 Or Turn = 1) And CanTake(CheckSquare(TempPieces, , , IndexTranslation(Val(Too))).Piece, TempPieces, True) And AutoDebug = False Then Exit Function

If MultiMove Then frmMain.Shape1(Too).Picture = frmMain.ImageList1.ListImages(5).Picture: Exit Function

If FromSquare.Player = 1 Then
  If TempPieces(FromSquare.Piece).Double Then
    frmMain.Shape1(Too).Picture = frmMain.ImageList1.ListImages(2).Picture
  Else
    frmMain.Shape1(Too).Picture = frmMain.ImageList1.ListImages(1).Picture
  End If
Else
  If TempPieces(FromSquare.Piece).Double Then
    frmMain.Shape1(Too).Picture = frmMain.ImageList1.ListImages(4).Picture
  Else
    frmMain.Shape1(Too).Picture = frmMain.ImageList1.ListImages(3).Picture
  End If
End If

If CheatSwitch <> 1 Then

  If Turn = 1 Then
    Turn = 2
    If CheckWin(Currentpieces) = True Then
      RefreshDisplay
      MovePiece = True
      Exit Function
    ElseIf CanMove(Currentpieces) = False Then
      If AutoDebug = False Then MsgBox Names(1) & " wins!", vbExclamation
      ResetGame
    Else
      RefreshDisplay
    End If
    DoEvents
    If frmMain.Option1 And CheatSwitch = 0 Then Call AIMove(Turn)
    If frmMain.Option2 And frmMain.CheckAutoSwitch Then Sleep 500: Reversed = True: RefreshBoard Currentpieces  ' Switch board
  Else
    Turn = 1
    If CheckWin(Currentpieces) = True Then
      RefreshDisplay
      MovePiece = True
      Exit Function
    ElseIf CanMove(Currentpieces) = False Then
      If AutoDebug = False Then MsgBox Names(2) & " wins!", vbExclamation
      ResetGame
    Else
      RefreshDisplay
    End If
    DoEvents
    If AutoDebug = True Then Call AIMove(Turn)
    If frmMain.Option2 And frmMain.CheckAutoSwitch Then Sleep 500: Reversed = False: RefreshBoard Currentpieces  ' Switch board
  End If

  VTurns = VTurns + 1
  frmMain.lblTurns = VTurns
End If

MovePiece = True

Exit Function
Error:
  RefreshBoard TempPieces
  MsgBox "There has been an internal error in the function 'Movepiece' in the Movement module" & vbCrLf & "Cause - " & Err.Description, vbCritical, "ERR_" & Err.Number
End Function

Public Function CanMove(Pieces() As PieceAttr) As Boolean
Dim Direction As Long, PieceN As Long, High As Long, Low As Long, TempSquare As SelectedSquare
Dim X As Long, Y As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, NewIndex As Long, NewIndex2 As Long

If Turn = 1 Then Low = 1: High = 20 Else Low = 20: High = 40

For PieceN = Low To High
  If Pieces(PieceN).Index = OFF_BOARD Then GoTo Next_Piece 'if piece is dead
  XYConvert Pieces(PieceN).Index, X, Y
  For Direction = 1 To 4
    
    NewIndex = Pieces(PieceN).Index + IndexMoves(Direction, 1)
    X1 = X + XYMoves(Direction, 1).X
    Y1 = Y + XYMoves(Direction, 1).Y
    If X1 < 1 Or X1 > 10 Or Y1 < 1 Or Y1 > 10 Then GoTo Next_Direction
    
    TempSquare = CheckSquare(Pieces, , , Val(NewIndex))
    If TempSquare.IsSquare = False Then GoTo Next_Direction
    If TempSquare.IsPiece Then
      '---------Place over piece to take--------
      NewIndex2 = Pieces(PieceN).Index + IndexMoves(Direction + 4, 1)
      X2 = X + XYMoves(Direction + 4, 1).X
      Y2 = Y + XYMoves(Direction + 4, 1).Y
      If X2 < 1 Or X2 > 10 Or Y2 < 1 Or Y2 > 10 Then GoTo Next_Direction
      
      If CheckSquare(Pieces, , , Val(NewIndex2)).IsPiece = False _
      And TempSquare.IsPiece = True And TempSquare.Player <> Turn Then CanMove = True: Exit Function

    Else
      If Turn = 1 Then 'if it can move in that direction
        If Pieces(PieceN).Double = False And (Direction = 2 Or Direction = 3) Then GoTo Next_Direction
      Else
        If Pieces(PieceN).Double = False And (Direction = 1 Or Direction = 4) Then GoTo Next_Direction
      End If
      CanMove = True: Exit Function
    End If
  
Next_Direction:
  Next Direction
Next_Piece:
Next PieceN

End Function

Public Function CanTake(PieceN As Long, Pieces() As PieceAttr, Optional ExcludeDouble As Boolean) As Boolean
Dim Direction As Long, Player As Long, NewIndex As Long, X As Long, Y As Long
Dim TempSquare As SelectedSquare, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
Dim NewIndex2 As Long, MoveLength As Long, MaxMoveLength As Long

Player = Pieces(PieceN).Player

If Pieces(PieceN).Index = OFF_BOARD Then Exit Function
  XYConvert Pieces(PieceN).Index, X, Y
  If Pieces(PieceN).Double = True And ExcludeDouble = False Then MaxMoveLength = 10 Else MaxMoveLength = 1
  For Direction = 1 To 4
    For MoveLength = 1 To MaxMoveLength
      NewIndex = Pieces(PieceN).Index + IndexMoves(Direction, MoveLength)
      X1 = X + XYMoves(Direction, MoveLength).X
      Y1 = Y + XYMoves(Direction, MoveLength).Y
      If X1 < 1 Or X1 > 10 Or Y1 < 1 Or Y1 > 10 Then GoTo Next_Direction
      
      TempSquare = CheckSquare(Pieces, , , Val(NewIndex))
      If TempSquare.IsSquare = False Then GoTo Next_Direction
      If TempSquare.IsPiece Then
      
        NewIndex2 = Pieces(PieceN).Index + IndexMoves(Direction + 4, MoveLength)
        X2 = X + XYMoves(Direction + 4, MoveLength).X
        Y2 = Y + XYMoves(Direction + 4, MoveLength).Y
        If X2 < 1 Or X2 > 10 Or Y2 < 1 Or Y2 > 10 Then GoTo Next_Direction
        
        If CheckSquare(Pieces, , , Val(NewIndex2)).IsPiece = False _
        And TempSquare.IsPiece = True And TempSquare.Player <> Player Then CanTake = True: Exit Function
        
        GoTo Next_Direction
        
      End If
    Next MoveLength
Next_Direction:
  Next Direction

End Function

Public Function IndexTranslation(Index As Long) As Long
  If Reversed Then IndexTranslation = 99 - Index Else IndexTranslation = Index
End Function

Public Sub ResetGame()
Dim Lng1 As Long, Lng2 As Long, X As Long, Y As Long

Turn = 1

'AutoDebug = False

GameStarted = False

With frmMain
  .lblMMatrixSize = "0"
  .lblP1Time = "0 Min 0 Sec"
  .lblP2Time = "0 Min 0 Sec"
  .lblTotalTime = "0 Min 0 Sec"
  .lblPlyDepth = "0"
  .lblTurns = "0"
End With

TotalTime.Minutes = 0
TotalTime.Seconds = 0
VP2Time.Minutes = 0
VP2Time.Seconds = 0
VP1Time.Minutes = 0
VP1Time.Seconds = 0
VTurns = 0

For Lng1 = 60 To 99
  XYConvert Lng1, X, Y
  If (X + Y) Mod 2 <> 0 Then
    Lng2 = Lng2 + 1
    Currentpieces(Lng2).X = X
    Currentpieces(Lng2).Y = Y
    Currentpieces(Lng2).Index = Lng1
  End If
Next

For Lng1 = 0 To 39
  XYConvert Lng1, X, Y
  If (X + Y) Mod 2 <> 0 Then
    Lng2 = Lng2 + 1
    Currentpieces(Lng2).X = X
    Currentpieces(Lng2).Y = Y
    Currentpieces(Lng2).Index = Lng1
  End If
Next

For Lng1 = 1 To 40
  Currentpieces(Lng1).Index = IConvert(Currentpieces(Lng1).X, Currentpieces(Lng1).Y)
  If Lng1 <= 20 Then Currentpieces(Lng1).Player = 1 Else Currentpieces(Lng1).Player = 2
  Currentpieces(Lng1).Double = False
Next

Call RefreshDisplay
Call RefreshBoard(Currentpieces)

GetSettings

Score(1) = 20
Score(2) = 20

If AutoDebug = True Then Call AIMove(Turn)

End Sub

Public Sub RefreshDisplay()
Dim PieceN As Long

  Score(1) = 0: Score(2) = 0

  If Turn = 1 Then
    frmMain.lblTurn = Names(1)
  Else
    frmMain.lblTurn = Names(2)
  End If
  
  frmMain.Labels(7) = Names(1) & " Time"
  frmMain.Labels(8) = Names(2) & " Time"
  
  For PieceN = 1 To 20
    If Currentpieces(PieceN).X <> 0 And Currentpieces(PieceN).Index <> OFF_BOARD Then Score(1) = Score(1) + 1
  Next
  For PieceN = 21 To 40
    If Currentpieces(PieceN).X <> 0 And Currentpieces(PieceN).Index <> OFF_BOARD Then Score(2) = Score(2) + 1
  Next
  
  frmMain.lblP1Points = Names(1) & " - " & 20 - Score(2)
  frmMain.lblP2Points = Names(2) & " - " & 20 - Score(1)
  
  If Score(1) = 0 Then
    If AutoDebug = False Then
      If frmMain.Option2 Then MsgBox Names(2) & " wins!", vbExclamation Else MsgBox "Had enough yet?", vbExclamation
    End If
    ResetGame
  ElseIf Score(2) = 0 Then
    If AutoDebug = False Then MsgBox Names(1) & " wins!", vbExclamation
    ResetGame
  End If
  
  If VP1Time.Seconds >= 60 Then VP1Time.Minutes = VP1Time.Minutes + Int(VP1Time.Seconds / 60): VP1Time.Seconds = VP2Time.Seconds - (Int(VP2Time.Seconds / 60) * 60)
  If InStr(1, CStr(Round(VP1Time.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
    frmMain.lblP1Time = VP1Time.Minutes & " Min " & Round(VP1Time.Seconds, 1) & ".0 Sec"
  Else
    frmMain.lblP1Time = VP1Time.Minutes & " Min " & Round(VP1Time.Seconds, 1) & " Sec"
  End If
  
  If VP2Time.Seconds >= 60 Then VP2Time.Minutes = VP2Time.Minutes + Int(VP2Time.Seconds / 60): VP2Time.Seconds = VP2Time.Seconds - (Int(VP2Time.Seconds / 60) * 60)
  If InStr(1, CStr(Round(VP2Time.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
    frmMain.lblP2Time = VP2Time.Minutes & " Min " & Round(VP2Time.Seconds, 1) & ".0 Sec"
  Else
    frmMain.lblP2Time = VP2Time.Minutes & " Min " & Round(VP2Time.Seconds, 1) & " Sec"
  End If
  
  TotalTime.Seconds = VP2Time.Seconds + VP1Time.Seconds
  TotalTime.Minutes = VP2Time.Minutes + VP2Time.Minutes
  If TotalTime.Seconds >= 60 Then TotalTime.Minutes = TotalTime.Minutes + Int(TotalTime.Seconds / 60): TotalTime.Seconds = TotalTime.Seconds - (Int(TotalTime.Seconds / 60) * 60)
  If InStr(1, CStr(Round(TotalTime.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
    frmMain.lblTotalTime = TotalTime.Minutes & " Min " & Round(TotalTime.Seconds, 1) & ".0 Sec"
  Else
    frmMain.lblTotalTime = TotalTime.Minutes & " Min " & Round(TotalTime.Seconds, 1) & " Sec"
  End If
  
  frmMain.lblTurns = VTurns
  
End Sub

Public Sub RefreshBoard(Pieces() As PieceAttr)
Dim PieceN As Long, X As Long, Y As Long, ShapeN As Long

For PieceN = 1 To 40
  If Pieces(PieceN).X <> 0 And Pieces(PieceN).X <> 0 Then
    ShapeN = IndexTranslation(IConvert(Pieces(PieceN).X, Pieces(PieceN).Y))
    If Pieces(PieceN).Player = 1 Then
      If Pieces(PieceN).Double Then
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(2).Picture 'frmMain.ImageList1.ListImages(2).Picture
      Else
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(1).Picture
      End If
    Else
      If Pieces(PieceN).Double Then
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(4).Picture
      Else
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(3).Picture
      End If
    End If
  End If
Next

For X = 1 To 10
  For Y = 1 To 10
    ShapeN = IConvert(X, Y)
    If frmMain.Shape1(ShapeN).Picture.Handle <> 0 And CheckSquare(Currentpieces, , , IndexTranslation(ShapeN)).IsPiece = False Then frmMain.Shape1(ShapeN).Picture = Nothing
  Next Y
Next X

End Sub

Public Sub StateRefreshBoard(Pieces() As StatePieceAttr)
Dim PieceN As Long, X As Long, Y As Long, ShapeN As Long

For PieceN = 1 To 40
  XYConvert CLng(Pieces(PieceN).Index), X, Y
  If X <> 0 And Y <> 0 Then
    ShapeN = IndexTranslation(CLng(Pieces(PieceN).Index))
    If PieceN < 21 Then
      If Pieces(PieceN).Double Then
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(2).Picture 'frmMain.ImageList1.ListImages(2).Picture
      Else
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(1).Picture
      End If
    Else
      If Pieces(PieceN).Double Then
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(4).Picture
      Else
        frmMain.Shape1(ShapeN).Picture = frmMain.ImageList1.ListImages(3).Picture
      End If
    End If
  End If
Next

For X = 1 To 10
  For Y = 1 To 10
    ShapeN = IConvert(X, Y)
    If frmMain.Shape1(ShapeN).Picture.Handle <> 0 And StateCheckSquare(Pieces, IndexTranslation(ShapeN)).Index = NOT_PIECE Then frmMain.Shape1(ShapeN).Picture = Nothing
  Next Y
Next X

End Sub

Public Function IConvert(X As Long, Y As Long) As Long
  If X > 10 Or X < 1 Or Y > 10 Or Y < 1 Then
    IConvert = OFF_BOARD
  Else
    IConvert = ((Y - 1) * 10) + X - 1
  End If
End Function

Public Sub XYConvert(Index As Long, ByRef X As Long, ByRef Y As Long)
  If Index > 99 Or Index < 0 Then X = 0: Y = 0: Exit Sub
  Y = (Index - (Index Mod 10)) / 10 + 1
  X = (Index Mod 10) + 1
End Sub

Public Sub CopyArray(ByRef Source() As PieceAttr, Desination() As StatePieceAttr)
Dim PieceN As Long

For PieceN = 1 To 40
  Desination(PieceN).Double = Source(PieceN).Double
  Desination(PieceN).Index = Source(PieceN).Index
Next

End Sub

Public Function StateCheckWin(Pieces() As StatePieceAttr) As Boolean
Dim PieceN As Long, Num1 As Long, Num2 As Long

For PieceN = 1 To 20
  If Pieces(PieceN).Index <> OFF_BOARD Then Num1 = Num1 + 1
Next

For PieceN = 21 To 40
  If Pieces(PieceN).Index <> OFF_BOARD Then Num2 = Num2 + 1
Next

If Num1 = 0 Or Num2 = 0 Then StateCheckWin = True

End Function

Public Function CheckWin(Pieces() As PieceAttr) As Boolean
Dim PieceN As Long, Num1 As Long, Num2 As Long

For PieceN = 1 To 20
  If Pieces(PieceN).Index <> OFF_BOARD Then Num1 = Num1 + 1
Next

For PieceN = 21 To 40
  If Pieces(PieceN).Index <> OFF_BOARD Then Num2 = Num2 + 1
Next

If Num1 = 0 Or Num2 = 0 Then CheckWin = True

End Function

Public Function StateCountpieces(Pieces() As StatePieceAttr) As Long
Dim PieceN As Long, Num1 As Long

For PieceN = 1 To 40
  If Pieces(PieceN).Index <> OFF_BOARD Then
    Num1 = Num1 + 1
  End If
Next

StateCountpieces = Num1

End Function

Public Function Countpieces(Pieces() As PieceAttr) As Long
Dim PieceN As Long, Num1 As Long

For PieceN = 1 To 40
  If Pieces(PieceN).Index <> OFF_BOARD Then
    Num1 = Num1 + 1
  End If
Next

Countpieces = Num1

End Function

Public Function SaveSettings()
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Move Speed", MoveSpeed
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Force Move", ForceMove
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Gametype Mode", PlayerType
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Alpha Beta Pruning Mode", ABPMode
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Max Thought Time", MaxThoughtTime
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Pruning Threshold", PruneThreshold
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Memory Limit", MemoryLimit
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Cheat", CheatSwitch
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 1", , Names(1)
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 2", , Names(2)
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Advanced", IsAdvanced
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Autoswitch", AutoSwitch
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "MaxDepth", MaxDepth
End Function

Public Function GetSettings()
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Move speed", MoveSpeed
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Force Move", ForceMove
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Gametype Mode", PlayerType
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Alpha Beta Pruning Mode", ABPMode
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Max Thought Time", MaxThoughtTime
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Pruning Threshold", PruneThreshold
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Memory Limit", MemoryLimit
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Cheat", CheatSwitch
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 1", , Names(1)
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 2", , Names(2)
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Advanced", IsAdvanced
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Autoswitch", AutoSwitch
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "MaxDepth", MaxDepth
End Function

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional ByRef KeyVal As Long, Optional ByRef KeyValStr As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = Val(tmpVal)                                ' Copy String Value
        KeyValStr = tmpVal
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Val(Format$("&h" + KeyVal))                ' Convert Double Word To String
        KeyValStr = CStr(Format$("&h" + KeyVal))
    End Select
    
    GetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = 0                                             ' Set Return Val To Empty String
    KeyValStr = ""
    GetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function SetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional ByRef KeyVal As Long, Optional ByRef KeyValStr As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If RC <> ERROR_SUCCESS Then
      If RC = ERROR_FILE_NOT_FOUND Then
        RC = RegCreateKey(KeyRoot, KeyName, hKey)
      Else
        GoTo GetKeyError          ' Handle Error...
      End If
    End If
    
    If KeyValStr = "" Then KeyValStr = CStr(KeyVal)
    KeyValSize = Len(KeyValStr) + 1
    KeyValType = REG_SZ
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegSetValueEx(hKey, SubKeyRef, 0, KeyValType, ByVal KeyValStr, KeyValSize)
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    SetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    SetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


