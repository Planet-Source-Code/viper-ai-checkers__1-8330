Attribute VB_Name = "AI"
Option Explicit
Dim ForceScore As Long

Public Sub AIMove(ByVal Player As Long)
Dim Depth As Long, Lng1 As Long, Chosenmove As Long, High As Long, Low As Long
Dim TempMove() As ArrayPieces, Time As Long, ChosenPiece As Long, ParentID As Long
Dim ForcedScore As Long, Lng2 As Long, EqualToScore As Long, TempCount As Long
Dim SubBranch As Long

RestartAI:

ERR7 = False 'toggle outof memory switch off
With frmMain
  .StatusBar1.SimpleText = "Initialising AI engine..."
  .MousePointer = MousePointerConstants.vbHourglass 'change mouse pointer to wait
  .cmdSwitch.Enabled = False
End With

Time = GetTickCount 'record current time

ReDim SMatrix(0 To 0)
ReDim MoveMatrix(0 To 0)
UpperMove = 0
UpperS = 0
SMatrix(0).Depth = 1
SMatrix(0).ArrayNum = 0
CopyArray Currentpieces, MoveMatrix(0).Pieces 'copy current boardstate into movematrix(0)

Lng1 = 0

If Player = 1 Then Low = 1: High = 20 Else Low = 21: High = 40 'get piece ranges

frmMain.StatusBar1.SimpleText = "Generating moves..."

For Depth = 1 To MaxDepth
  Do While SMatrix(Lng1).Depth = Depth
    If Depth <> 1 And ForceMove Then
      EqualToScore = 0
      ForcedScore = 50
      
      For Lng2 = UpperS To 1 Step -1
        If SMatrix(Lng2).Depth = Depth Then
          ParentID = SMatrix(Lng2).ParentID
          For SubBranch = Lng2 To 1 Step -1
            If SMatrix(SubBranch).ParentID = ParentID Then
              TempCount = StateCountpieces(MoveMatrix(SMatrix(SubBranch).ArrayNum).Pieces)
              If TempCount < ForcedScore Then
                ForcedScore = TempCount: EqualToScore = 1
              ElseIf TempCount = ForcedScore Then
                EqualToScore = EqualToScore + 1
              End If
            Else
              Exit For
            End If
          Next SubBranch
          
          For SubBranch = Lng2 To SubBranch Step -1
            If StateCountpieces(MoveMatrix(SMatrix(SubBranch).ArrayNum).Pieces) <> ForcedScore Then
              With SMatrix(SubBranch)
                .Finished = LEAVE
                .Depth = 0
                .ParentID = 0
              End With
            End If
          Next SubBranch
          Lng2 = SubBranch
        End If
      Next Lng2
      If EqualToScore = 1 And Depth = 2 Then GoTo SkipGenerate
    End If
    
    If SMatrix(Lng1).Finished = NORMAL Then
      GenerateStates Lng1, False, Player
      If ERR7 = True Then Exit For 'exit if error or out of memory
      If Int((GetTickCount - Time) / 1000) > MaxThoughtTime And MaxThoughtTime <> 0 Then 'if maximum time has been reached
        frmMain.lblPlyDepth = Depth + 1
        frmMain.lblMMatrixSize = UpperMove
        Exit For
      End If
    End If
    Lng1 = Lng1 + 1
  Loop
  frmMain.lblMMatrixSize = UpperMove 'Move stats
  frmMain.lblPlyDepth = Depth 'Depth stats
  DoEvents 'prevents freezing....ish
  Player = Player Mod 2 + 1
Next Depth

If ABPMode = 1 Then

'---------Simple Alpha Beta routine----
  Lng1 = 1
  
  If ERR7 = False Then
    Do While UpperS >= Lng1 'do loop used because uppers will change and a for loop would only check it once
      If SMatrix(Lng1).Depth >= MaxDepth And SMatrix(Lng1).Finished <> LEAVE Then
        If StateCountpieces(MoveMatrix(SMatrix(SMatrix(Lng1).ParentID).ArrayNum).Pieces) <> StateCountpieces(MoveMatrix(SMatrix(Lng1).ArrayNum).Pieces) Then 'if score boardstate has changed in the last 2 moves of the branch
          GenerateStates Lng1, True, Player
          SMatrix(SMatrix(Lng1).ParentID).Score = 0 'Erase lower score in favour of higher depths
        End If
      End If
      Lng1 = Lng1 + 1
    Loop
    frmMain.lblMMatrixSize = UpperMove
  End If

End If

SkipGenerate:

frmMain.StatusBar1.SimpleText = "Evaluating moves..."

DoEvents

Chosenmove = ChooseMove(Player) 'store smatrix arraynumber to move into chosenmove

For Lng1 = Low To High 'check which piece to move
  If MoveMatrix(SMatrix(Chosenmove).ArrayNum).Pieces(Lng1).Index <> Currentpieces(Lng1).Index Then Exit For
Next

ChosenPiece = Lng1

If SMatrix(Chosenmove).SubParentId <> 0 Then
  For Lng1 = Low To High
    If MoveMatrix(SMatrix(SMatrix(Chosenmove).SubParentId).ArrayNum).Pieces(Lng1).Index <> Currentpieces(Lng1).Index Then Exit For
  Next
End If

If ChosenPiece < Lng1 Then
  ChosenPiece = Lng1
End If

If ChosenPiece > High Then 'If it hasn't moved a piece
  MsgBox Names(1) & " wins!", vbExclamation
  ReDim SMatrix(0 To 0)
  ReDim MoveMatrix(0 To 0)
  frmMain.MousePointer = MousePointerConstants.vbArrow
  ResetGame
  Exit Sub
End If

With frmMain
  .Shape1(IndexTranslation(Currentpieces(ChosenPiece).Index)).Picture = frmMain.ImageList1.ListImages(5).Picture 'Select piece to move
  .StatusBar1.SimpleText = ""
  .cmdSwitch.Enabled = True
End With

Sleep MoveSpeed

'StateRefreshBoard MoveMatrix(Chosenmove).Pieces
'RefreshBoard Currentpieces

If MovePiece(Val(Currentpieces(ChosenPiece).Index), Val(MoveMatrix(SMatrix(Chosenmove).ArrayNum).Pieces(Lng1).Index), Chosenmove, False) = False Then
  MsgBox "There has been an error moving the computer piece", vbExclamation
  Stop
  StateRefreshBoard MoveMatrix(Chosenmove).Pieces
  RefreshBoard Currentpieces
End If

ReDim SMatrix(0 To 0)
ReDim MoveMatrix(0 To 0)
UpperS = 0
UpperMove = 0

'---------Display Times---------

VP2Time.Seconds = VP2Time.Seconds + Round((GetTickCount - Time) / 1000, 1)
If VP2Time.Seconds >= 60 Then VP2Time.Minutes = VP2Time.Minutes + Int(VP2Time.Seconds / 60): VP2Time.Seconds = VP2Time.Seconds - (Int(VP2Time.Seconds / 60) * 60)
If InStr(1, CStr(Round(VP2Time.Seconds, 1)), ".", vbBinaryCompare) = 0 Then
  frmMain.lblP2Time = VP2Time.Minutes & " Min " & Round(VP2Time.Seconds, 1) & ".0 Sec"
Else
  frmMain.lblP2Time = VP2Time.Minutes & " Min " & Round(VP2Time.Seconds, 1) & " Sec"
End If

frmMain.MousePointer = MousePointerConstants.vbArrow

End Sub

Public Sub GenerateStates(ParentID As Long, TakingOnly As Boolean, Player As Long)
On Error GoTo Error
Dim PieceN As Long, Direction As Long, NewIndex As Long, ArrayNum As Long
Dim Low As Long, High As Long, TempSquare As StateSelectedSquare, NewIndex2 As Long
Dim Lng1 As Long, X As Long, Y As Long, Hasmoved As Boolean
Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, MoveLength As Long, MaxMoveLength As Long

If ERR7 Then Exit Sub

If Player = 1 Then Low = 1: High = 20 Else Low = 21: High = 40 'Get piece ranges

ArrayNum = SMatrix(ParentID).ArrayNum 'Get Movematrix array number of parent move

For PieceN = Low To High
  If MoveMatrix(ArrayNum).Pieces(PieceN).Index = OFF_BOARD Then GoTo Next_Piece 'if piece is dead
  Y = (MoveMatrix(ArrayNum).Pieces(PieceN).Index - (MoveMatrix(ArrayNum).Pieces(PieceN).Index Mod 10)) / 10 + 1
  X = (MoveMatrix(ArrayNum).Pieces(PieceN).Index Mod 10) + 1
  If MoveMatrix(ArrayNum).Pieces(PieceN).Double = True Then MaxMoveLength = 10 Else MaxMoveLength = 1
  For Direction = 1 To 4
    For MoveLength = 1 To MaxMoveLength
      If ERR7 Then Exit Sub
      NewIndex = MoveMatrix(ArrayNum).Pieces(PieceN).Index + IndexMoves(Direction, MoveLength)
      X1 = X + XYMoves(Direction, MoveLength).X
      Y1 = Y + XYMoves(Direction, MoveLength).Y
      If X1 < 1 Or X1 > 10 Or Y1 < 1 Or Y1 > 10 Then GoTo Next_Direction
      
      TempSquare = StateCheckSquare(MoveMatrix(ArrayNum).Pieces, NewIndex)
      If TempSquare.Index = OFF_BOARD Then GoTo Next_Direction
      If TempSquare.Index <> NOT_PIECE Then
        '---------Place over piece to take--------
        If TempSquare.Player = Player Then GoTo Next_Direction
        NewIndex2 = MoveMatrix(ArrayNum).Pieces(PieceN).Index + IndexMoves(Direction + 4, MoveLength)
        X2 = X + XYMoves(Direction + 4, MoveLength).X
        Y2 = Y + XYMoves(Direction + 4, MoveLength).Y
        If X2 < 1 Or X2 > 10 Or Y2 < 1 Or Y2 > 10 Then GoTo Next_Direction
        
        If StateCheckSquare(MoveMatrix(ArrayNum).Pieces, NewIndex2).Index = NOT_PIECE Then 'if there is a piece to take and it's not ours!
          UpperMove = UpperMove + 1
          If UpperMove + 100 >= UBound(MoveMatrix) Then ReDim Preserve MoveMatrix(0 To UBound(MoveMatrix) + ReDimInterval) 'Call ReDimension(ReDimInterval, False, True)
          For Lng1 = 1 To 40 'Record new boardstate
            MoveMatrix(UpperMove).Pieces(Lng1) = MoveMatrix(ArrayNum).Pieces(Lng1)
          Next Lng1
          
          MoveMatrix(UpperMove).Pieces(PieceN).Index = NewIndex2 'Change the current piece co-ordinate
          
          MoveMatrix(UpperMove).Pieces(TempSquare.Piece).Index = OFF_BOARD ' Delete taken piece
          
          Select Case Player 'Set piece as a double if it reaches the other side
            Case 1
              If NewIndex2 > 0 And NewIndex2 < 10 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = True
            Case 2
              If NewIndex2 > 89 And NewIndex2 < 99 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = True
          End Select
          
          '------Save stuff about move----------
          
          UpperS = UpperS + 1
          If UpperS + 100 >= UBound(SMatrix) Then ReDim Preserve SMatrix(0 To UBound(SMatrix) + ReDimInterval) 'Call ReDimension(ReDimInterval, True, False)
          Hasmoved = True
          With SMatrix(UpperS)
            .Depth = SMatrix(ParentID).Depth + 1 'increase depth
            .ArrayNum = UpperS
            .ParentID = ParentID
            If StateCheckWin(MoveMatrix(UpperMove).Pieces) Then
              If .Depth Mod 2 = 0 Then 'Is Player 2 turn (Computer)
                .Score = 1000 / (.Depth - 1)
                .Finished = LEAVE
              Else
                .Score = 1000 / (-.Depth - 1)
                .Finished = LEAVE
              End If
            Else
              .Score = EvalBoard(MoveMatrix(UpperMove).Pieces)
              Call GenerateMultiStates(UpperS, PieceN, Player) 'if computer has taken then check if it can take again
            End If
            
          End With
          
        End If
        
        GoTo Next_Direction
  
      ElseIf TakingOnly = False Then
        If Player = 1 Then 'if it can move in that direction
          If MoveMatrix(ArrayNum).Pieces(PieceN).Double = False And (Direction = 2 Or Direction = 3) Then GoTo Next_Direction
        Else
          If MoveMatrix(ArrayNum).Pieces(PieceN).Double = False And (Direction = 1 Or Direction = 4) Then GoTo Next_Direction
        End If
      
        UpperMove = UpperMove + 1
        If UpperMove >= UBound(MoveMatrix) Then ReDim Preserve MoveMatrix(0 To UBound(MoveMatrix) + ReDimInterval) 'Call ReDimension(ReDimInterval, False, True)
        For Lng1 = 1 To 40 'Save boardstate
          MoveMatrix(UpperMove).Pieces(Lng1) = MoveMatrix(SMatrix(ParentID).ArrayNum).Pieces(Lng1)
        Next Lng1
        
        MoveMatrix(UpperMove).Pieces(PieceN).Index = NewIndex ' Change index of piece which has moved
        
        Select Case Player 'Set piece as a double if it reaches the other side
          Case 1
            If NewIndex > 0 And NewIndex < 10 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = True
          Case 2
            If NewIndex > 89 And NewIndex < 99 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = True
        End Select
        
        '-------------Save stuff about move---------
        
        UpperS = UpperS + 1
        If UpperS >= UBound(SMatrix) Then ReDim Preserve SMatrix(0 To UBound(SMatrix) + ReDimInterval) 'Call ReDimension(ReDimInterval, True, False)
        Hasmoved = True
        With SMatrix(UpperS)
          .Depth = SMatrix(ParentID).Depth + 1
          .ArrayNum = UpperMove
          .ParentID = ParentID
          .Score = EvalBoard(MoveMatrix(UpperMove).Pieces)
        End With
      End If

    Next MoveLength
Next_Direction:
  Next Direction
Next_Piece:
Next PieceN

'--------Changes the score of parent move to a losing (or winning) score if there are no moves possilbe from the branch

If Hasmoved = False Then
  With SMatrix(ParentID)
    If .Depth Mod 2 = 0 Then 'Is Player 2 turn (Computer)
      .Score = 1000 / (.Depth - 1)
      .Finished = CANNOTMOVE
    Else
      .Score = 1000 / (-.Depth - 1)
      .Finished = CANNOTMOVE
    End If
  End With
End If

Exit Sub
Error:
  If Err.Number = 7 Then ERR7 = True
  MemoryLimit = UBound(SMatrix)
  MsgBox Err.Description & vbCrLf & "GenerateStates", vbExclamation, "ERR_" & Err.Number
  Err.Clear
End Sub

Private Sub GenerateMultiStates(ByVal ParentID As Long, PieceN As Long, Player As Long)
On Error GoTo Error
Dim Direction As Long, NewIndex As Long, ArrayNum As Long
Dim TempSquare As StateSelectedSquare, NewIndex2 As Long
Dim Lng1 As Long, X As Long, Y As Long, Hasmoved As Boolean
Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Index As Long

If ERR7 Then Exit Sub

ArrayNum = SMatrix(ParentID).ArrayNum 'Get movematrix array number from parent

  If MoveMatrix(ArrayNum).Pieces(PieceN).Index = OFF_BOARD Then Exit Sub
  Index = MoveMatrix(ArrayNum).Pieces(PieceN).Index
  Y = (Index - (Index Mod 10)) / 10 + 1
  X = (Index Mod 10) + 1
  'If X = 5 And Y = 4 Then Stop
  For Direction = 1 To 4
    If ERR7 Then Exit Sub
    
    NewIndex = Index + IndexMoves(Direction, 1)
    X1 = X + XYMoves(Direction, 1).X
    Y1 = Y + XYMoves(Direction, 1).Y
    If X1 < 1 Or X1 > 10 Or Y1 < 1 Or Y1 > 10 Then GoTo Next_Direction
    
    TempSquare = StateCheckSquare(MoveMatrix(ArrayNum).Pieces, NewIndex)
    If TempSquare.Index = OFF_BOARD Then GoTo Next_Direction
    If TempSquare.Index <> NOT_PIECE Then
    
      If TempSquare.Player = Player Then GoTo Next_Direction
      NewIndex2 = Index + IndexMoves(Direction + 4, 1)
      X2 = X + XYMoves(Direction + 4, 1).X
      Y2 = Y + XYMoves(Direction + 4, 1).Y
      If X2 < 1 Or X2 > 10 Or Y2 < 1 Or Y2 > 10 Then GoTo Next_Direction
      
      If StateCheckSquare(MoveMatrix(ArrayNum).Pieces, NewIndex2).Index = NOT_PIECE Then
        UpperMove = UpperMove + 1
        For Lng1 = 1 To 40
          MoveMatrix(UpperMove).Pieces(Lng1) = MoveMatrix(ArrayNum).Pieces(Lng1)
        Next Lng1
        
        MoveMatrix(UpperMove).Pieces(PieceN).Index = NewIndex2
        
        MoveMatrix(UpperMove).Pieces(TempSquare.Piece).Index = OFF_BOARD
        
        Select Case Player
          Case 1
            If Index >= 0 And Index < 10 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = False
          Case 2
            If Index > 89 And Index < 99 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = False
        End Select
        
        Select Case Player
          Case 1
            If NewIndex2 >= 0 And NewIndex2 < 10 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = True
          Case 2
            If NewIndex2 > 89 And NewIndex2 < 9 Then MoveMatrix(UpperMove).Pieces(PieceN).Double = True
        End Select
        
        UpperS = UpperS + 1
        Hasmoved = True
        With SMatrix(UpperS)
          .Depth = SMatrix(ParentID).Depth
          .ArrayNum = UpperMove
          .ParentID = SMatrix(ParentID).ParentID
          .SubParentId = ParentID
          If StateCheckWin(MoveMatrix(UpperMove).Pieces) Then
            If .Depth Mod 2 = 0 Then 'Is Player 2 turn (Computer)
              .Score = 1000 / (.Depth - 1)
              .Finished = LEAVE
            Else
              .Score = 1000 / (-.Depth - 1)
              .Finished = LEAVE
            End If
          Else
            .Score = EvalBoard(MoveMatrix(UpperMove).Pieces)
            Call GenerateMultiStates(UpperS, PieceN, Player)
          End If
          
        End With
        
      End If
    
    Else
    
      If Player = 1 Then 'if it can move in that direction
        If MoveMatrix(ArrayNum).Pieces(PieceN).Double = False And (Direction = 2 Or Direction = 3) Then GoTo Next_Direction
      Else
        If MoveMatrix(ArrayNum).Pieces(PieceN).Double = False And (Direction = 1 Or Direction = 4) Then GoTo Next_Direction
      End If
      
      Hasmoved = True

    End If
  
Next_Direction:
  Next Direction
  
'--------Changes the score of parent move to a losing (or winning) score if there are no moves possilbe from the branch

If Hasmoved = False Then
  With SMatrix(ParentID)
    If .Depth Mod 2 = 0 Then 'Is Player 2 turn (Computer)
      .Finished = CANNOTMOVE
    Else
      .Finished = CANNOTMOVE
    End If
  End With
End If

Exit Sub
Error:
  If Err.Number = 7 Then ERR7 = True
  MemoryLimit = UBound(SMatrix)
  MsgBox Err.Description & vbCrLf & "GenerateMultiStates", vbExclamation, "ERR_" & Err.Number
  Err.Clear
End Sub

Private Function EvalBoard(Pieces() As StatePieceAttr) As Long
Dim PieceN As Long, Num1 As Long, Num2 As Long, Ratio As Double, Score As Long
Dim High1 As Long, High2 As Long

If Turn = 2 Then
  High1 = 20
  High2 = 40
Else
  High1 = 40
  High2 = 20
End If

For PieceN = (High1 - 19) To High1
  If Pieces(PieceN).Index <> OFF_BOARD Then
    If Pieces(PieceN).Index < 99 And Pieces(PieceN).Index > 89 And Pieces(PieceN).Double = False Then Score = Score - CStrategy 'Makes com keep back line as long as poss
    If Pieces(PieceN).Double Then
      Num1 = Num1 + CDouble
    Else
      Num1 = Num1 + CSingle
    End If
  End If
Next

For PieceN = (High2 - 19) To High2
  If Pieces(PieceN).Index <> OFF_BOARD Then
    If Pieces(PieceN).Index < 10 And Pieces(PieceN).Index >= 0 And Pieces(PieceN).Double = False Then Score = Score + CStrategy 'Makes com keep back line as long as poss
    If Pieces(PieceN).Double Then
      Num2 = Num2 + CDouble
    Else
      Num2 = Num2 + CSingle
    End If
  End If
Next

'A ratio is used so it will make a swap if is winning etc
'Score comes from strategic points

Ratio = Num2 / Num1
EvalBoard = Int((Ratio * 100) + Score)

End Function

Public Function ChooseMove(Player As Long) As Long
Dim Depth As Long, StateID As Long, MaxScore As Long, MaxSources() As Long, Lng1 As Long
Dim ABMaxDepth As Long, FirstDepthEval() As Long, MaxSources2() As Long

ReDim MaxSources(1 To 1)
ReDim FirstDepthEval(0 To 0)

'-------Saves Depth 1 Eval Score------

For StateID = 0 To UpperS
  If SMatrix(StateID).Depth <= 2 Then
    ReDim Preserve FirstDepthEval(0 To StateID)
    If StateCheckWin(MoveMatrix(UpperMove).Pieces) Then
      FirstDepthEval(UBound(FirstDepthEval)) = 1000
    Else
      FirstDepthEval(UBound(FirstDepthEval)) = EvalBoard(MoveMatrix(UpperMove).Pieces)
    End If
  Else
    Exit For
  End If
Next

'-------Gets maximum depth (which is changed because of the alpha beta pruning)----

For StateID = 0 To UpperS
  If SMatrix(StateID).Depth > ABMaxDepth Then ABMaxDepth = SMatrix(StateID).Depth
Next

'-------Erases current scores used for Alpha Beta Branching------------

If ABMaxDepth < MaxDepth Then Lng1 = ABMaxDepth Else Lng1 = MaxDepth

For StateID = 0 To UpperS
  If SMatrix(StateID).Depth < Lng1 And SMatrix(StateID).Finished <> LEAVE Then SMatrix(StateID).Score = 0
Next

'-------Passes up socer using MinMax to the depth 2 moves (depth 1 is original board state)

For Depth = ABMaxDepth To 2 Step -1
  For StateID = UpperS To 0 Step -1
    If SMatrix(StateID).Depth = Depth Then
      If Depth Mod 2 = 0 Then 'current player turn
        If SMatrix(StateID).Score > SMatrix(SMatrix(StateID).ParentID).Score Or SMatrix(SMatrix(StateID).ParentID).Score = 0 Then
          SMatrix(SMatrix(StateID).ParentID).Score = SMatrix(StateID).Score
        End If
      Else
        If SMatrix(StateID).Score < SMatrix(SMatrix(StateID).ParentID).Score Or SMatrix(SMatrix(StateID).ParentID).Score = 0 Then
          SMatrix(SMatrix(StateID).ParentID).Score = SMatrix(StateID).Score
        End If
      End If
    End If
  Next StateID
Next Depth

MaxScore = -1000 'Forces move (even if it is sub zero)

'actually chooses piece(s) to move

For StateID = 0 To UpperS
  If SMatrix(StateID).Depth = 2 Then
    If SMatrix(StateID).Score > MaxScore Then
      ReDim MaxSources(1 To 1)
      MaxSources(1) = StateID
      MaxScore = SMatrix(StateID).Score
    ElseIf SMatrix(StateID).Score = MaxScore Then
      ReDim Preserve MaxSources(1 To UBound(MaxSources) + 1)
      MaxSources(UBound(MaxSources)) = StateID
    End If
  End If
Next

MaxScore = -1000

'----Cross-references equal moves (in maxsources) to that of first level moves preceding them----

If ForceMove = False Then
  For StateID = 1 To UBound(MaxSources)
    If FirstDepthEval(MaxSources(StateID)) > MaxScore Then
      ReDim MaxSources2(1 To 1)
      MaxSources2(1) = StateID
      MaxScore = FirstDepthEval(StateID)
    ElseIf FirstDepthEval(StateID) = MaxScore Then
      ReDim Preserve MaxSources2(1 To UBound(MaxSources2) + 1)
      MaxSources2(UBound(MaxSources2)) = StateID
    End If
  Next
  ChooseMove = MaxSources(Int(UBound(MaxSources2) * Rnd(Left(GetTickCount, 2)) + 1)) 'Randomly chooses one of the best
Else
  ChooseMove = MaxSources(Int(UBound(MaxSources) * Rnd(Left(GetTickCount, 2)) + 1)) 'Randomly chooses one of the best
End If

End Function
