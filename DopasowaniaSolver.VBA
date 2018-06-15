Sub START()

For sym = 1 To 7

Sheets("dane").Select
Range("A:C").Clear

If sym = 1 Then
Range("A:C").Value = Range("D:F").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If


If sym = 2 Then
Range("A:C").Value = Range("G:I").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 3 Then
Range("A:C").Value = Range("J:L").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 4 Then
Range("A:C").Value = Range("M:O").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 5 Then
Range("A:C").Value = Range("P:R").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 6 Then
Range("A:C").Value = Range("S:U").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 7 Then
Range("A:C").Value = Range("V:X").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 8 Then
Range("A:C").Value = Range("Y:AA").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 9 Then
Range("A:C").Value = Range("AB:AD").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

If sym = 10 Then
Range("A:C").Value = Range("AE:AG").Value
Range("A:C").Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If


Call symul


Next sym



End Sub
Function ShowTrial(Reason As Integer)
'  MsgBox Reason
  ShowTrial = 3
End Function

Sub statyst()
'
' Makro1 Makro
'
'


For blad = 1 To 4

If (blad = 1) Then

Sheets("DANE").Select
Range("D1").Select
Range("D1").Value = 0.01
nblad = 0.01

End If

If (blad = 2) Then

Sheets("DANE").Select
Range("D1").Select
Range("D1").Value = 0.03
nblad = 0.03

End If

If (blad = 3) Then

Sheets("DANE").Select
Range("D1").Select
Range("D1").Value = 0.05
nblad = 0.05

End If

If (blad = 4) Then

Sheets("DANE").Select
Range("D1").Select
Range("D1").Value = 0.1
nblad = 0.1

End If


For sym = 2 To 100

nazwa = "NBLUK00_ERR0_" & nblad & "_" & sym & ""

ActiveWorkbook.SaveAs Filename:="D:\_Projekty\BLUEGAS\szumy\Statystyka\DUONG_" & nazwa & ".xlsm", FileFormat:=52

Sheets("DANE").Select

    Range("C1").Select
    Range("C1").Value = sym


Call symul



Sheets("Wyniki").Select




ActiveWorkbook.Save

Next sym


Next blad

End Sub



Sub symul()
Dim waga051()
Dim waga12()
Dim waga14()
Dim waga15()
Dim czas()
Dim FGPR()
Dim FGPT()








Sheets("DANE").Select


nazwa = Range("A1").Value

Sheets("DANE").Select
Columns("A:C").Select
Selection.Copy
Sheets("wagi").Select
Columns("F:H").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Call wagi


Sheets("wagi").Select
t = 0
iwaga = 4
rozm = 1
waga1 = Range("A" & iwaga & "").Value
waga2 = Range("B" & iwaga & "").Value
waga3 = Range("C" & iwaga & "").Value
waga4 = Range("D" & iwaga & "").Value

Do Until waga1 = ""

rozm = rozm + 1

iwaga = iwaga + 1
waga1 = Range("A" & iwaga & "").Value
Loop

rozm = rozm - 1

ReDim waga051(rozm)
ReDim waga12(rozm)
ReDim waga14(rozm)
ReDim waga15(rozm)
ReDim czas(rozm)
ReDim FGPR(rozm)
ReDim FGPT(rozm)

iwaga = 4
waga1 = Range("A" & iwaga & "").Value
waga2 = Range("B" & iwaga & "").Value
waga3 = Range("C" & iwaga & "").Value
waga4 = Range("D" & iwaga & "").Value
qg = Range("G" & iwaga & "").Value
kroki = Range("F" & iwaga & "").Value
GP = Range("H" & iwaga & "").Value

Do Until waga1 = ""

waga051(t) = waga1
waga12(t) = waga2
waga14(t) = waga3
waga15(t) = waga4
czas(t) = kroki
FGPR(t) = qg
FGPT(t) = GP

iwaga = iwaga + 1
waga1 = Range("A" & iwaga & "").Value
waga2 = Range("B" & iwaga & "").Value
waga3 = Range("C" & iwaga & "").Value
waga4 = Range("D" & iwaga & "").Value
qg = Range("G" & iwaga & "").Value
kroki = Range("F" & iwaga & "").Value
GP = Range("H" & iwaga & "").Value

t = t + 1
Loop



'*************************************************dopasowanie 0.5 i 1***********************************************

Sheets("dop 0.5 - 1").Select

Range("A8:AB20000").Clear

przeg = 0
czasprzeg = 0


For i = 0 To rozm - 1
Range("A" & 8 + i & "").Value = czas(i)
Range("B" & 8 + i & "").Value = FGPR(i)
Range("C" & 8 + i & "").Select
   ActiveCell.FormulaR1C1 = _
        "=R3C11/RC[-2]^R5C11*EXP(R4C11*(RC[-2]/RC[-2]^R5C11-1))"

Range("E" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"
Range("F" & 8 + i & "").Value = FGPT(i)

Range("G" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]<R1C11,RC[-1],  (EXP(R4C11*(1/R1C11^R6C11-1))   -   EXP(R4C11*(1/RC[-6]^R6C11-1)))*R7C11+R2C11)"

Range("H" & 8 + i & "").Value = waga051(i)
Range("D" & 8 + i & "").Value = waga051(i)

Range("I" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"






If (przeg < FGPR(i)) Then
przeg = FGPR(i)
czasprzeg = czas(i)
End If

Next i



Range("J" & 8 + i - 1 & "").Value = "=(F" & 8 + i - 1 & "-G" & 8 + i - 1 & ")/F" & 8 + i - 1 & ""






        ' ***************************DUONG*****************************************
    
    FseNew = 0
    FseOLD = 1

    
    Range("L3").Value = 1
    Range("L4").Value = 1
    Range("L5").Value = 1
    pzero = 0.000001
       
    Do Until FseOLD = FseNew
    
    FseOLD = Range("I5").Value
    
    
    
    SolverReset
    SolverAdd "L3", 3, pzero
    SolverAdd "L4", 3, pzero
    SolverAdd "L5", 3, pzero
    
    SolverOk SetCell:="$I$5", MaxMinVal:=2, ValueOf:=0, ByChange:="$L$3:$L$5", _
    Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=1000, Iterations:=100, Precision:=0.000001, _
    Convergence:=0.000000000001, StepThru:=False, Scaling:=False, AssumeNonNeg:= _
    False, Derivatives:=2
    SolverFinish KeepFinal = True
    SolverSolve UserFinish = False, ShowRef:="ShowTrial"
    
    FseNew = Range("I5").Value
    
    Loop
    
    blad051 = -Range("J" & 8 + i - 1 & "").Value
   
For i = 0 To rozm - 1
   If (czas(i) >= P1 And czas(i) <= P3) Then
   Range("L" & 8 + i & "").Value = Range("C" & 8 + i & "").Value
   Range("M" & 8 + i & "").Value = Range("G" & 8 + i & "").Value
   End If
Next i

    
  Sheets("Wyniki").Select
  Range("C2").Value = "Czas przegięcia wydajności"
  Range("D2").Value = czasprzeg
  Range("E4").Value = blad051

'*************************************************dopasowanie 1 i 2***********************************************

Sheets("dop 1 - 2").Select

Range("A8:AB20000").Clear




For i = 0 To rozm - 1
Range("A" & 8 + i & "").Value = czas(i)
Range("B" & 8 + i & "").Value = FGPR(i)
Range("C" & 8 + i & "").Select
   ActiveCell.FormulaR1C1 = _
        "=R3C11/RC[-2]^R5C11*EXP(R4C11*(RC[-2]/RC[-2]^R5C11-1))"

Range("E" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"
Range("F" & 8 + i & "").Value = FGPT(i)

Range("G" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]<R1C11,RC[-1],  (EXP(R4C11*(1/R1C11^R6C11-1))   -   EXP(R4C11*(1/RC[-6]^R6C11-1)))*R7C11+R2C11)"

Range("H" & 8 + i & "").Value = waga12(i)
Range("D" & 8 + i & "").Value = waga12(i)

Range("I" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"








Next i



Range("J" & 8 + i - 1 & "").Value = "=(F" & 8 + i - 1 & "-G" & 8 + i - 1 & ")/F" & 8 + i - 1 & ""






        ' ***************************DUONG*****************************************
    
    FseNew = 0
    FseOLD = 1

    
    Range("L3").Value = 1
    Range("L4").Value = 1
    Range("L5").Value = 1
    pzero = 0.000001
       
    Do Until FseOLD = FseNew
    
    FseOLD = Range("I5").Value
    
    
    
    SolverReset
    SolverAdd "L3", 3, pzero
    SolverAdd "L4", 3, pzero
    SolverAdd "L5", 3, pzero
    
    SolverOk SetCell:="$I$5", MaxMinVal:=2, ValueOf:=0, ByChange:="$L$3:$L$5", _
    Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=1000, Iterations:=100, Precision:=0.000001, _
    Convergence:=0.000000000001, StepThru:=False, Scaling:=False, AssumeNonNeg:= _
    False, Derivatives:=2
    SolverFinish KeepFinal = True
    SolverSolve UserFinish = False, ShowRef:="ShowTrial"
    
    FseNew = Range("I5").Value
    
    Loop
    
    blad12 = -Range("J" & 8 + i - 1 & "").Value
   
For i = 0 To rozm - 1
   If (czas(i) >= P11 And czas(i) <= P31) Then
   Range("L" & 8 + i & "").Value = Range("C" & 8 + i & "").Value
   Range("M" & 8 + i & "").Value = Range("G" & 8 + i & "").Value
   End If
Next i

    
  Sheets("Wyniki").Select
  Range("E5").Value = blad12

'*************************************************dopasowanie 1 i 3***********************************************

Sheets("dop 1 - 3").Select

Range("A8:AB20000").Clear




For i = 0 To rozm - 1
Range("A" & 8 + i & "").Value = czas(i)
Range("B" & 8 + i & "").Value = FGPR(i)
Range("C" & 8 + i & "").Select
   ActiveCell.FormulaR1C1 = _
        "=R3C11/RC[-2]^R5C11*EXP(R4C11*(RC[-2]/RC[-2]^R5C11-1))"

Range("E" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"
Range("F" & 8 + i & "").Value = FGPT(i)

Range("G" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]<R1C11,RC[-1],  (EXP(R4C11*(1/R1C11^R6C11-1))   -   EXP(R4C11*(1/RC[-6]^R6C11-1)))*R7C11+R2C11)"

Range("H" & 8 + i & "").Value = waga14(i)
Range("D" & 8 + i & "").Value = waga14(i)

Range("I" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"








Next i



Range("J" & 8 + i - 1 & "").Value = "=(F" & 8 + i - 1 & "-G" & 8 + i - 1 & ")/F" & 8 + i - 1 & ""






        ' ***************************DUONG*****************************************
    
    FseNew = 0
    FseOLD = 1

    
    Range("L3").Value = 1
    Range("L4").Value = 1
    Range("L5").Value = 1
    pzero = 0.000001
       
    Do Until FseOLD = FseNew
    
    FseOLD = Range("I5").Value
    
    
    
    SolverReset
    SolverAdd "L3", 3, pzero
    SolverAdd "L4", 3, pzero
    SolverAdd "L5", 3, pzero
    
    SolverOk SetCell:="$I$5", MaxMinVal:=2, ValueOf:=0, ByChange:="$L$3:$L$5", _
    Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=1000, Iterations:=100, Precision:=0.000001, _
    Convergence:=0.000000000001, StepThru:=False, Scaling:=False, AssumeNonNeg:= _
    False, Derivatives:=2
    SolverFinish KeepFinal = True
    SolverSolve UserFinish = False, ShowRef:="ShowTrial"
    
    FseNew = Range("I5").Value
    
    Loop
    
    blad14 = -Range("J" & 8 + i - 1 & "").Value
   
For i = 0 To rozm - 1
   If (czas(i) >= P12 And czas(i) <= P32) Then
   Range("L" & 8 + i & "").Value = Range("C" & 8 + i & "").Value
   Range("M" & 8 + i & "").Value = Range("G" & 8 + i & "").Value
   End If
Next i

    
  Sheets("Wyniki").Select
  Range("E6").Value = blad14

'*************************************************dopasowanie 1 i 4***********************************************

Sheets("dop 1 - 4").Select

Range("A8:AB20000").Clear




For i = 0 To rozm - 1
Range("A" & 8 + i & "").Value = czas(i)
Range("B" & 8 + i & "").Value = FGPR(i)
Range("C" & 8 + i & "").Select
   ActiveCell.FormulaR1C1 = _
        "=R3C11/RC[-2]^R5C11*EXP(R4C11*(RC[-2]/RC[-2]^R5C11-1))"

Range("E" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"
Range("F" & 8 + i & "").Value = FGPT(i)

Range("G" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]<R1C11,RC[-1],  (EXP(R4C11*(1/R1C11^R6C11-1))   -   EXP(R4C11*(1/RC[-6]^R6C11-1)))*R7C11+R2C11)"

Range("H" & 8 + i & "").Value = waga15(i)
Range("D" & 8 + i & "").Value = waga15(i)

Range("I" & 8 + i & "").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>0,((RC[-2]-RC[-3])/RC[-3])^2*RC[-1],0)"








Next i



Range("J" & 8 + i - 1 & "").Value = "=(F" & 8 + i - 1 & "-G" & 8 + i - 1 & ")/F" & 8 + i - 1 & ""






        ' ***************************DUONG*****************************************
    
    FseNew = 0
    FseOLD = 1

    
    Range("L3").Value = 1
    Range("L4").Value = 1
    Range("L5").Value = 1
    pzero = 0.000001
       
    Do Until FseOLD = FseNew
    
    FseOLD = Range("I5").Value
    
    
    
    SolverReset
    SolverAdd "L3", 3, pzero
    SolverAdd "L4", 3, pzero
    SolverAdd "L5", 3, pzero
    
    SolverOk SetCell:="$I$5", MaxMinVal:=2, ValueOf:=0, ByChange:="$L$3:$L$5", _
    Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOptions MaxTime:=1000, Iterations:=100, Precision:=0.000001, _
    Convergence:=0.000000000001, StepThru:=False, Scaling:=False, AssumeNonNeg:= _
    False, Derivatives:=2
    SolverFinish KeepFinal = True
    SolverSolve UserFinish = False, ShowRef:="ShowTrial"
    
    FseNew = Range("I5").Value
    
    Loop
    
    blad15 = -Range("J" & 8 + i - 1 & "").Value
   
For i = 0 To rozm - 1
   If (czas(i) >= P13 And czas(i) <= P33) Then
   Range("L" & 8 + i & "").Value = Range("C" & 8 + i & "").Value
   Range("M" & 8 + i & "").Value = Range("G" & 8 + i & "").Value
   End If
Next i

    
  Sheets("Wyniki").Select
  Range("E7").Value = blad15

'ActiveWorkbook.SaveAs Filename:="D:\_projekty\BLUEGAS\DOpasowania FOPT\DUONG_" & nazwa & ".xlsm", FileFormat:=52

Sheets("Wyniki").Select




ActiveWorkbook.Save


End Sub
