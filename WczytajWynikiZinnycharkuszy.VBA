Sub wczytaj()

For i = 3 To 102
'For i = 3 To 3

sym = Range("A" & i & "").Value
'hiperboliczna
Workbooks.Open Filename:="D:\_Projekty\BLUEGAS\szumy\Statystyka\DUONG_NBLUK00_ERR0_0.01_" & sym & ".xlsm"
Sheets("wyniki").Select

ht = Range("D2").Value
h105 = Range("E4").Value
h21 = Range("E5").Value
h41 = Range("E6").Value
h51 = Range("E7").Value

ThisWorkbook.Activate

Sheets("zestawienie").Select
Range("B" & i & "").Value = ht
Range("C" & i & "").Value = h105
Range("D" & i & "").Value = h21
Range("E" & i & "").Value = h41
Range("F" & i & "").Value = h51

Workbooks("DUONG_NBLUK00_ERR0_0.01_" & sym & ".xlsm").Save
Workbooks("DUONG_NBLUK00_ERR0_0.01_" & sym & ".xlsm").Close

Workbooks.Open Filename:="D:\_Projekty\BLUEGAS\szumy\Statystyka\DUONG_NBLUK00_ERR0_0.03_" & sym & ".xlsm"
Sheets("wyniki").Select

ht = Range("D2").Value
h105 = Range("E4").Value
h21 = Range("E5").Value
h41 = Range("E6").Value
h51 = Range("E7").Value

ThisWorkbook.Activate

Sheets("zestawienie").Select
Range("G" & i & "").Value = ht
Range("H" & i & "").Value = h105
Range("I" & i & "").Value = h21
Range("J" & i & "").Value = h41
Range("K" & i & "").Value = h51

Workbooks("DUONG_NBLUK00_ERR0_0.03_" & sym & ".xlsm").Save
Workbooks("DUONG_NBLUK00_ERR0_0.03_" & sym & ".xlsm").Close

Workbooks.Open Filename:="D:\_Projekty\BLUEGAS\szumy\Statystyka\DUONG_NBLUK00_ERR0_0.05_" & sym & ".xlsm"
Sheets("wyniki").Select

ht = Range("D2").Value
h105 = Range("E4").Value
h21 = Range("E5").Value
h41 = Range("E6").Value
h51 = Range("E7").Value

ThisWorkbook.Activate

Sheets("zestawienie").Select
Range("L" & i & "").Value = ht
Range("M" & i & "").Value = h105
Range("N" & i & "").Value = h21
Range("O" & i & "").Value = h41
Range("P" & i & "").Value = h51

Workbooks("DUONG_NBLUK00_ERR0_0.05_" & sym & ".xlsm").Save
Workbooks("DUONG_NBLUK00_ERR0_0.05_" & sym & ".xlsm").Close

Workbooks.Open Filename:="D:\_Projekty\BLUEGAS\szumy\Statystyka\DUONG_NBLUK00_ERR0_0.1_" & sym & ".xlsm"
Sheets("wyniki").Select

ht = Range("D2").Value
h105 = Range("E4").Value
h21 = Range("E5").Value
h41 = Range("E6").Value
h51 = Range("E7").Value

ThisWorkbook.Activate

Sheets("zestawienie").Select
Range("Q" & i & "").Value = ht
Range("R" & i & "").Value = h105
Range("S" & i & "").Value = h21
Range("T" & i & "").Value = h41
Range("U" & i & "").Value = h51

Workbooks("DUONG_NBLUK00_ERR0_0.1_" & sym & ".xlsm").Save
Workbooks("DUONG_NBLUK00_ERR0_0.1_" & sym & ".xlsm").Close


Next i




End Sub
