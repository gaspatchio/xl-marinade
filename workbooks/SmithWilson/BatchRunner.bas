Attribute VB_Name = "BatchRunner"
Option Explicit

' Recalculates the workbook and reconciles the full calibration and
' extrapolation results - the zeta vector and the P and R curves -
' against the lifelib values on Lifelib_Reference. Results land on
' Batch_Results (summary C4:C11, one row per value from row 14).
' Never MsgBox (invisible modal in a hidden Excel) - errors are
' written to the sheet.

Sub RunCurveReconciliation()
    Dim wsOut As Worksheet
    Dim hasCalcState As Boolean
    Dim tol As Double
    Dim zetaWB As Variant, zetaRef As Variant
    Dim pWB As Variant, pRef As Variant
    Dim rWB As Variant, rRef As Variant
    Dim outData() As Variant
    Dim i As Long, nRows As Long, nOK As Long, nTot As Long
    Dim absDiff As Double, relDiff As Double
    Dim maxAbsZeta As Double, maxRelP As Double, maxRelR As Double
    Dim ok As Boolean
    Dim t0 As Double
    Dim savedCalc As XlCalculation

    Set wsOut = ThisWorkbook.Worksheets("Batch_Results")

    On Error Resume Next
    hasCalcState = (Application.CalculationState = xlDone) Or True
    hasCalcState = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    t0 = Timer
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    tol = ThisWorkbook.Names("BatchTol").RefersToRange.Value

    Application.Calculate
    If hasCalcState Then
        Do While Application.CalculationState <> xlDone
            DoEvents
        Loop
    End If

    zetaWB = ThisWorkbook.Names("ZetaVec").RefersToRange.Value
    zetaRef = ThisWorkbook.Names("Ref_Zeta").RefersToRange.Value
    pWB = ThisWorkbook.Names("Ext_P").RefersToRange.Value
    pRef = ThisWorkbook.Names("Ref_P").RefersToRange.Value
    rWB = ThisWorkbook.Names("Ext_R").RefersToRange.Value
    rRef = ThisWorkbook.Names("Ref_R").RefersToRange.Value

    nTot = UBound(zetaWB, 1) + UBound(pWB, 1) + UBound(rWB, 1)
    ReDim outData(1 To nTot, 1 To 7)
    wsOut.Range(wsOut.Cells(14, 2), wsOut.Cells(wsOut.Rows.Count, 8)).ClearContents

    nRows = 0
    nOK = 0
    maxAbsZeta = 0#
    maxRelP = 0#
    maxRelR = 0#

    For i = 1 To UBound(zetaWB, 1)
        absDiff = Abs(zetaWB(i, 1) - zetaRef(i, 1))
        relDiff = absDiff / WorksheetFunction.Max(1#, Abs(zetaRef(i, 1)))
        If absDiff > maxAbsZeta Then maxAbsZeta = absDiff
        ok = (relDiff <= tol)
        If ok Then nOK = nOK + 1
        nRows = nRows + 1
        outData(nRows, 1) = "zeta"
        outData(nRows, 2) = i
        outData(nRows, 3) = zetaWB(i, 1)
        outData(nRows, 4) = zetaRef(i, 1)
        outData(nRows, 5) = absDiff
        outData(nRows, 6) = relDiff
        outData(nRows, 7) = IIf(ok, "TRUE", "FALSE")
    Next i

    For i = 1 To UBound(pWB, 1)
        absDiff = Abs(pWB(i, 1) - pRef(i, 1))
        relDiff = absDiff / WorksheetFunction.Max(1#, Abs(pRef(i, 1)))
        If relDiff > maxRelP Then maxRelP = relDiff
        ok = (relDiff <= tol)
        If ok Then nOK = nOK + 1
        nRows = nRows + 1
        outData(nRows, 1) = "P"
        outData(nRows, 2) = i
        outData(nRows, 3) = pWB(i, 1)
        outData(nRows, 4) = pRef(i, 1)
        outData(nRows, 5) = absDiff
        outData(nRows, 6) = relDiff
        outData(nRows, 7) = IIf(ok, "TRUE", "FALSE")
    Next i

    For i = 1 To UBound(rWB, 1)
        absDiff = Abs(rWB(i, 1) - rRef(i, 1))
        relDiff = absDiff / WorksheetFunction.Max(1#, Abs(rRef(i, 1)))
        If relDiff > maxRelR Then maxRelR = relDiff
        ok = (relDiff <= tol)
        If ok Then nOK = nOK + 1
        nRows = nRows + 1
        outData(nRows, 1) = "R"
        outData(nRows, 2) = i
        outData(nRows, 3) = rWB(i, 1)
        outData(nRows, 4) = rRef(i, 1)
        outData(nRows, 5) = absDiff
        outData(nRows, 6) = relDiff
        outData(nRows, 7) = IIf(ok, "TRUE", "FALSE")
    Next i

    wsOut.Range(wsOut.Cells(14, 2), wsOut.Cells(13 + nTot, 8)).Value = outData

    wsOut.Range("C4").Value = nTot
    wsOut.Range("C5").Value = nOK
    wsOut.Range("C6").Value = nTot - nOK
    wsOut.Range("C7").Value = maxAbsZeta
    wsOut.Range("C8").Value = maxRelP
    wsOut.Range("C9").Value = maxRelR
    wsOut.Range("C10").Value = Round(Timer - t0, 1)
    wsOut.Range("C11").Value = IIf(nOK = nTot, "PASS", "FAIL")
    wsOut.Range("C11").Font.Bold = True

CleanUp:
    Application.Calculation = savedCalc
    Application.Calculate
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        wsOut.Range("C11").Value = "ERROR"
        wsOut.Range("D11").Value = Err.Description
        wsOut.Range("C11").Font.Bold = True
    End If
End Sub
