Attribute VB_Name = "BatchRunner"
Option Explicit

' Runs the projection for every model point on Lifelib_Reference and
' reconciles ALL six stored lifelib quantities: office premium and the
' PVs of premiums, claims, expenses, commissions and net cash flow.
' Results land on Batch_Results:
'   - summary block in C4:C10 (points run, matches, max diffs, verdict)
'   - one row per model point from row 13 down
'
' Control sheet settings:
'   BatchLimit - cap on points to run (blank = all)
'   BatchTol   - relative tolerance: reconciles when
'                |wb - lifelib| <= tol * Max(1, |lifelib|)

Sub RunAllModelPoints()
    Dim wsRef As Worksheet, wsOut As Worksheet, wsSum As Worksheet
    Dim nPoints As Long, i As Long, q As Long, retries As Long
    Dim hasCalcState As Boolean
    Dim limitVal As Variant, tol As Double
    Dim refData As Variant
    Dim outData() As Variant
    Dim pointRange As Range
    Dim wbVals(1 To 5) As Double
    Dim relPV(1 To 5) As Double
    Dim premWB As Double, premRef As Double
    Dim diffPrem As Double
    Dim maxDiffPrem As Double, maxRelPV As Double
    Dim nOK As Long
    Dim ok As Boolean
    Dim t0 As Double
    Dim savedPoint As Variant, savedCalc As XlCalculation

    Set wsRef = ThisWorkbook.Worksheets("Lifelib_Reference")
    Set wsOut = ThisWorkbook.Worksheets("Batch_Results")
    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set pointRange = ThisWorkbook.Names("PointID").RefersToRange

    ' Probe once whether this Excel exposes Application.CalculationState
    On Error Resume Next
    hasCalcState = (Application.CalculationState = xlDone) Or True
    hasCalcState = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    t0 = Timer
    savedPoint = pointRange.Value
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' reference table: point_id, premium_pp, pv_premiums, pv_claims,
    ' pv_expenses, pv_commissions, pv_net_cf (headers on row 2)
    nPoints = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row - 2
    limitVal = ThisWorkbook.Names("BatchLimit").RefersToRange.Value
    If IsNumeric(limitVal) And Not IsEmpty(limitVal) Then
        If limitVal > 0 And limitVal < nPoints Then nPoints = CLng(limitVal)
    End If
    tol = ThisWorkbook.Names("BatchTol").RefersToRange.Value

    refData = wsRef.Range(wsRef.Cells(3, 1), wsRef.Cells(nPoints + 2, 7)).Value
    ReDim outData(1 To nPoints, 1 To 12)

    ' clear previous run
    wsOut.Range(wsOut.Cells(13, 2), wsOut.Cells(wsOut.Rows.Count, 13)).ClearContents

    maxDiffPrem = 0#
    maxRelPV = 0#
    nOK = 0

    For i = 1 To nPoints
        pointRange.Value = refData(i, 1)
        Application.Calculate

        ' Calculate can return before the chain is done under sustained VBA
        ' looping (observed: stale reads of the previous point's results).
        ' Wait for the engine, then confirm via the Summary sentinel cell
        ' (=PointID) that this iteration's point has propagated.
        If hasCalcState Then
            Do While Application.CalculationState <> xlDone
                DoEvents
            Loop
        End If
        retries = 0
        Do While wsSum.Range("C4").Value <> refData(i, 1) And retries < 20
            DoEvents
            Application.Calculate
            If hasCalcState Then
                Do While Application.CalculationState <> xlDone
                    DoEvents
                Loop
            End If
            retries = retries + 1
        Loop
        If wsSum.Range("C4").Value <> refData(i, 1) Then
            Err.Raise vbObjectError + 1, , _
                "Calculation did not propagate for model point " & refData(i, 1)
        End If

        premWB = ThisWorkbook.Names("PremiumPP").RefersToRange.Value
        premRef = refData(i, 2)
        wbVals(1) = ThisWorkbook.Names("PV_Premiums").RefersToRange.Value
        wbVals(2) = ThisWorkbook.Names("PV_Claims").RefersToRange.Value
        wbVals(3) = ThisWorkbook.Names("PV_Expenses").RefersToRange.Value
        wbVals(4) = ThisWorkbook.Names("PV_Commissions").RefersToRange.Value
        wbVals(5) = ThisWorkbook.Names("PV_NetCF").RefersToRange.Value

        diffPrem = Abs(premWB - premRef)
        If diffPrem > maxDiffPrem Then maxDiffPrem = diffPrem
        ok = (diffPrem <= tol * WorksheetFunction.Max(1#, Abs(premRef)))
        For q = 1 To 5
            relPV(q) = Abs(wbVals(q) - refData(i, 2 + q)) _
                / WorksheetFunction.Max(1#, Abs(refData(i, 2 + q)))
            If relPV(q) > maxRelPV Then maxRelPV = relPV(q)
            If relPV(q) > tol Then ok = False
        Next q
        If ok Then nOK = nOK + 1

        outData(i, 1) = refData(i, 1)
        outData(i, 2) = premWB
        outData(i, 3) = premRef
        outData(i, 4) = diffPrem
        outData(i, 5) = relPV(1)
        outData(i, 6) = relPV(2)
        outData(i, 7) = relPV(3)
        outData(i, 8) = relPV(4)
        outData(i, 9) = wbVals(5)
        outData(i, 10) = refData(i, 7)
        outData(i, 11) = relPV(5)
        outData(i, 12) = IIf(ok, "TRUE", "FALSE")
    Next i

    wsOut.Range(wsOut.Cells(13, 2), wsOut.Cells(12 + nPoints, 13)).Value = outData

    wsOut.Range("C4").Value = nPoints
    wsOut.Range("C5").Value = nOK
    wsOut.Range("C6").Value = nPoints - nOK
    wsOut.Range("C7").Value = maxDiffPrem
    wsOut.Range("C8").Value = maxRelPV
    wsOut.Range("C9").Value = Round(Timer - t0, 1)
    wsOut.Range("C10").Value = IIf(nOK = nPoints, "PASS", "FAIL")
    wsOut.Range("C10").Font.Bold = True

CleanUp:
    pointRange.Value = savedPoint
    Application.Calculation = savedCalc
    Application.Calculate
    Application.ScreenUpdating = True
    ' Never MsgBox here: in an unattended (hidden) Excel a dialog is an
    ' invisible modal that blocks all Apple events. Report on the sheet.
    If Err.Number <> 0 Then
        wsOut.Range("C10").Value = "ERROR"
        wsOut.Range("D10").Value = Err.Description
        wsOut.Range("C10").Font.Bold = True
    End If
End Sub
