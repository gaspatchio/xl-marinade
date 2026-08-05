Attribute VB_Name = "BatchRunner"
Option Explicit

' Runs the projection for every (model point, scenario) row on
' Lifelib_Reference and reconciles ALL seven stored lifelib PVs:
' premiums, claims, expenses, commissions, investment income, change
' in account value and net cash flow. Results land on Batch_Results
' (summary C4:C10, one row per run from row 13). Never MsgBox
' (invisible modal in a hidden Excel) - errors are written to the sheet.

Sub RunAllModelPoints()
    Dim wsRef As Worksheet, wsOut As Worksheet, wsSum As Worksheet
    Dim nRuns As Long, i As Long, retries As Long
    Dim hasCalcState As Boolean
    Dim limitVal As Variant, tol As Double
    Dim refData As Variant
    Dim outData() As Variant
    Dim pointRange As Range, scenRange As Range
    Dim wbVals(1 To 7) As Double
    Dim relPV(1 To 7) As Double
    Dim q As Long
    Dim maxRelPrem As Double, maxRelPV As Double
    Dim nOK As Long
    Dim ok As Boolean
    Dim t0 As Double
    Dim savedPoint As Variant, savedScen As Variant, savedCalc As XlCalculation

    Set wsRef = ThisWorkbook.Worksheets("Lifelib_Reference")
    Set wsOut = ThisWorkbook.Worksheets("Batch_Results")
    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set pointRange = ThisWorkbook.Names("PointID").RefersToRange
    Set scenRange = ThisWorkbook.Names("ScenID").RefersToRange

    On Error Resume Next
    hasCalcState = (Application.CalculationState = xlDone) Or True
    hasCalcState = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    t0 = Timer
    savedPoint = pointRange.Value
    savedScen = scenRange.Value
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    nRuns = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row - 2
    limitVal = ThisWorkbook.Names("BatchLimit").RefersToRange.Value
    If IsNumeric(limitVal) And Not IsEmpty(limitVal) Then
        If limitVal > 0 And limitVal < nRuns Then nRuns = CLng(limitVal)
    End If
    tol = ThisWorkbook.Names("BatchTol").RefersToRange.Value

    ' reference: point_id, scen_id, pv_premiums, pv_claims, pv_expenses,
    ' pv_commissions, pv_inv_income, pv_av_change, pv_net_cf (cols 3-9)
    refData = wsRef.Range(wsRef.Cells(3, 1), wsRef.Cells(nRuns + 2, 9)).Value
    ReDim outData(1 To nRuns, 1 To 12)

    wsOut.Range(wsOut.Cells(13, 2), wsOut.Cells(wsOut.Rows.Count, 13)).ClearContents

    maxRelPrem = 0#
    maxRelPV = 0#
    nOK = 0

    For i = 1 To nRuns
        pointRange.Value = refData(i, 1)
        scenRange.Value = refData(i, 2)
        Application.Calculate
        If hasCalcState Then
            Do While Application.CalculationState <> xlDone
                DoEvents
            Loop
        End If
        retries = 0
        Do While (wsSum.Range("C4").Value <> refData(i, 1) _
                  Or wsSum.Range("C5").Value <> refData(i, 2)) And retries < 20
            DoEvents
            Application.Calculate
            If hasCalcState Then
                Do While Application.CalculationState <> xlDone
                    DoEvents
                Loop
            End If
            retries = retries + 1
        Loop
        If wsSum.Range("C4").Value <> refData(i, 1) _
           Or wsSum.Range("C5").Value <> refData(i, 2) Then
            Err.Raise vbObjectError + 1, , _
                "Calculation did not propagate for run " & refData(i, 1) & "/" & refData(i, 2)
        End If

        wbVals(1) = ThisWorkbook.Names("PV_Premiums").RefersToRange.Value
        wbVals(2) = ThisWorkbook.Names("PV_Claims").RefersToRange.Value
        wbVals(3) = ThisWorkbook.Names("PV_Expenses").RefersToRange.Value
        wbVals(4) = ThisWorkbook.Names("PV_Commissions").RefersToRange.Value
        wbVals(5) = ThisWorkbook.Names("PV_InvIncome").RefersToRange.Value
        wbVals(6) = ThisWorkbook.Names("PV_AvChange").RefersToRange.Value
        wbVals(7) = ThisWorkbook.Names("PV_NetCF").RefersToRange.Value

        ok = True
        For q = 1 To 7
            relPV(q) = Abs(wbVals(q) - refData(i, 2 + q)) _
                / WorksheetFunction.Max(1#, Abs(refData(i, 2 + q)))
            If relPV(q) > tol Then ok = False
            If q >= 2 And relPV(q) > maxRelPV Then maxRelPV = relPV(q)
        Next q
        If relPV(1) > maxRelPrem Then maxRelPrem = relPV(1)
        If relPV(1) > maxRelPV Then maxRelPV = relPV(1)
        If ok Then nOK = nOK + 1

        outData(i, 1) = refData(i, 1)
        outData(i, 2) = refData(i, 2)
        outData(i, 3) = wbVals(7)
        outData(i, 4) = refData(i, 9)
        For q = 1 To 7
            outData(i, 4 + q) = relPV(q)
        Next q
        outData(i, 12) = IIf(ok, "TRUE", "FALSE")
    Next i

    wsOut.Range(wsOut.Cells(13, 2), wsOut.Cells(12 + nRuns, 13)).Value = outData

    wsOut.Range("C4").Value = nRuns
    wsOut.Range("C5").Value = nOK
    wsOut.Range("C6").Value = nRuns - nOK
    wsOut.Range("C7").Value = maxRelPrem
    wsOut.Range("C8").Value = maxRelPV
    wsOut.Range("C9").Value = Round(Timer - t0, 1)
    wsOut.Range("C10").Value = IIf(nOK = nRuns, "PASS", "FAIL")
    wsOut.Range("C10").Font.Bold = True

CleanUp:
    pointRange.Value = savedPoint
    scenRange.Value = savedScen
    Application.Calculation = savedCalc
    Application.Calculate
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        wsOut.Range("C10").Value = "ERROR"
        wsOut.Range("D10").Value = Err.Description
        wsOut.Range("C10").Font.Bold = True
    End If
End Sub
