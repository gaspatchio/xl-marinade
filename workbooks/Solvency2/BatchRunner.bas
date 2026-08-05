Attribute VB_Name = "BatchRunner"
Option Explicit

' Computes the life SCR for every policy on Lifelib_Reference (t0 = 0,
' scenario 1) and reconciles all nine stored quantities against lifelib:
' the SCR, the four live risk charges, the three lapse sub-shock charges
' and the base NAV. Results land on Batch_Results (summary C4:C11, one
' row per policy from row 14). Never MsgBox (invisible modal in a
' hidden Excel) - errors are written to the sheet.

Sub RunAllPolicies()
    Dim wsRef As Worksheet, wsOut As Worksheet, wsSum As Worksheet
    Dim nRuns As Long, i As Long, q As Long, retries As Long
    Dim hasCalcState As Boolean
    Dim limitVal As Variant, tol As Double
    Dim refData As Variant
    Dim outData() As Variant
    Dim polRange As Range
    Dim wbVals(1 To 9) As Double, refVals(1 To 9) As Double
    Dim relDiff(1 To 9) As Double
    Dim maxRelSCR As Double, maxRelLife As Double, maxRelNAV As Double
    Dim nOK As Long
    Dim ok As Boolean
    Dim t0 As Double
    Dim savedPol As Variant, savedScen As Variant, savedT0 As Variant
    Dim savedCalc As XlCalculation

    Set wsRef = ThisWorkbook.Worksheets("Lifelib_Reference")
    Set wsOut = ThisWorkbook.Worksheets("Batch_Results")
    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set polRange = ThisWorkbook.Names("PolicyID").RefersToRange

    On Error Resume Next
    hasCalcState = (Application.CalculationState = xlDone) Or True
    hasCalcState = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    t0 = Timer
    savedPol = polRange.Value
    savedScen = ThisWorkbook.Names("ScenID").RefersToRange.Value
    savedT0 = ThisWorkbook.Names("T0").RefersToRange.Value
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ThisWorkbook.Names("ScenID").RefersToRange.Value = 1
    ThisWorkbook.Names("T0").RefersToRange.Value = 0

    nRuns = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row - 2
    limitVal = ThisWorkbook.Names("BatchLimit").RefersToRange.Value
    If IsNumeric(limitVal) And Not IsEmpty(limitVal) Then
        If limitVal > 0 And limitVal < nRuns Then nRuns = CLng(limitVal)
    End If
    tol = ThisWorkbook.Names("BatchTol").RefersToRange.Value

    ' reference: policy_id, scr_life, life_mort, life_longev, life_lapse,
    ' life_exps, lapse_up, lapse_down, lapse_mass, nav_base
    refData = wsRef.Range(wsRef.Cells(3, 1), wsRef.Cells(nRuns + 2, 10)).Value
    ReDim outData(1 To nRuns, 1 To 13)

    wsOut.Range(wsOut.Cells(14, 2), wsOut.Cells(wsOut.Rows.Count, 14)).ClearContents

    maxRelSCR = 0#
    maxRelLife = 0#
    maxRelNAV = 0#
    nOK = 0

    For i = 1 To nRuns
        polRange.Value = refData(i, 1)
        Application.Calculate
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
                "Calculation did not propagate for policy " & refData(i, 1)
        End If

        wbVals(1) = ThisWorkbook.Names("SCR_Life").RefersToRange.Value
        wbVals(2) = ThisWorkbook.Names("Life_Mort").RefersToRange.Value
        wbVals(3) = ThisWorkbook.Names("Life_Longev").RefersToRange.Value
        wbVals(4) = ThisWorkbook.Names("Life_Lapse").RefersToRange.Value
        wbVals(5) = ThisWorkbook.Names("LapseUp_Risk").RefersToRange.Value
        wbVals(6) = ThisWorkbook.Names("LapseDown_Risk").RefersToRange.Value
        wbVals(7) = ThisWorkbook.Names("LapseMass_Risk").RefersToRange.Value
        wbVals(8) = ThisWorkbook.Names("Life_Exps").RefersToRange.Value
        wbVals(9) = ThisWorkbook.Names("NAV_Base").RefersToRange.Value
        refVals(1) = refData(i, 2)
        refVals(2) = refData(i, 3)
        refVals(3) = refData(i, 4)
        refVals(4) = refData(i, 5)
        refVals(5) = refData(i, 7)
        refVals(6) = refData(i, 8)
        refVals(7) = refData(i, 9)
        refVals(8) = refData(i, 6)
        refVals(9) = refData(i, 10)

        ok = True
        For q = 1 To 9
            relDiff(q) = Abs(wbVals(q) - refVals(q)) _
                / WorksheetFunction.Max(1#, Abs(refVals(q)))
            If relDiff(q) > tol Then ok = False
        Next q
        If relDiff(1) > maxRelSCR Then maxRelSCR = relDiff(1)
        For q = 2 To 8
            If relDiff(q) > maxRelLife Then maxRelLife = relDiff(q)
        Next q
        If relDiff(9) > maxRelNAV Then maxRelNAV = relDiff(9)
        If ok Then nOK = nOK + 1

        outData(i, 1) = refData(i, 1)
        outData(i, 2) = wbVals(1)
        outData(i, 3) = refVals(1)
        For q = 1 To 9
            outData(i, 3 + q) = relDiff(q)
        Next q
        outData(i, 13) = IIf(ok, "TRUE", "FALSE")
    Next i

    wsOut.Range(wsOut.Cells(14, 2), wsOut.Cells(13 + nRuns, 14)).Value = outData

    wsOut.Range("C4").Value = nRuns
    wsOut.Range("C5").Value = nOK
    wsOut.Range("C6").Value = nRuns - nOK
    wsOut.Range("C7").Value = maxRelSCR
    wsOut.Range("C8").Value = maxRelLife
    wsOut.Range("C9").Value = maxRelNAV
    wsOut.Range("C10").Value = Round(Timer - t0, 1)
    wsOut.Range("C11").Value = IIf(nOK = nRuns, "PASS", "FAIL")
    wsOut.Range("C11").Font.Bold = True

CleanUp:
    polRange.Value = savedPol
    ThisWorkbook.Names("ScenID").RefersToRange.Value = savedScen
    ThisWorkbook.Names("T0").RefersToRange.Value = savedT0
    Application.Calculation = savedCalc
    Application.Calculate
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        wsOut.Range("C11").Value = "ERROR"
        wsOut.Range("D11").Value = Err.Description
        wsOut.Range("C11").Font.Bold = True
    End If
End Sub
