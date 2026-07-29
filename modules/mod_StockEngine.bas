Attribute VB_Name = "mod_StockEngine"
' ============================================================================
' Academix v13.2 - DSS Logistique El Bayadh
' Copyright (c) 2025-2026 Mahi Kamel Abdelghani
' ============================================================================

Option Explicit

Public Type ArticleDetails
    Code        As String
    Designation As String
    Category    As String
    PU          As Double
    Stock       As Long
End Type

' ... (existing code)


' ================================================================================
' CONSTANTS ? Synchronized with Unitï¿½ de traitement VBA GROUND_TRUTH
' ================================================================================
Private Const ORDER_COST_S  As Double = 801.45 ' DZD ? full order cycle cost (field-refined from 500)
Private Const HOLDING_RATE  As Double = 0.2    ' 20% of unit price per year
Private Const LEAD_TIME_DEFAULT As Integer = 2 ' Default delivery days

' Article-specific safety stocks ? mirrors UnitÃ© de traitement VBA GROUND_TRUTH (Calibrated v13.4 from v7 historical)
Public Function GetSafetyStock(ByVal sku As String) As Double
    Select Case UCase(Trim(sku))
        Case "ART-001": GetSafetyStock = 400  ' Papier A4 (v7 Stock Min)
        Case "ART-002": GetSafetyStock = 200  ' Toner G030 (case study value)
        Case "ART-003": GetSafetyStock = 30   ' Papier A3 (v7 Stock Min)
        Case "ART-004": GetSafetyStock = 20   ' BoÃ®te archives (v7 Stock Min)
        Case "ART-005": GetSafetyStock = 2    ' Agrafeuse (v7 Stock Min)
        Case "ART-006": GetSafetyStock = 5    ' Stylos (v7 Stock Min)
        Case "ART-007": GetSafetyStock = 2    ' Registre 5m (v7 Stock Min)
        Case "ART-008": GetSafetyStock = 2    ' Encre tampon (v7 Stock Min)
        Case "ART-009": GetSafetyStock = 10   ' Sous-chemise (v7 Stock Min)
        Case "ART-010": GetSafetyStock = 5    ' Chemise (v7 Stock Min)
        Case "ART-011": GetSafetyStock = 1    ' Fax (v7 Stock Min)
        Case "ART-012": GetSafetyStock = 5    ' Marqueur (v7 Stock Min)
        Case "ART-013": GetSafetyStock = 10   ' Encre cachets (v7 Stock Min)
        Case "ART-014": GetSafetyStock = 5    ' Classeur (v7 Stock Min)
        Case "ART-015": GetSafetyStock = 39   ' Toner gÃ©nÃ©rique (v7 Stock Min)
        Case Else:      GetSafetyStock = 50
    End Select
End Function

' ================================================================================
' FUNCTION: ComputeEOQ
' Formula: Q* = SQRT(2 x D x S / (P x t))
' ================================================================================
Public Function ComputeEOQ(ByVal AnnualDemand As Double, _
                            ByVal unitPrice As Double) As Double
    If unitPrice <= 0 Or AnnualDemand <= 0 Then
        ComputeEOQ = 0
        Exit Function
    End If

    Dim holdingCostH As Double
    holdingCostH = unitPrice * HOLDING_RATE

    ComputeEOQ = Sqr((2 * AnnualDemand * ORDER_COST_S) / holdingCostH)
End Function

' ================================================================================
' FUNCTION: ComputeROP
' Formula: ROP = (avg_daily_demand x lead_time) + safety_stock
' Added: Session 22 - was referenced by mod_StockEngine.ValidateStockLevel,
'       mod_Dashboard (2 sites), mod_UI_Setup.DrawStockoutBanner, mod_Procurement
'       but the function itself was missing. Signature mirrors external
'       reference (external/lsm-vba-core) so 2-arg callsites work via the
'       Optional default on LeadTimeDays.
' ================================================================================
Public Function ComputeROP(ByVal AvgDailyDemand As Double, _
                            ByVal sku As String, _
                            Optional ByVal LeadTimeDays As Integer = LEAD_TIME_DEFAULT) As Double
    ComputeROP = (AvgDailyDemand * LeadTimeDays) + GetSafetyStock(sku)
End Function

' ================================================================================
' FUNCTION: GetServices
' Returns a list of services from the SYS_STRINGS sheet.
' ================================================================================
Public Function GetServices() As Collection
    Dim services As New Collection
    Dim wsStr As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set wsStr = ThisWorkbook.Sheets(mod_Config.SHEET_SYS_STRINGS)
    On Error GoTo 0
    
    If Not wsStr Is Nothing Then
        lastRow = wsStr.Cells(wsStr.Rows.count, COL_SYS_ID).End(xlUp).Row
        For i = 2 To lastRow
            If Left(Trim(CStr(wsStr.Cells(i, COL_SYS_ID).Value)), 4) = "SVC_" Then
                services.Add Trim(CStr(wsStr.Cells(i, COL_SYS_VALUE).Value))
            End If
        Next i
    End If
    
    If services.count = 0 Then
        services.Add "Service 1"
        services.Add "Service 2"
        services.Add "Fournisseur Externe"
    End If
    
    Set GetServices = services
End Function

' ================================================================================
' FUNCTION: GetArticles
' Returns a list of articles (Code | Designation) from the ARTICLES sheet.
' ================================================================================
Public Function GetArticles(ByVal filterCat As String) As Collection
    Dim articles As New Collection
    Dim wsArt As Worksheet
    Dim lastRow As Long, i As Long
    Dim code As String, desig As String, cat As String
    Dim noFilter As Boolean

    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0

    If wsArt Is Nothing Then
        articles.Add "ART-001 | Papier A4 80g/mÂ²"
        articles.Add "ART-002 | Toner G030 (noir)"
        articles.Add "ART-003 | Papier A3 80g/mÂ²"
        articles.Add "ART-004 | BoÃ®te archives carton"
        articles.Add "ART-005 | Agrafeuse de bureau"
        articles.Add "ART-006 | Stylos bille boÃ®te/50"
        articles.Add "ART-007 | Registre grand format 5m"
        articles.Add "ART-008 | Encre tampon"
        articles.Add "ART-009 | Sous-chemise carton"
        articles.Add "ART-010 | Chemise cartonnÃ©e"
        articles.Add "ART-011 | Rouleau papier fax"
        articles.Add "ART-012 | Marqueur permanent noir"
        articles.Add "ART-013 | Encre pour cachets"
        articles.Add "ART-014 | Classeur Ã  levier"
        articles.Add "ART-015 | Cartouche toner gÃ©nÃ©rique"
        Set GetArticles = articles
        Exit Function
    End If

    noFilter = (filterCat = "" Or filterCat = "(Toutes)")
    lastRow = wsArt.Cells(wsArt.Rows.count, COL_ART_CODE).End(xlUp).Row

    For i = 3 To lastRow
        code = Trim(CStr(wsArt.Cells(i, COL_ART_CODE).Value))
        desig = Trim(CStr(wsArt.Cells(i, COL_ART_DESIGNATION).Value))
        cat = Trim(CStr(wsArt.Cells(i, COL_ART_CATEGORIE).Value))

        If code <> "" Then
            If noFilter Or (cat = filterCat) Then
                articles.Add code & " | " & desig
            End If
        End If
    Next i

    Set GetArticles = articles
End Function


' ================================================================================
' SUB: ValidateStockLevel
' Fires a UI alert if current stock breaches ROP.
' ================================================================================
Public Sub ValidateStockLevel(ByVal sku As String, _
                               ByVal CurrentStock As Double, _
                               ByVal AnnualDemand As Double, _
                               ByVal unitPrice As Double)
    If AnnualDemand <= 0 Then Exit Sub

    Dim avgDaily As Double: avgDaily = AnnualDemand / mod_Config.WORKING_DAYS_PER_YEAR
    Dim rop As Double: rop = ComputeROP(avgDaily, sku)
    Dim ss As Double: ss = GetSafetyStock(sku)

    If CurrentStock <= rop Then
        Dim eoq As Double: eoq = ComputeEOQ(AnnualDemand, unitPrice)
        Dim alertLevel As String: alertLevel = IIf(CurrentStock <= ss, "RUPTURE IMMINENTE", "SEUIL D'ALERTE ATTEINT")

        MsgBox alertLevel & vbCrLf & vbCrLf & _
               "Article  : " & sku & vbCrLf & _
               "Stock    : " & CurrentStock & " unites" & vbCrLf & _
               "ROP      : " & Round(rop, 1) & " unites" & vbCrLf & _
               "SS       : " & ss & " unites" & vbCrLf & _
               "EOQ (Q*) : " & Round(eoq, 0) & " unites a commander", _
               vbExclamation, mod_Config.SYS_TITLE
    End If
End Sub

' ================================================================================
' FUNCTION: GetArticleDetails
' Returns a complete ArticleDetails struct for a given SKU.
' ================================================================================
Public Function GetArticleDetails(ByVal sku As String) As ArticleDetails
    Dim details As ArticleDetails
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        details.Code = sku
        GetArticleDetails = details
        Exit Function
    End If
    
    foundRow = Application.Match(sku, wsArt.Columns(COL_ART_CODE), 0)
    
    If IsError(foundRow) Then
        details.Code = sku
        details.Stock = 0
        GetArticleDetails = details
        Exit Function
    End If
    
    details.Code = Trim(CStr(wsArt.Cells(foundRow, COL_ART_CODE).Value))
    details.Designation = Trim(CStr(wsArt.Cells(foundRow, COL_ART_DESIGNATION).Value))
    details.Category = Trim(CStr(wsArt.Cells(foundRow, COL_ART_CATEGORIE).Value))
    details.PU = CDbl(mod_Utilities.SafeVal(wsArt.Cells(foundRow, COL_ART_PU).Value))
    details.Stock = CLng(mod_Utilities.SafeVal(wsArt.Cells(foundRow, COL_ART_STOCK).Value))
    
    GetArticleDetails = details
End Function

' ================================================================================
' SUB: UpdateArticleStockBalance
' Directly updates the stock quantity in the ARTICLES sheet based on movements.
' ================================================================================
Public Sub UpdateArticleStockBalance(ByVal artCode As String, ByVal mvtSign As String, ByVal qty As Long)
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then Exit Sub
    
    foundRow = Application.Match(artCode, wsArt.Columns(COL_ART_CODE), 0)
    
    If Not IsError(foundRow) Then
        wsArt.Unprotect Password:=mod_Config.MASTER_PWD
        
        Dim currentQty As Double: currentQty = Val(wsArt.Cells(foundRow, COL_ART_STOCK).Value) ' Column C: Stock
        
        If mvtSign = "IN" Then
            wsArt.Cells(foundRow, COL_ART_STOCK).Value = currentQty + qty
        Else
            wsArt.Cells(foundRow, COL_ART_STOCK).Value = currentQty - qty
        End If
        
        wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End If
End Sub

' ================================================================================
' FUNCTION: GetAnnualDemandFromHistory
' Aggregates annual demand from MOUVEMENTS sheet for a given SKU.
' ================================================================================
Public Function GetAnnualDemandFromHistory(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then GetAnnualDemandFromHistory = 0: Exit Function
    Dim currentYear As Integer: currentYear = Year(Date)
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    GetAnnualDemandFromHistory = WorksheetFunction.SumIfs( _
        wsMouv.Columns(COL_MOUV_QTE), _
        wsMouv.Columns(COL_MOUV_CODE_ARTICLE), sku, _
        wsMouv.Columns(COL_MOUV_TYPE), "OUT", _
        wsMouv.Columns(COL_MOUV_DATE), ">=" & DateSerial(currentYear, 1, 1))
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: CalculateCMUP
' Chronological moving weighted average.
'
' Basis: SCF, arrete du 26/07/2008, points 123-6 / 123-7 (aligne IAS 2), and the
' TAG1801 definition (SEMESTRE III, p.33): the average is taken over the stock
' HELD AT THE DATE OF ISSUE - "en divisant le cout total des stocks par le nombre
' de stocks en stock a la date de la sortie". Opening stock must therefore be
' carried forward.
'
' The previous implementation divided total IN value by total IN quantity. That
' ignores opening stock entirely and is correct only when opening stock is zero.
'
' Pass 1 derives the opening quantity by backing the movements out of the live
' balance:  initialQty = currentBalance - totalIn + totalOut
' Pass 2 walks the movements in DATE order, recomputing the average after each.
'
' Note: this domain defines no TVA parameters (there is no TAX_RATE or
' PU_INCLUDES_TVA in mod_Config), so unit costs are taken as recorded - the price
' applied per unit is unchanged from the previous implementation.
'
' EXPECT REPORTED VALUATIONS TO MOVE. The change is not limited to opening stock.
' The previous method was a lifetime average of all inbound, so it also ignored
' the effect of issues; this one averages over the stock actually held. The two
' agree only when opening stock is zero AND no issue has occurred. Where they
' differ, the previous figure was generally an OVERSTATEMENT when opening stock
' was cheaper than later receipts - which is the inventory-valuation exposure the
' correction removes.
' ================================================================================
Public Function CalculateCMUP(ByVal sku As String) As Double
    On Error GoTo ErrorHandler

    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    If wsMouv Is Nothing Or wsArt Is Nothing Then CalculateCMUP = 0: Exit Function

    ' --- Locate the article to read its live balance and unit price ---
    Dim foundRow As Variant
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(foundRow) Then CalculateCMUP = 0: Exit Function

    Dim currentBalance As Double: currentBalance = Val(wsArt.Cells(foundRow, COL_ART_STOCK).Value)
    Dim unitPrice As Double: unitPrice = Val(wsArt.Cells(foundRow, COL_ART_PU).Value)

    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD

    Dim lastRow As Long
    lastRow = wsMouv.Cells(wsMouv.Rows.Count, COL_MOUV_CODE_ARTICLE).End(xlUp).Row

    ' --- Pass 1: net the movements so the opening stock can be backed out ---
    Dim i As Long, mvType As String
    Dim totalInQty As Double, totalOutQty As Double
    For i = 2 To lastRow
        If Trim(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            mvType = UCase(Trim(wsMouv.Cells(i, COL_MOUV_TYPE).Value))
            If mvType = "IN" Then
                totalInQty = totalInQty + Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)
            ElseIf mvType = "OUT" Then
                totalOutQty = totalOutQty + Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)
            End If
        End If
    Next i

    Dim initialQty As Double: initialQty = currentBalance - totalInQty + totalOutQty
    If initialQty < 0 Then initialQty = 0

    ' --- Collect this article's movements ---
    Dim mvDate() As Double, mvTypeArr() As String
    Dim mvQty() As Double, mvVal() As Double, mvPU() As Double
    Dim n As Long: n = 0
    For i = 2 To lastRow
        If Trim(wsMouv.Cells(i, COL_MOUV_CODE_ARTICLE).Value) = sku Then
            n = n + 1
            ReDim Preserve mvDate(1 To n)
            ReDim Preserve mvTypeArr(1 To n)
            ReDim Preserve mvQty(1 To n)
            ReDim Preserve mvVal(1 To n)
            ReDim Preserve mvPU(1 To n)
            ' A date cell must go through CDate: Val() on a date string would
            ' parse only the leading day number and corrupt the ordering.
            If IsDate(wsMouv.Cells(i, COL_MOUV_DATE).Value) Then
                mvDate(n) = CDbl(CDate(wsMouv.Cells(i, COL_MOUV_DATE).Value))
            Else
                mvDate(n) = Val(wsMouv.Cells(i, COL_MOUV_DATE).Value)
            End If
            mvTypeArr(n) = UCase(Trim(wsMouv.Cells(i, COL_MOUV_TYPE).Value))
            mvQty(n) = Val(wsMouv.Cells(i, COL_MOUV_QTE).Value)
            mvVal(n) = Val(wsMouv.Cells(i, COL_MOUV_VALEUR).Value)
            mvPU(n) = Val(wsMouv.Cells(i, COL_MOUV_PU).Value)
        End If
    Next i

    ' --- Sort chronologically; row order is not guaranteed to be by date ---
    Dim j As Long, tD As Double, tT As String, tQ As Double, tV As Double, tP As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If mvDate(j) < mvDate(i) Then
                tD = mvDate(i): mvDate(i) = mvDate(j): mvDate(j) = tD
                tT = mvTypeArr(i): mvTypeArr(i) = mvTypeArr(j): mvTypeArr(j) = tT
                tQ = mvQty(i): mvQty(i) = mvQty(j): mvQty(j) = tQ
                tV = mvVal(i): mvVal(i) = mvVal(j): mvVal(j) = tV
                tP = mvPU(i): mvPU(i) = mvPU(j): mvPU(j) = tP
            End If
        Next j
    Next i

    ' --- Pass 2: moving weighted average, in date order ---
    Dim qtyOnHand As Double: qtyOnHand = initialQty
    Dim valueOnHand As Double: valueOnHand = initialQty * unitPrice
    Dim cmupNow As Double
    If qtyOnHand > 0 Then cmupNow = unitPrice Else cmupNow = 0

    Dim unitCost As Double
    For i = 1 To n
        If mvTypeArr(i) = "IN" Then
            ' VALEUR is written as qty * pu, so it is the reliable cost basis;
            ' fall back to the per-unit column when VALEUR is absent.
            If mvQty(i) > 0 And mvVal(i) > 0 Then
                unitCost = mvVal(i) / mvQty(i)
            ElseIf mvPU(i) > 0 Then
                unitCost = mvPU(i)
            Else
                unitCost = 0
            End If
            qtyOnHand = qtyOnHand + mvQty(i)
            valueOnHand = valueOnHand + (mvQty(i) * unitCost)

        ElseIf mvTypeArr(i) = "OUT" Then
            ' An issue leaves the unit cost unchanged; it removes qty at the
            ' average prevailing at that date.
            If qtyOnHand > 0 Then
                valueOnHand = valueOnHand - (mvQty(i) * cmupNow)
                qtyOnHand = qtyOnHand - mvQty(i)
                If qtyOnHand < 0 Then qtyOnHand = 0
                If valueOnHand < 0 Then valueOnHand = 0
            End If
        End If

        If qtyOnHand > 0 Then cmupNow = valueOnHand / qtyOnHand Else cmupNow = 0
    Next i

    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    CalculateCMUP = cmupNow
    On Error GoTo 0
    Exit Function

ErrorHandler:
    On Error Resume Next
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    CalculateCMUP = 0
    On Error GoTo 0
End Function

' ================================================================================
' SUB: RefreshAllCMUP
' Recalculates CMUP for all articles in ARTICLES sheet
' ================================================================================
Public Sub RefreshAllCMUP()
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.count, COL_ART_CODE).End(xlUp).Row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Long, cmup As Double
    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        If sku <> "" Then
            cmup = CalculateCMUP(sku)
            If cmup > 0 Then wsArt.Cells(i, COL_ART_CMUP).Value = cmup
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Application.UserControl Then
        MsgBox "CMUP (Prix Moyen) mis " & Chr(233) & " jour.", vbInformation, mod_Config.SYS_TITLE
    End If
End Sub

' ================================================================================
' SUB: UpdateAllABCClassifications
' Calculates ABC classification based on annual consumption value.
' A: Top 80%, B: 15%, C: 5%
' Added: Session 22 - was referenced by mod_SyncBridge.SyncMetricsFromLedger,
'       mod_Analysis.RunAnalysis, and mod_TaskOrchestrator task registry
'       (string label only) but the sub itself was missing.
' ================================================================================
Public Sub UpdateAllABCClassifications(Optional ByVal silent As Boolean = False)
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.count, COL_ART_CODE).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim i As Long
    Dim totalValue As Double: totalValue = 0
    Dim articleValues() As Double: ReDim articleValues(2 To lastRow)
    Dim articleCodes() As String: ReDim articleCodes(2 To lastRow)

    wsArt.Unprotect Password:=mod_Config.MASTER_PWD

    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, COL_ART_CODE).Value)
        If sku <> "" Then
            Dim AnnualDemand As Double: AnnualDemand = GetAnnualDemandFromHistory(sku)
            Dim pu As Double: pu = Val(wsArt.Cells(i, COL_ART_PU).Value)
            articleValues(i) = AnnualDemand * pu
            articleCodes(i) = sku
            totalValue = totalValue + articleValues(i)
        End If
    Next i

    If totalValue = 0 Then
        wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        Exit Sub
    End If

    Dim j As Long, tempVal As Double, tempCode As String
    For i = 2 To lastRow - 1
        For j = i + 1 To lastRow
            If articleValues(i) < articleValues(j) Then
                tempVal = articleValues(i): articleValues(i) = articleValues(j): articleValues(j) = tempVal
                tempCode = articleCodes(i): articleCodes(i) = articleCodes(j): articleCodes(j) = tempCode
            End If
        Next j
    Next i

    Dim cumulativeValue As Double: cumulativeValue = 0
    For i = 2 To lastRow
        cumulativeValue = cumulativeValue + articleValues(i)
        Dim ratio As Double: ratio = cumulativeValue / totalValue
        Dim abcClass As String

        If ratio <= 0.8 Then
            abcClass = "A"
        ElseIf ratio <= 0.95 Then
            abcClass = "B"
        Else
            abcClass = "C"
        End If

        Dim foundRow As Variant
        foundRow = Application.Match(articleCodes(i), wsArt.Columns(COL_ART_CODE), 0)
        If Not IsError(foundRow) Then
            wsArt.Cells(foundRow, COL_ART_CLASSE_ABC).Value = abcClass
        End If
    Next i

    wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    If Not silent Then
        MsgBox "Classifications ABC mises a jour.", vbInformation, mod_Config.SYS_TITLE
    End If
End Sub

' ================================================================================
' FUNCTION: GetArticleStock
' Returns current stock quantity for an article (reads from ARTICLES column C)
' ================================================================================
Public Function GetArticleStock(ByVal sku As String) As Double
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticleStock = 0
        Exit Function
    End If
    
    foundRow = Application.Match(sku, wsArt.Columns(COL_ART_CODE), 0)
    
    If IsError(foundRow) Then
        GetArticleStock = 0
        Exit Function
    End If
    
    GetArticleStock = mod_Utilities.SafeVal(wsArt.Cells(foundRow, COL_ART_STOCK).Value)
End Function

' ================================================================================
' FUNCTION: GetNextSequence
' Returns next sequence number for a given doc prefix (e.g. "BS-", "BR-", "BC-")
' Scans MOUVEMENTS sheet column COL_MOUV_REF_DOC for existing refs with the
' prefix, finds the max suffix number, and returns max+1.
' Added: Session 22 - was called from mod_StockEntry_Logic.GenerateAutoRef
'       (mod_StockEngine.GetNextSequence) but the function itself was missing.
'       Reference impl lifted from external/lsm-vba-core (where it lived as
'       Private in mod_StockEntry_Logic); moved here to keep stock math co-located.
' ================================================================================
Public Function GetNextSequence(ByVal prefix As String) As Long
    Dim wsMouv   As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim maxSeq   As Long
    Dim refStr   As String

    maxSeq = 0

    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0

    If wsMouv Is Nothing Then
        GetNextSequence = 1
        Exit Function
    End If

    lastRow = wsMouv.Cells(wsMouv.Rows.Count, COL_MOUV_REF_DOC).End(xlUp).Row

    For i = 3 To lastRow
        refStr = CStr(wsMouv.Cells(i, COL_MOUV_REF_DOC).Value)
        If Left(refStr, Len(prefix)) = prefix And InStr(refStr, "-") > 0 Then
            Dim parts() As String
            parts = Split(refStr, "-")
            If UBound(parts) >= 2 Then
                Dim seqNum As Long
                On Error Resume Next
                seqNum = CLng(parts(UBound(parts)))
                If Err.Number = 0 And seqNum > maxSeq Then maxSeq = seqNum
                On Error GoTo 0
            End If
        End If
    Next i

    GetNextSequence = maxSeq + 1
End Function

