Attribute VB_Name = "GlobE_XML_Export"
Option Explicit

' ===========================================================================
' GloBE Information Return (GIR) XML Export
' Swiss QDMTT 2024  |  OECD GloBE XML Schema, January 2025
' ===========================================================================

' --- UPDATE THESE FOR EACH CLIENT -------------------------------------------
Private Const COMPANY_NAME  As String = "PLACEHOLDER_COMPANY_AG"
Private Const TIN_VALUE     As String = "CHE-123456789"
Private Const TIN_ISSUED_BY As String = "CH"
Private Const TIN_TYPE      As String = "GIR3001"
Private Const GLOBE_STATUS  As String = "GIR301"
Private Const RULES         As String = "GIR204"
Private Const CCY           As String = "CHF"
Private Const JURISDICTION  As String = "CH"
Private Const FAS_STANDARD  As String = "Swiss GAAP FER"
Private Const CFS_OF_UPE    As String = "GIR501"
Private Const PERIOD_START  As String = "2024-01-01"
Private Const PERIOD_END    As String = "2024-12-31"
Private Const SHEET_NAME    As String = "QDMTT 2024"
Private Const DATA_COL      As Long = 14   ' Column N = jurisdictional totals

' --- SUMMARY ROW NUMBERS (Column N) -----------------------------------------
Private Const ROW_FANIL    As Long = 236
Private Const ROW_NGL_INC  As Long = 264
Private Const ROW_CURR_TAX As Long = 295
Private Const ROW_COV_TAX  As Long = 314


' ===========================================================================
' MAIN EXPORT
' ===========================================================================
Public Sub ExportToGlobEXML()
    Dim ws As Worksheet
    Dim fn As Integer
    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    ' Read summary values from Column N
    Dim fanil   As Long:  fanil   = CellLong(ws, ROW_FANIL)
    Dim nglInc  As Long:  nglInc  = CellLong(ws, ROW_NGL_INC)
    Dim currTax As Long:  currTax = CellLong(ws, ROW_CURR_TAX)
    Dim covTax  As Long:  covTax  = CellLong(ws, ROW_COV_TAX)
    Dim etr     As String: etr    = CalcETR(covTax, nglInc)

    Dim msgRef As String
    msgRef = JURISDICTION & "2024" & JURISDICTION & Format(Now(), "YYYYMMDDHHmmss")
    Dim ts As String
    ts = Format(Now(), "YYYY-MM-DD") & "T" & Format(Now(), "HH:MM:SS")

    ' Build XML string
    Dim LF As String: LF = vbCrLf
    Dim x  As String

    x = "<?xml version='1.0' encoding='utf-8'?>" & LF
    x = x & "<GloBE_Message xmlns=""urn:oecd:ties:gir:v1"">" & LF

    ' MessageHeader
    x = x & "  <MessageHeader>" & LF
    x = x & "    <TransmittingCountry>" & JURISDICTION & "</TransmittingCountry>" & LF
    x = x & "    <ReceivingCountry>" & JURISDICTION & "</ReceivingCountry>" & LF
    x = x & "    <MessageType>GIR</MessageType>" & LF
    x = x & "    <MessageRefID>" & msgRef & "</MessageRefID>" & LF
    x = x & "    <MessageTypeIndic>GIR101</MessageTypeIndic>" & LF
    x = x & "    <ReportingPeriod>" & PERIOD_END & "</ReportingPeriod>" & LF
    x = x & "    <Timestamp>" & ts & "</Timestamp>" & LF
    x = x & "  </MessageHeader>" & LF

    ' FilingInfo
    x = x & "  <GloBE_Body>" & LF
    x = x & "    <FilingInfo>" & LF
    x = x & "      <FilingCE>" & LF
    x = x & "        <ID>" & LF
    x = x & "          <Name>" & COMPANY_NAME & "</Name>" & LF
    x = x & "          <ResCountryCode>" & JURISDICTION & "</ResCountryCode>" & LF
    x = x & "          <TIN issuedBy=""" & TIN_ISSUED_BY & """ TypeOfTIN=""" & TIN_TYPE & """>" & TIN_VALUE & "</TIN>" & LF
    x = x & "          <Rules>" & RULES & "</Rules>" & LF
    x = x & "          <GlobeStatus>" & GLOBE_STATUS & "</GlobeStatus>" & LF
    x = x & "        </ID>" & LF
    x = x & "      </FilingCE>" & LF
    x = x & "      <AccountingInfo>" & LF
    x = x & "        <CFSofUPE>" & CFS_OF_UPE & "</CFSofUPE>" & LF
    x = x & "        <FAS>" & FAS_STANDARD & "</FAS>" & LF
    x = x & "        <Currency currCode=""" & CCY & """ />" & LF
    x = x & "      </AccountingInfo>" & LF
    x = x & "      <Period>" & LF
    x = x & "        <Start>" & PERIOD_START & "</Start>" & LF
    x = x & "        <End>" & PERIOD_END & "</End>" & LF
    x = x & "      </Period>" & LF
    x = x & "      <NameMNE>" & COMPANY_NAME & "</NameMNE>" & LF
    x = x & "    </FilingInfo>" & LF

    ' JurisdictionSection
    x = x & "    <JurisdictionSection>" & LF
    x = x & "      <GloBE_Tax>" & LF
    x = x & "        <ETR>" & LF
    x = x & "          <ETR_Status>" & LF
    x = x & "            <ETR_Computation>" & LF
    x = x & "              <OverallComputation>" & LF
    x = x & "                <FANIL>" & fanil & "</FANIL>" & LF
    x = x & "                <AdjustedFANIL>" & fanil & "</AdjustedFANIL>" & LF

    ' NetGlobeIncome + GIR2001-GIR2026
    x = x & "                <NetGlobeIncome>" & LF
    x = x & "                  <Total>" & nglInc & "</Total>" & LF
    x = x & AdjRow(ws, 238, "GIR2001")
    x = x & AdjRow(ws, 239, "GIR2002")
    x = x & AdjRow(ws, 240, "GIR2003")
    x = x & AdjRow(ws, 241, "GIR2004")
    x = x & AdjRow(ws, 242, "GIR2005")
    x = x & AdjRow(ws, 243, "GIR2006")
    x = x & AdjRow(ws, 244, "GIR2007")
    x = x & AdjRow(ws, 245, "GIR2008")
    x = x & AdjRow(ws, 246, "GIR2009")
    x = x & AdjRow(ws, 247, "GIR2010")
    x = x & AdjRow(ws, 248, "GIR2011")
    x = x & AdjRow(ws, 249, "GIR2012")
    x = x & AdjRow(ws, 250, "GIR2013")
    x = x & AdjRow(ws, 251, "GIR2014")
    x = x & AdjRow(ws, 252, "GIR2015")
    x = x & AdjRow(ws, 253, "GIR2016")
    x = x & AdjRow(ws, 254, "GIR2017")
    x = x & AdjRow(ws, 255, "GIR2018")
    x = x & AdjRow(ws, 256, "GIR2019")
    x = x & AdjRow(ws, 257, "GIR2020")
    x = x & AdjRow(ws, 258, "GIR2021")
    x = x & AdjRow(ws, 259, "GIR2022")
    x = x & AdjRow(ws, 260, "GIR2023")
    x = x & AdjRow(ws, 261, "GIR2024")
    x = x & AdjRow(ws, 262, "GIR2025")
    x = x & AdjRow(ws, 263, "GIR2026")
    x = x & "                </NetGlobeIncome>" & LF

    x = x & "                <IncomeTaxExpense>" & currTax & "</IncomeTaxExpense>" & LF
    x = x & "                <ETRRate>" & etr & "</ETRRate>" & LF
    x = x & "                <TopUpTaxPercentage>0.0000</TopUpTaxPercentage>" & LF

    ' AdjustedCoveredTax + GIR2701-GIR2720
    x = x & "                <AdjustedCoveredTax>" & LF
    x = x & "                  <Total>" & covTax & "</Total>" & LF
    x = x & "                  <AggregrateCurrentTax>" & currTax & "</AggregrateCurrentTax>" & LF
    x = x & AdjRow(ws, 297, "GIR2701")
    x = x & AdjRow(ws, 298, "GIR2703")
    x = x & AdjRow(ws, 299, "GIR2704")
    x = x & AdjRow(ws, 300, "GIR2705")
    x = x & AdjRow(ws, 301, "GIR2706")
    x = x & AdjRow(ws, 302, "GIR2707")
    x = x & AdjRow(ws, 303, "GIR2708")
    x = x & AdjRow(ws, 304, "GIR2709")
    x = x & AdjRow(ws, 305, "GIR2710")
    x = x & AdjRow(ws, 306, "GIR2711")
    x = x & AdjRow(ws, 307, "GIR2712")
    x = x & AdjRow(ws, 308, "GIR2713")
    x = x & AdjRow(ws, 309, "GIR2714")
    x = x & AdjRow(ws, 310, "GIR2715")
    x = x & AdjRow(ws, 311, "GIR2716")
    x = x & AdjRow(ws, 312, "GIR2717")
    x = x & AdjRow(ws, 313, "GIR2718")
    x = x & AdjRowCol(ws, 95, 8, "GIR2719")
    x = x & AdjRowCol(ws, 96, 8, "GIR2720")
    x = x & "                </AdjustedCoveredTax>" & LF

    x = x & "              </OverallComputation>" & LF
    x = x & "            </ETR_Computation>" & LF
    x = x & "          </ETR_Status>" & LF
    x = x & "        </ETR>" & LF
    x = x & "      </GloBE_Tax>" & LF
    x = x & "    </JurisdictionSection>" & LF
    x = x & "  </GloBE_Body>" & LF
    x = x & "</GloBE_Message>"

    ' Save XML file
    Dim outDir  As String: outDir  = ThisWorkbook.Path & Application.PathSeparator & "output"
    Dim outPath As String: outPath = outDir & Application.PathSeparator & "gir_2024_CH.xml"

    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    fn = FreeFile
    Open outPath For Output As #fn
    Print #fn, x
    Close #fn

    MsgBox "Export successful!" & vbCrLf & vbCrLf & outPath, vbInformation, "GloBE XML Export"
    Exit Sub

ErrHandler:
    If fn > 0 Then Close #fn
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Export Failed"
End Sub


' ===========================================================================
' ONE-TIME SETUP: run once to add the Export button to the sheet
' ===========================================================================
Public Sub AddExportButton()
    Dim ws  As Worksheet
    Dim btn As Object

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    On Error Resume Next
    ws.Buttons("btnExportXML").Delete
    On Error GoTo 0

    Dim btnLeft   As Double: btnLeft   = ws.Cells(1, 26).Left
    Dim btnTop    As Double: btnTop    = ws.Cells(1, 26).Top

    Set btn = ws.Buttons.Add(btnLeft, btnTop, 160, 30)
    With btn
        .Name     = "btnExportXML"
        .Caption  = "Export to GloBE XML"
        .OnAction = "GlobE_XML_Export.ExportToGlobEXML"
        .Font.Bold = True
        .Font.Size = 11
    End With

    MsgBox "Button added. You're all set!", vbInformation, "Setup Complete"
End Sub


' ===========================================================================
' PRIVATE HELPERS
' ===========================================================================

Private Function CellLong(ws As Worksheet, rowNum As Long) As Long
    Dim v As Variant
    v = ws.Cells(rowNum, DATA_COL).Value
    If IsEmpty(v) Or IsNull(v) Or Not IsNumeric(v) Then
        CellLong = 0
    Else
        CellLong = CLng(v)
    End If
End Function

Private Function CalcETR(covTax As Long, nglInc As Long) As String
    If nglInc = 0 Then
        CalcETR = "0.0000"
        Exit Function
    End If
    Dim r As Double
    r = CDbl(covTax) / CDbl(nglInc)
    If r < 0 Then r = 0
    If r > 1 Then r = 1
    CalcETR = Format(r, "0.0000")
End Function

Private Function AdjRow(ws As Worksheet, rowNum As Long, girCode As String) As String
    Dim LF  As String: LF  = vbCrLf
    Dim p   As String: p   = "                  "
    Dim amt As Long:   amt = CellLong(ws, rowNum)
    AdjRow = p & "<Adjustments>" & LF & _
             p & "  <Amount>" & amt & "</Amount>" & LF & _
             p & "  <AdjustmentItem>" & girCode & "</AdjustmentItem>" & LF & _
             p & "</Adjustments>" & LF
End Function

Private Function AdjRowCol(ws As Worksheet, rowNum As Long, colNum As Long, girCode As String) As String
    Dim LF  As String: LF = vbCrLf
    Dim p   As String: p  = "                  "
    Dim v   As Variant: v = ws.Cells(rowNum, colNum).Value
    Dim amt As Long
    If IsEmpty(v) Or IsNull(v) Or Not IsNumeric(v) Then
        amt = 0
    Else
        amt = CLng(v)
    End If
    AdjRowCol = p & "<Adjustments>" & LF & _
                p & "  <Amount>" & amt & "</Amount>" & LF & _
                p & "  <AdjustmentItem>" & girCode & "</AdjustmentItem>" & LF & _
                p & "</Adjustments>" & LF
End Function
