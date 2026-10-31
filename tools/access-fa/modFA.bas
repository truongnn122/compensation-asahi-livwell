Option Compare Database
Option Explicit

' =====================================================================
' TINH PHU CAP CO DINH (FA) - ALL/AD/2210-005, Phu luc 1, Muc A.II
' Logic doi chieu: kiem-tra/fa_reference.py
' Toan bo chu tieng Viet hien thi lay tu bang tblText (Unicode) qua ham T().
' =====================================================================

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Const FA_VERSION As String = "2026.10.31-4"
Private Const FORM_MAIN As String = "frmMain"
Private Const LOCK_TIMEOUT_MIN As Long = 30
Private Const FA_TITLES As String = ",LP,"    ' Chuc danh duoc nhan FA (Muc A.II: LP toan thoi gian)
Private Const MB_OK As Long = 0
Private Const MB_YESNO As Long = 4
Private Const MB_ICONERROR As Long = 16
Private Const MB_ICONQUESTION As Long = 32
Private Const MB_ICONINFO As Long = 64
Private Const IDYES As Long = 6

' ---- Tham so, doc tu tblParam ----
Private pStdDays As Double
Private pMinAttPct As Double
Private pMinDays As Double
Private pNewGuarAllow As Double
Private pNewGuarMonths As Long
Private pNewGuarMinCt As Double
Private pQtrGuarAllow As Double
Private pQtrMinCt As Double
Private pQtrMinFYC As Double
Private pK2NACoef As Double
Private pMDRTCoef As Double

Private allowRows As Variant   ' (i, 0)=MinFYC, (i,1)=MinRecruits, (i,2)=AllowancePerDay
Private k2Rows As Variant      ' (i, 0)=MinK2, (i,1)=Coef - sap xep giam dan
Private ctRows As Variant      ' (i, 0)=MinContracts, (i,1)=Coef - sap xep giam dan
Private holidays As Object     ' Scripting.Dictionary
Private mErr As String

' =====================================================================
' Tien ich
' =====================================================================

Public Function FaVersion() As String
    FaVersion = FA_VERSION
End Function

' Sao luu file du lieu truoc thao tac ghi (dung Application.Run de khong phu thuoc modUpdate khi bien dich)
Private Function AutoBackup() As Boolean
    On Error Resume Next
    AutoBackup = (Application.Run("BackupData", True) <> "")
    If Err.Number <> 0 Then AutoBackup = True   ' chua cai modUpdate: bo qua
End Function

Public Function T(ByVal key As String) As String
    T = Nz(DLookup("TextValue", "tblText", "TextKey='" & key & "'"), key)
End Function

' Giai ma chuoi \uXXXX -> Unicode (module chi duoc chua ky tu ASCII)
Public Function U(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" Then
            out = out & ChrW$(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function

Private Sub EnsureText(ByVal key As String, ByVal value As String)
    If IsNull(DLookup("TextKey", "tblText", "TextKey='" & key & "'")) Then
        CurrentDb.Execute "INSERT INTO tblText (TextKey, TextValue) VALUES ('" & key & "', " & Q(value) & ")", dbFailOnError
    End If
End Sub

' Bo sung nhan cho file FA_App da dung truoc phien ban nay
Private Sub EnsureTexts()
    EnsureText "R_NotLP", U("Ch\u1EE9c danh kh\u00F4ng ph\u1EA3i LP (FA ch\u1EC9 \u00E1p d\u1EE5ng LP to\u00E0n th\u1EDDi gian)")
    EnsureText "E_AgentCols", U("Sheet Agent thi\u1EBFu c\u1ED9t b\u1EAFt bu\u1ED9c: m\u00E3 \u0111\u1EA1i l\u00FD (AgentCode/AGENT_CODE), ng\u00E0y b\u1ED5 nhi\u1EC7m (AppointDate/APPOINTMENT_DATE), m\u00F4 h\u00ECnh (WorkModel/AGENCY_TYPE).")
    EnsureText "S_AgentsImported", U("\u0110\u00E3 n\u1EA1p danh s\u00E1ch agent, s\u1ED1 d\u00F2ng:")
End Sub

Private Function BackEndPath() As String
    Dim c As String
    c = CurrentDb.TableDefs("tblAgent").Connect
    BackEndPath = Mid$(c, InStr(1, c, "DATABASE=", vbTextCompare) + 9)
End Function

' Bo sung cot Title, AgentStatus cho file FA_Data da dung truoc phien ban nay
Private Sub EnsureAgentColumns()
    Dim be As DAO.Database, f As DAO.Field, hasTitle As Boolean, hasStatus As Boolean
    Set be = DBEngine.OpenDatabase(BackEndPath())
    For Each f In be.TableDefs("tblAgent").Fields
        If f.Name = "Title" Then hasTitle = True
        If f.Name = "AgentStatus" Then hasStatus = True
    Next f
    If Not hasTitle Then be.Execute "ALTER TABLE tblAgent ADD COLUMN Title TEXT(20)", dbFailOnError
    If Not hasStatus Then be.Execute "ALTER TABLE tblAgent ADD COLUMN AgentStatus TEXT(20)", dbFailOnError
    be.Close
    If Not (hasTitle And hasStatus) Then CurrentDb.TableDefs("tblAgent").RefreshLink
End Sub

Private Function MsgU(ByVal text As String, Optional ByVal flags As Long = MB_ICONINFO) As Long
    MsgU = MessageBoxW(Application.hWndAccessApp, StrPtr(text), StrPtr(T("AppTitle")), flags)
End Function

Private Sub SetStatus(ByVal text As String)
    If CurrentProject.AllForms(FORM_MAIN).IsLoaded Then
        Forms(FORM_MAIN)!lblStatus.Caption = text
        DoEvents
    End If
End Sub

' Lam tron half-up (Muc D.s). Round() cua VBA la banker's rounding nen khong dung.
Public Function RoundHU(ByVal x As Double, Optional ByVal digits As Long = 0) As Double
    Dim f As Double
    f = 10 ^ digits
    If x >= 0 Then
        RoundHU = Int(x * f + 0.5 + 0.000000001) / f
    Else
        RoundHU = -Int(-x * f + 0.5 + 0.000000001) / f
    End If
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function Q(ByVal s As String) As String
    Q = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function PeriodStart(ByVal period As String) As Date
    PeriodStart = DateSerial(CInt(Left$(period, 4)), CInt(Right$(period, 2)), 1)
End Function

Private Function PeriodAdd(ByVal period As String, ByVal n As Long) As String
    PeriodAdd = Format$(DateAdd("m", n, PeriodStart(period)), "yyyymm")
End Function

Private Function IsValidPeriod(ByVal period As String) As Boolean
    If Len(period) <> 6 Or Not IsNumeric(period) Then Exit Function
    If CInt(Right$(period, 2)) < 1 Or CInt(Right$(period, 2)) > 12 Then Exit Function
    IsValidPeriod = True
End Function

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = 1
End Function

' So ngay lam viec (tru Thu Bay, Chu Nhat, ngay le) tu d1 den d2
Private Function WorkDaysBetween(ByVal d1 As Date, ByVal d2 As Date) As Long
    Dim d As Date, n As Long
    For d = d1 To d2
        If Weekday(d, vbMonday) <= 5 Then
            If Not holidays.Exists(CLng(d)) Then n = n + 1
        End If
    Next d
    WorkDaysBetween = n
End Function

' =====================================================================
' Tham so
' =====================================================================

Private Function P(ByVal key As String) As Double
    Dim v As Variant
    v = DLookup("ParamValue", "tblParam", "ParamKey='" & key & "'")
    If IsNull(v) Then Err.Raise vbObjectError + 1, , T("E_ParamMissing") & " " & key
    P = CDbl(v)
End Function

Private Function LoadRows(ByVal sql As String, ByVal nCols As Long) As Variant
    Dim rs As DAO.Recordset, arr() As Double, i As Long, j As Long
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 2, , T("E_ParamMissing") & " " & sql
    rs.MoveLast: rs.MoveFirst
    ReDim arr(rs.RecordCount - 1, nCols - 1)
    Do Until rs.EOF
        For j = 0 To nCols - 1
            arr(i, j) = Nz(rs.Fields(j).Value, 0)
        Next j
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    LoadRows = arr
End Function

Private Sub LoadParams()
    Dim rs As DAO.Recordset

    pStdDays = P("StdDays")
    pMinAttPct = P("MinAttendancePct")
    pMinDays = P("MinWorkDays")
    pNewGuarAllow = P("NewGuarAllowance")
    pNewGuarMonths = CLng(P("NewGuarMonths"))
    pNewGuarMinCt = P("NewGuarMinContracts")
    pQtrGuarAllow = P("QtrGuarAllowance")
    pQtrMinCt = P("QtrGuarMinContracts")
    pQtrMinFYC = P("QtrGuarMinFYC")
    pK2NACoef = P("K2_NA_Coef")
    pMDRTCoef = P("MDRT_Coef")

    allowRows = LoadRows("SELECT MinFYC, MinRecruits, AllowancePerDay FROM tblParam_Allowance", 3)
    k2Rows = LoadRows("SELECT MinK2, Coef FROM tblParam_K2 ORDER BY MinK2 DESC", 2)
    ctRows = LoadRows("SELECT MinContracts, Coef FROM tblParam_Contract ORDER BY MinContracts DESC", 2)

    Set holidays = NewDict()
    Set rs = CurrentDb.OpenRecordset("SELECT HolidayDate FROM tblHoliday", dbOpenSnapshot)
    Do Until rs.EOF
        holidays(CLng(DateValue(rs!HolidayDate))) = True
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Muc A: lay muc phu cap cao nhat ma LP thoa (Lua chon 1 hoac Lua chon 2)
Private Function AllowanceFor(ByVal fyc As Double, ByVal recruits As Double) As Double
    Dim i As Long, best As Double
    For i = 0 To UBound(allowRows, 1)
        If fyc >= allowRows(i, 0) And recruits >= allowRows(i, 1) Then
            If allowRows(i, 2) > best Then best = allowRows(i, 2)
        End If
    Next i
    AllowanceFor = best
End Function

' Muc B: %He so K2. K2 = Null nghia la N/A (chua co K2)
Private Function K2CoefFor(ByVal k2 As Variant) As Double
    Dim i As Long, k As Double
    If IsNull(k2) Then
        K2CoefFor = pK2NACoef
        Exit Function
    End If
    k = RoundHU(CDbl(k2), 1)
    For i = 0 To UBound(k2Rows, 1)
        If k >= k2Rows(i, 0) Then
            K2CoefFor = k2Rows(i, 1)
            Exit Function
        End If
    Next i
End Function

' Muc D: %He so hop dong
Private Function ContractCoefFor(ByVal ct As Double) As Double
    Dim i As Long
    For i = 0 To UBound(ctRows, 1)
        If ct >= ctRows(i, 0) Then
            ContractCoefFor = ctRows(i, 1)
            Exit Function
        End If
    Next i
End Function

' =====================================================================
' Trang thai ky (khoa chay trung / chot ky) - bang tblPeriodStatus o file Data
' =====================================================================

Private Function GetUserId() As String
    GetUserId = Environ$("USERNAME") & "@" & Environ$("COMPUTERNAME")
End Function

Private Function GetStatus(ByVal period As String, ByVal ky As Long, _
                           Optional ByRef byUser As String, Optional ByRef at As Variant) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblPeriodStatus WHERE Period=" & Q(period) & " AND Ky=" & ky, dbOpenSnapshot)
    If Not rs.EOF Then
        GetStatus = Nz(rs!Status, "")
        byUser = Nz(rs!UpdatedBy, "")
        at = rs!UpdatedAt
    End If
    rs.Close
End Function

Private Sub SetPeriodStatus(ByVal period As String, ByVal ky As Long, ByVal status As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "DELETE FROM tblPeriodStatus WHERE Period=" & Q(period) & " AND Ky=" & ky, dbFailOnError
    db.Execute "INSERT INTO tblPeriodStatus (Period, Ky, Status, UpdatedBy, UpdatedAt) VALUES (" & _
               Q(period) & ", " & ky & ", " & Q(status) & ", " & Q(GetUserId()) & ", Now())", dbFailOnError
End Sub

Private Function IsLocked(ByVal period As String, ByVal ky As Long) As Boolean
    IsLocked = (GetStatus(period, ky) = "LOCKED")
End Function

' =====================================================================
' Doc du lieu ky truoc vao Dictionary (tranh DLookup tung dong)
' =====================================================================

' Ket qua cua 1 thang, lay theo ky chi tra lon nhat da tinh cua tung LP
Private Function LoadLatestResults(ByVal period As String) As Object
    Dim d As Object, rs As DAO.Recordset
    Set d = NewDict()
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT r.AgentCode, r.Contracts, r.NewGuar, r.NewGuarPending, r.FA_Month, r.QtrGuarMonth " & _
        "FROM tblFA_Result AS r WHERE r.Period=" & Q(period) & " AND r.Ky = " & _
        "(SELECT Max(x.Ky) FROM tblFA_Result AS x WHERE x.Period=r.Period AND x.AgentCode=r.AgentCode)", dbOpenSnapshot)
    Do Until rs.EOF
        d(CStr(rs!AgentCode)) = Array(Nz(rs!Contracts, 0), Nz(rs!NewGuar, 0), Nz(rs!NewGuarPending, 0), _
                                      Nz(rs!FA_Month, 0), Nz(rs!QtrGuarMonth, 0))
        rs.MoveNext
    Loop
    rs.Close
    Set LoadLatestResults = d
End Function

' LP dat dieu kien bao dam FA quy ke tiep: >= 6 HD va >= 24 trieu FYC phat sinh trong quy
Private Function LoadQtrQualified(ByVal qtrEnd As String) As Object
    Dim d As Object, rs As DAO.Recordset, inList As String
    Set d = NewDict()
    inList = Q(PeriodAdd(qtrEnd, -2)) & "," & Q(PeriodAdd(qtrEnd, -1)) & "," & Q(qtrEnd)
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT i.AgentCode, Sum(i.Contracts) AS SumCt, Sum(i.FYC) AS SumFYC " & _
        "FROM tblFA_Input AS i WHERE i.Period IN (" & inList & ") AND i.Ky = " & _
        "(SELECT Max(x.Ky) FROM tblFA_Input AS x WHERE x.Period=i.Period AND x.AgentCode=i.AgentCode) " & _
        "GROUP BY i.AgentCode", dbOpenSnapshot)
    Do Until rs.EOF
        If Nz(rs!SumCt, 0) >= pQtrMinCt And Nz(rs!SumFYC, 0) >= pQtrMinFYC Then d(CStr(rs!AgentCode)) = True
        rs.MoveNext
    Loop
    rs.Close
    Set LoadQtrQualified = d
End Function

Private Function LoadPaidPrevKy(ByVal period As String, ByVal ky As Long) As Object
    Dim d As Object, rs As DAO.Recordset
    Set d = NewDict()
    Set rs = CurrentDb.OpenRecordset("SELECT AgentCode, Sum(Payable) AS S FROM tblFA_Result WHERE Period=" & _
        Q(period) & " AND Ky<" & ky & " GROUP BY AgentCode", dbOpenSnapshot)
    Do Until rs.EOF
        d(CStr(rs!AgentCode)) = Nz(rs!S, 0)
        rs.MoveNext
    Loop
    rs.Close
    Set LoadPaidPrevKy = d
End Function

' =====================================================================
' TINH FA
' =====================================================================

Public Function RunFA(ByVal period As String, ByVal ky As Long) As Boolean
    Dim db As DAO.Database, ws As DAO.Workspace
    Dim rsIn As DAO.Recordset, rsOut As DAO.Recordset
    Dim st As String, stUser As String, stAt As Variant
    Dim m1Res As Object, prevQtrOK As Object, paidPrev As Object, m1Qtr As Object, m2Qtr As Object
    Dim monthStart As Date, monthEnd As Date, monthInQtr As Long, prevQtrEnd As String
    Dim inTrans As Boolean, n As Long, total As Double

    On Error GoTo Fail
    Set db = CurrentDb

    ' ---- Kiem tra trang thai ky ----
    st = GetStatus(period, ky, stUser, stAt)
    If st = "LOCKED" Then
        MsgU T("E_Locked"), MB_ICONERROR: Exit Function
    End If
    If st = "RUNNING" And stUser <> GetUserId() Then
        If DateDiff("n", Nz(stAt, #1/1/2000#), Now()) < LOCK_TIMEOUT_MIN Then
            MsgU T("E_Running") & " " & stUser, MB_ICONERROR: Exit Function
        End If
    End If
    If DCount("*", "tblFA_Result", "Period=" & Q(period) & " AND Ky>" & ky) > 0 Then
        MsgU T("E_LaterKy"), MB_ICONERROR: Exit Function
    End If
    If DCount("*", "tblFA_Input", "Period=" & Q(period) & " AND Ky=" & ky) = 0 Then
        MsgU T("E_NoInput"), MB_ICONERROR: Exit Function
    End If
    If DCount("*", "tblFA_Result", "Period=" & Q(period) & " AND Ky=" & ky) > 0 Then
        If MsgU(T("Q_Recalc"), MB_YESNO Or MB_ICONQUESTION) <> IDYES Then Exit Function
    End If

    EnsureTexts
    EnsureAgentSchema
    SetPeriodStatus period, ky, "RUNNING"
    SetStatus T("S_Loading")
    LoadParams

    monthStart = PeriodStart(period)
    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
    monthInQtr = ((Month(monthStart) - 1) Mod 3) + 1
    prevQtrEnd = PeriodAdd(period, -monthInQtr)

    Set m1Res = LoadLatestResults(PeriodAdd(period, -1))
    Set prevQtrOK = LoadQtrQualified(prevQtrEnd)
    Set paidPrev = LoadPaidPrevKy(period, ky)
    If monthInQtr = 3 And ky = 4 Then
        Set m1Qtr = LoadLatestResults(PeriodAdd(period, -2))
        Set m2Qtr = LoadLatestResults(PeriodAdd(period, -1))
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    db.Execute "DELETE FROM tblFA_Result WHERE Period=" & Q(period) & " AND Ky=" & ky, dbFailOnError

    Set rsIn = db.OpenRecordset( _
        "SELECT i.*, IIf(h.AgentCode Is Null, a.AppointDate, h.AppointDate) AS AppointDate, " & _
        "IIf(h.AgentCode Is Null, a.WorkModel, h.WorkModel) AS WorkModel, " & _
        "IIf(h.AgentCode Is Null, a.TerminateDate, h.TerminateDate) AS TerminateDate, " & _
        "IIf(h.AgentCode Is Null, a.Title, h.Title) AS Title " & _
        "FROM (tblFA_Input AS i LEFT JOIN tblAgent AS a ON i.AgentCode = a.AgentCode) " & _
        "LEFT JOIN (SELECT * FROM tblAgentHist WHERE SNAPSHOT_PERIOD=" & Q(period) & ") AS h ON i.AgentCode = h.AgentCode " & _
        "WHERE i.Period=" & Q(period) & " AND i.Ky=" & ky & " ORDER BY i.AgentCode", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset("tblFA_Result", dbOpenDynaset, dbAppendOnly)

    Do Until rsIn.EOF
        Dim code As String, reason As String, eligible As Boolean
        Dim wm As Long, baseDays As Long, actual As Double, att As Double, attPct As Double, dayF As Double
        Dim fyc As Double, ct As Double, rec As Double, allow As Double
        Dim k2c As Double, mdrtc As Double, ctc As Double, faFormula As Double
        Dim newGuar As Double, pending As Double, retro As Double, guarAmt As Double, cumCt As Double
        Dim faMonth As Double, qtrOK As Boolean, qtrMonth As Double, topup As Double
        Dim faTotal As Double, paid As Double, m1 As Variant

        code = CStr(rsIn!AgentCode)
        reason = ""
        actual = Nz(rsIn!ActualDays, 0)
        fyc = Nz(rsIn!FYC, 0)
        ct = Nz(rsIn!Contracts, 0)
        rec = Nz(rsIn!Recruits, 0)
        newGuar = 0: pending = 0: retro = 0: topup = 0: qtrMonth = 0
        wm = 0: baseDays = 0: attPct = 0: dayF = 0

        ' ---- Thang lam viec & ty le cham cong (Muc D.k, D.z) ----
        If IsNull(rsIn!AppointDate) Then
            reason = T("R_NoAgent")
        ElseIf rsIn!AppointDate > monthEnd Then
            reason = T("R_NotAppointed")
        Else
            wm = DateDiff("m", rsIn!AppointDate, monthStart) + 1
            If wm = 1 Then
                baseDays = WorkDaysBetween(DateValue(rsIn!AppointDate), monthEnd)
            Else
                baseDays = WorkDaysBetween(monthStart, monthEnd)
            End If
            If baseDays > 0 Then att = (actual + 1) / baseDays Else att = 0
            If att > 1 Then att = 1
            attPct = RoundHU(att * 100, 1)
            If wm = 1 Then
                dayF = actual
                If dayF > pStdDays Then dayF = pStdDays
                dayF = RoundHU(dayF, 1)
            Else
                dayF = RoundHU(attPct / 100 * pStdDays, 1)
            End If
        End If

        ' ---- Dieu kien toi thieu ----
        If reason = "" Then
            Dim title As String
            title = UCase$(Trim$(Nz(rsIn!Title, "")))
            If Len(title) > 0 And InStr(FA_TITLES, "," & title & ",") = 0 Then
                reason = T("R_NotLP")
            ElseIf Nz(rsIn!WorkModel, "") <> "FT" Then
                reason = T("R_NotFT")
            ElseIf Not IsNull(rsIn!TerminateDate) Then
                If rsIn!TerminateDate <= monthEnd Then reason = T("R_Terminated")
            End If
        End If
        If reason = "" Then
            If UCase$(Trim$(Nz(rsIn!StatusEOM, ""))) = "SE" Then
                reason = T("R_SE")
            ElseIf Not Nz(rsIn!TrainingDone, False) Then
                reason = T("R_Training")
            ElseIf attPct < pMinAttPct And actual < pMinDays Then
                reason = T("R_Attendance")
            End If
        End If
        eligible = (reason = "")

        ' ---- Cong thuc FA ----
        allow = AllowanceFor(fyc, rec)
        k2c = K2CoefFor(rsIn!K2Pct)
        If Nz(rsIn!MDRTPrev, False) Then mdrtc = pMDRTCoef Else mdrtc = 1
        ctc = ContractCoefFor(ct)
        faFormula = RoundHU(allow * dayF * k2c * mdrtc * ctc, 0)

        ' ---- Bao dam 0,3 trieu/ngay cho 02 thang lam viec dau ----
        If wm >= 1 And wm <= pNewGuarMonths Then
            cumCt = ct
            m1 = Empty
            If wm = 2 And m1Res.Exists(code) Then
                m1 = m1Res(code)
                cumCt = cumCt + m1(0)
            End If
            guarAmt = RoundHU(pNewGuarAllow * dayF, 0)
            If eligible Then
                If cumCt >= pNewGuarMinCt Then newGuar = guarAmt Else pending = guarAmt
            End If
            ' Thang 1 chua du 0,5 HD -> chi bu khi thang 2 dat
            If wm = 2 And Not IsEmpty(m1) And cumCt >= pNewGuarMinCt And reason <> T("R_Terminated") Then
                If m1(1) = 0 And m1(2) > 0 Then retro = MaxD(0, m1(2) - m1(3))
            End If
            If pending > 0 And reason = "" Then reason = T("R_GuarPending")
        End If

        If eligible Then faMonth = MaxD(faFormula, newGuar) Else faMonth = 0

        ' ---- Bao dam FA theo quy: 0,15 trieu/ngay ----
        qtrOK = prevQtrOK.Exists(code)
        If qtrOK And eligible Then qtrMonth = RoundHU(pQtrGuarAllow * dayF, 0)
        If qtrOK And monthInQtr = 3 And ky = 4 Then
            Dim sumGuar As Double, sumFA As Double
            sumGuar = qtrMonth: sumFA = faMonth
            If m1Qtr.Exists(code) Then sumGuar = sumGuar + m1Qtr(code)(4): sumFA = sumFA + m1Qtr(code)(3)
            If m2Qtr.Exists(code) Then sumGuar = sumGuar + m2Qtr(code)(4): sumFA = sumFA + m2Qtr(code)(3)
            topup = MaxD(0, sumGuar - sumFA)
        End If

        faTotal = faMonth + retro + topup
        If paidPrev.Exists(code) Then paid = paidPrev(code) Else paid = 0

        ' ---- Ghi ket qua ----
        rsOut.AddNew
        rsOut!Period = period
        rsOut!Ky = ky
        rsOut!AgentCode = code
        rsOut!WorkMonthNo = wm
        rsOut!WorkDaysBase = baseDays
        rsOut!ActualDays = actual
        rsOut!AttendancePct = attPct
        rsOut!DayFactor = dayF
        rsOut!Eligible = eligible
        rsOut!Reason = reason
        rsOut!FYC = fyc
        rsOut!Contracts = ct
        rsOut!Recruits = rec
        rsOut!AllowancePerDay = allow
        rsOut!K2Pct = rsIn!K2Pct
        rsOut!K2Coef = k2c
        rsOut!MDRTCoef = mdrtc
        rsOut!ContractCoef = ctc
        rsOut!FA_Formula = faFormula
        rsOut!NewGuar = newGuar
        rsOut!NewGuarPending = pending
        rsOut!NewGuarRetro = retro
        rsOut!FA_Month = faMonth
        rsOut!QtrGuarFlag = qtrOK
        rsOut!QtrGuarMonth = qtrMonth
        rsOut!QtrGuarTopup = topup
        rsOut!FA_Total = faTotal
        rsOut!PaidPrevKy = paid
        rsOut!Payable = faTotal - paid
        rsOut!CalcAt = Now()
        rsOut!CalcBy = GetUserId()
        rsOut.Update

        n = n + 1
        total = total + faTotal - paid
        If n Mod 200 = 0 Then SetStatus T("S_Calculating") & " " & n
        rsIn.MoveNext
    Loop
    rsIn.Close
    rsOut.Close

    ws.CommitTrans
    inTrans = False
    SetPeriodStatus period, ky, "CALC"

    SetStatus T("S_Done") & " " & n & " LP | " & Format$(total, "#,##0")
    MsgU T("S_Done") & " " & n & " LP" & vbCrLf & T("S_TotalPayable") & " " & Format$(total, "#,##0")
    RunFA = True
    Exit Function

Fail:
    Dim msg As String
    msg = Err.Description
    If inTrans Then ws.Rollback
    On Error Resume Next
    SetPeriodStatus period, ky, IIf(st = "", "ERROR", st)
    SetStatus T("E_Calc")
    MsgU T("E_Calc") & vbCrLf & msg, MB_ICONERROR
End Function

' =====================================================================
' NHAP DU LIEU TU EXCEL (sheet Agent, Holiday, FA_Input)
' =====================================================================

Private Function PickFile(ByVal title As String, ByVal filterName As String, ByVal filterExt As String) As String
    With Application.FileDialog(3)   ' msoFileDialogFilePicker
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add filterName, filterExt
        If .Show Then PickFile = .SelectedItems(1)
    End With
End Function

Private Sub DropTable(ByVal name As String)
    On Error Resume Next
    DoCmd.DeleteObject acTable, name
End Sub

Private Function SheetExists(ByVal path As String, ByVal sheet As String) As Boolean
    On Error GoTo NoSheet
    DropTable "stg_" & sheet
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "stg_" & sheet, path, True, sheet & "!"
    SheetExists = True
NoSheet:
End Function

' Cau truc 98 cot cua file danh sach agent xuat tu he thong (Tbl_Agent.xlsx).
' Moi dong: TieuDeGoc|LanXuatHien|TenTruong|Kieu(T=chu, D=ngay, C/CD=truong FA co san)|TieuDeHienThi
Private Function AgentSpec() As Variant
    Dim s As String
    s = s & "Period|1|SNAPSHOT_PERIOD|T|Period;"
    s = s & "Channel|1|CHANNEL|T|Channel;"
    s = s & "REGION_CODE|1|REGION_CODE|T|REGION_CODE;"
    s = s & "REGION_NAME|1|REGION_NAME|T|REGION_NAME;"
    s = s & "ZONE_CODE|1|ZONE_CODE|T|ZONE_CODE;"
    s = s & "ZONE_NAME|1|ZONE_NAME|T|ZONE_NAME;"
    s = s & "BRANCH_CODE|1|BRANCH_CODE|T|BRANCH_CODE;"
    s = s & "BRANCH_NAME|1|BRANCH_NAME|T|BRANCH_NAME;"
    s = s & "SH_CODE|1|SH_CODE|T|SH_CODE;"
    s = s & "SH_NAME|1|SH_NAME|T|SH_NAME;"
    s = s & "SD_CODE|1|SD_CODE|T|SD_CODE;"
    s = s & "SD_NAME|1|SD_NAME|T|SD_NAME;"
    s = s & "INDIRECT 2_NAME|1|INDIRECT_2_CODE|T|INDIRECT 2_CODE;"
    s = s & "INDIRECT 2_NAME|2|INDIRECT_2_NAME|T|INDIRECT 2_NAME;"
    s = s & "INDIRECT 1_CODE|1|INDIRECT_1_CODE|T|INDIRECT 1_CODE;"
    s = s & "INDIRECT 1_NAME|1|INDIRECT_1_NAME|T|INDIRECT 1_NAME;"
    s = s & "DIRECT_MANAGER_CODE|1|DIRECT_MANAGER_CODE|T|DIRECT_MANAGER_CODE;"
    s = s & "DIRECT_MANAGER_NAME|1|DIRECT_MANAGER_NAME|T|DIRECT_MANAGER_NAME;"
    s = s & "LEADER|1|LEADER|T|LEADER;"
    s = s & "EXT_REF|1|EXT_REF|T|EXT_REF;"
    s = s & "AGENT_CODE|1|AgentCode|C|AGENT_CODE;"
    s = s & "FULL_NAME|1|AgentName|C|FULL_NAME;"
    s = s & "GENDER|1|GENDER|T|GENDER;"
    s = s & "TITLE (Code)|1|Title|C|TITLE (Code);"
    s = s & "AGENT_STATUS|1|AgentStatus|C|AGENT_STATUS;"
    s = s & "DOB|1|DOB|D|DOB;"
    s = s & "NATIONAL_ID|1|NATIONAL_ID|T|NATIONAL_ID;"
    s = s & "ID_ISSUE_DATE|1|ID_ISSUE_DATE|D|ID_ISSUE_DATE;"
    s = s & "ID_ISSUED_BY|1|ID_ISSUED_BY|T|ID_ISSUED_BY;"
    s = s & "ID_EXPIRATION_DATE|1|ID_EXPIRATION_DATE|D|ID_EXPIRATION_DATE;"
    s = s & "OLD_ID_NO|1|OLD_ID_NO|T|OLD_ID_NO;"
    s = s & "PLACE_OF_BIRTH|1|PLACE_OF_BIRTH|T|PLACE_OF_BIRTH;"
    s = s & "NATIONALITY|1|NATIONALITY|T|NATIONALITY;"
    s = s & "MARITAL_STATUS|1|MARITAL_STATUS|T|MARITAL_STATUS;"
    s = s & "TAX_CODE|1|TAX_CODE|T|TAX_CODE;"
    s = s & "INCOME_RANGE|1|INCOME_RANGE|T|INCOME_RANGE;"
    s = s & "EDUCATION|1|EDUCATION|T|EDUCATION;"
    s = s & "PUBLIC_SECTOR|1|PUBLIC_SECTOR|T|PUBLIC_SECTOR;"
    s = s & "OCCUPATION|1|OCCUPATION|T|OCCUPATION;"
    s = s & "EXPERIENCE (Year)|1|EXPERIENCE_YEAR|T|EXPERIENCE (Year);"
    s = s & "PERM_PROVINCE|1|PERM_PROVINCE|T|PERM_PROVINCE;"
    s = s & "PERM_WARD|1|PERM_WARD|T|PERM_WARD;"
    s = s & "PERM_STREET|1|PERM_STREET|T|PERM_STREET;"
    s = s & "CONTACT_PROVINCE|1|CONTACT_PROVINCE|T|CONTACT_PROVINCE;"
    s = s & "CONTACT_WARD|1|CONTACT_WARD|T|CONTACT_WARD;"
    s = s & "CONTACT_STREET|1|CONTACT_STREET|T|CONTACT_STREET;"
    s = s & "MOBILE|1|MOBILE|T|MOBILE;"
    s = s & "PERSONAL_EMAIL|1|PERSONAL_EMAIL|T|PERSONAL_EMAIL;"
    s = s & "COMPANY_EMAIL|1|COMPANY_EMAIL|T|COMPANY_EMAIL;"
    s = s & "CHANNEL|2|CHANNEL_2|T|CHANNEL (2);"
    s = s & "AGENCY_TYPE|1|WorkModel|C|AGENCY_TYPE;"
    s = s & "FIRST_RECRUITER_CODE|1|FIRST_RECRUITER_CODE|T|FIRST_RECRUITER_CODE;"
    s = s & "REFERRER_CODE|1|REFERRER_CODE|T|REFERRER_CODE;"
    s = s & "APPOINTMENT_DATE|1|AppointDate|CD|APPOINTMENT_DATE;"
    s = s & "TERMINATION_DATE|1|TerminateDate|CD|TERMINATION_DATE;"
    s = s & "FIRST APPOINTMENT_DATE|1|FIRST_APPOINTMENT_DATE|D|FIRST APPOINTMENT_DATE;"
    s = s & "PROMOTION_DATE|1|PROMOTION_DATE|D|PROMOTION_DATE;"
    s = s & "DEMOTION_DATE|1|DEMOTION_DATE|D|DEMOTION_DATE;"
    s = s & "BANK_ACCOUNT_HOLDER|1|BANK_ACCOUNT_HOLDER|T|BANK_ACCOUNT_HOLDER;"
    s = s & "BANK_ACCOUNT_NO|1|BANK_ACCOUNT_NO|T|BANK_ACCOUNT_NO;"
    s = s & "BANK_CODE|1|BANK_CODE|T|BANK_CODE;"
    s = s & "BANK_BRANCH|1|BANK_BRANCH|T|BANK_BRANCH;"
    s = s & "BANK_BRANCH CODE (Payment)|1|BANK_BRANCH_CODE_PAYMENT|T|BANK_BRANCH CODE (Payment);"
    s = s & "MOF_CERT_NO|1|MOF_CERT_NO|T|MOF_CERT_NO;"
    s = s & "MOF_CERT_ISSUE_DATE|1|MOF_CERT_ISSUE_DATE|D|MOF_CERT_ISSUE_DATE;"
    s = s & "MOF_EXAM_PASS_DATE|1|MOF_EXAM_PASS_DATE|D|MOF_EXAM_PASS_DATE;"
    s = s & "AVICAD_STATUS|1|AVICAD_STATUS|T|AVICAD_STATUS;"
    s = s & "AVICAD_DESC|1|AVICAD_DESC|T|AVICAD_DESC;"
    s = s & "AVICAD_CHECKED_AT|1|AVICAD_CHECKED_AT|D|AVICAD_CHECKED_AT;"
    s = s & "CONTRACT_NO Agent|1|CONTRACT_NO_AGENT|T|CONTRACT_NO Agent;"
    s = s & "CONTRACT_TYPE|1|CONTRACT_TYPE_AGENT|T|CONTRACT_TYPE Agent;"
    s = s & "CONTRACT_NO Leader|1|CONTRACT_NO_LEADER|T|CONTRACT_NO Leader;"
    s = s & "CONTRACT_TYPE|2|CONTRACT_TYPE_LEADER|T|CONTRACT_TYPE Leader;"
    s = s & "ESIGN_ENVELOPE_ID (DOCUMENT ID)|1|ESIGN_ENVELOPE_ID_DOCUMENT_ID|T|ESIGN_ENVELOPE_ID (DOCUMENT ID);"
    s = s & "CONTRACT_SENT_DATE|1|CONTRACT_SENT_DATE|D|CONTRACT_SENT_DATE;"
    s = s & "CONTRACT_SIGNED_DATE|1|CONTRACT_SIGNED_DATE|D|CONTRACT_SIGNED_DATE;"
    s = s & "SIGN_AUTH_METHOD|1|SIGN_AUTH_METHOD|T|SIGN_AUTH_METHOD;"
    s = s & "SIGN_STATUS|1|SIGN_STATUS|T|SIGN_STATUS;"
    s = s & "CONTRACT_FILE_NAME|1|CONTRACT_FILE_NAME|T|CONTRACT_FILE_NAME;"
    s = s & "CONTRACT_DMS_REF|1|CONTRACT_DMS_REF|T|CONTRACT_DMS_REF;"
    s = s & "PROGRAM_CODES|1|PROGRAM_CODES|T|PROGRAM_CODES;"
    s = s & "COMMIT_1|1|COMMIT_1|T|COMMIT_1;"
    s = s & "COMMIT_2|1|COMMIT_2|T|COMMIT_2;"
    s = s & "COMMIT_3|1|COMMIT_3|T|COMMIT_3;"
    s = s & "COMMIT_4|1|COMMIT_4|T|COMMIT_4;"
    s = s & "CONSENT_NOTIFY_OPS|1|CONSENT_NOTIFY_OPS|T|CONSENT_NOTIFY_OPS;"
    s = s & "CONSENT_NOTIFY_PROMO|1|CONSENT_NOTIFY_PROMO|T|CONSENT_NOTIFY_PROMO;"
    s = s & "PDPD_1_BASIC|1|PDPD_1_BASIC|T|PDPD_1_BASIC;"
    s = s & "PDPD_2_SENSITIVE|1|PDPD_2_SENSITIVE|T|PDPD_2_SENSITIVE;"
    s = s & "PDPD_3_SHARE|1|PDPD_3_SHARE|T|PDPD_3_SHARE;"
    s = s & "CONSENT_DATE|1|CONSENT_DATE|D|CONSENT_DATE;"
    s = s & "CONSENT_METHOD|1|CONSENT_METHOD|T|CONSENT_METHOD;"
    s = s & "ONBOARD_COURSE_CODE|1|ONBOARD_COURSE_CODE|T|ONBOARD_COURSE_CODE;"
    s = s & "ONBOARD_COURSE_DATE|1|ONBOARD_COURSE_DATE|D|ONBOARD_COURSE_DATE;"
    s = s & "TRAINING_COMPLETED|1|TRAINING_COMPLETED|T|TRAINING_COMPLETED;"
    s = s & "PRODUCT_AUTH_CODES|1|PRODUCT_AUTH_CODES|T|PRODUCT_AUTH_CODES;"
    s = s & "DOC_RECEIVED|1|DOC_RECEIVED|T|DOC_RECEIVED;"
    s = s & "DOC_DMS_REF|1|DOC_DMS_REF|T|DOC_DMS_REF;"
    AgentSpec = Split(Left$(s, Len(s) - 1), ";")
End Function

Private Function NormHeader(ByVal h As Variant) As String
    NormHeader = UCase$(Replace(Replace(Trim$(h & ""), " ", ""), "_", ""))
End Function

' Them cac cot cua file he thong vao tblAgent (file FA_Data). Chay lai nhieu lan khong sao.
Private Sub EnsureAgentSchema()
    Dim be As DAO.Database, have As Object, f As DAO.Field, spec As Variant, p As Variant, i As Long, added As Boolean
    EnsureAgentColumns
    Set be = DBEngine.OpenDatabase(BackEndPath())
    Set have = NewDict()
    For Each f In be.TableDefs("tblAgent").Fields
        have(UCase$(f.Name)) = True
    Next f
    spec = AgentSpec()
    For i = 0 To UBound(spec)
        p = Split(spec(i), "|")
        If Left$(p(3), 1) <> "C" And Not have.Exists(UCase$(p(2))) Then
            If p(3) = "D" Then
                be.Execute "ALTER TABLE tblAgent ADD COLUMN [" & p(2) & "] DATETIME", dbFailOnError
            Else
                be.Execute "ALTER TABLE tblAgent ADD COLUMN [" & p(2) & "] TEXT(255)", dbFailOnError
            End If
            added = True
        End If
    Next i
    be.TableDefs.Refresh
    EnsureHistTable be
    be.Close
    If added Then CurrentDb.TableDefs("tblAgent").RefreshLink
    LinkTable "tblAgentHist"
End Sub

' tblAgentHist: ban chup danh sach agent theo tung thang (SNAPSHOT_PERIOD), cung cau truc voi tblAgent
Private Sub EnsureHistTable(ByVal be As DAO.Database)
    Dim td As DAO.TableDef, f As DAO.Field, have As Object, exists As Boolean
    For Each td In be.TableDefs
        If td.Name = "tblAgentHist" Then exists = True
    Next td
    If Not exists Then
        be.Execute "SELECT * INTO tblAgentHist FROM tblAgent WHERE 1=0", dbFailOnError
        be.Execute "CREATE INDEX pkAgentHist ON tblAgentHist (SNAPSHOT_PERIOD, AgentCode) WITH PRIMARY", dbFailOnError
        Exit Sub
    End If
    Set have = NewDict()
    For Each f In be.TableDefs("tblAgentHist").Fields
        have(UCase$(f.Name)) = True
    Next f
    For Each f In be.TableDefs("tblAgent").Fields
        If Not have.Exists(UCase$(f.Name)) Then
            If f.Type = dbDate Then
                be.Execute "ALTER TABLE tblAgentHist ADD COLUMN [" & f.Name & "] DATETIME", dbFailOnError
            Else
                be.Execute "ALTER TABLE tblAgentHist ADD COLUMN [" & f.Name & "] TEXT(255)", dbFailOnError
            End If
        End If
    Next f
End Sub

Private Sub LinkTable(ByVal name As String)
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If td.Name = name Then
            td.RefreshLink
            Exit Sub
        End If
    Next td
    DoCmd.TransferDatabase acLink, "Microsoft Access", BackEndPath(), acTable, name, name
End Sub

' Query hien thi tblAgent theo dung thu tu va tieu de cot cua file he thong
Private Sub BuildAgentQuery()
    Dim spec As Variant, p As Variant, i As Long, sql As String
    spec = AgentSpec()
    For i = 0 To UBound(spec)
        p = Split(spec(i), "|")
        sql = sql & IIf(i = 0, "", ", ") & "[" & p(2) & "] AS [" & p(4) & "]"
    Next i
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "qryAgent"
    On Error GoTo 0
    CurrentDb.CreateQueryDef "qryAgent", "SELECT " & sql & " FROM tblAgent ORDER BY AgentCode"
End Sub

Private Function FormExists(ByVal name As String) As Boolean
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        If ao.Name = name Then FormExists = True
    Next ao
End Function

Private Sub DeleteForm(ByVal name As String)
    If FormExists(name) Then
        If CurrentProject.AllForms(name).IsLoaded Then DoCmd.Close acForm, name
        DoCmd.DeleteObject acForm, name
    End If
End Sub

Private Function SaveNewForm(ByVal f As Form, ByVal name As String) As String
    Dim tmp As String
    tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.Rename name, acForm, tmp
End Function

Private Function AddCtl(ByVal formName As String, ByVal ctlType As Long, ByVal ctlName As String, _
                        ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long, _
                        Optional ByVal section As Long = acDetail) As Control
    Set AddCtl = CreateControl(formName, ctlType, section, , , l, t, w, h)
    AddCtl.Name = ctlName
    If ctlType <> acSubform Then AddCtl.FontName = "Segoe UI"
End Function

' frmAgentList: bang dang datasheet giong file Excel. frmAgent: tim kiem, loc chuc danh, nap/xuat file.
Private Sub BuildAgentForms()
    Dim f As Form, c As Control, lb As Control, spec As Variant, p As Variant, i As Long, fn As String

    DeleteForm "frmAgent"
    DeleteForm "frmAgentList"
    BuildAgentQuery

    Set f = CreateForm()
    fn = f.Name
    f.RecordSource = "qryAgent"
    f.DefaultView = 2
    f.Caption = U("Danh s\u00E1ch Agent")
    spec = AgentSpec()
    For i = 0 To UBound(spec)
        p = Split(spec(i), "|")
        Set c = AddCtl(fn, acTextBox, "c" & i, 2000, 300 * i, 1800, 280)
        c.ControlSource = "[" & p(4) & "]"
        If p(3) = "D" Or p(3) = "CD" Then c.Format = "dd/mm/yyyy"
        Set lb = CreateControl(fn, acLabel, acDetail, c.Name, , 100, 300 * i, 1800, 280)
        lb.Name = "l" & i
        lb.Caption = p(4)
        If p(2) = "AgentName" Or p(3) = "T" And InStr(p(2), "NAME") > 0 Then c.ColumnWidth = 2800 Else c.ColumnWidth = 1700
    Next i
    SaveNewForm f, "frmAgentList"

    Set f = CreateForm()
    fn = f.Name
    f.Caption = U("Danh s\u00E1ch Agent")
    f.RecordSelectors = False
    f.NavigationButtons = False
    f.DividingLines = False
    f.ScrollBars = 0
    f.Width = 16000
    f.Section(acDetail).Height = 9000
    f.OnLoad = "=AgentFilter()"

    Set c = AddCtl(fn, acLabel, "lblSearch", 200, 200, 2400, 320)
    c.Caption = U("T\u00ECm m\u00E3 / h\u1ECD t\u00EAn")
    Set c = AddCtl(fn, acTextBox, "txtSearch", 2600, 200, 2600, 320)
    c.AfterUpdate = "=AgentFilter()"
    Set c = AddCtl(fn, acLabel, "lblTitle", 5500, 200, 1200, 320)
    c.Caption = U("Ch\u1EE9c danh")
    Set c = AddCtl(fn, acComboBox, "cboTitle", 6700, 200, 1300, 320)
    c.RowSourceType = "Table/Query"
    c.RowSource = "SELECT DISTINCT Title FROM tblAgent WHERE Title Is Not Null ORDER BY Title"
    c.AfterUpdate = "=AgentFilter()"
    Set c = AddCtl(fn, acCommandButton, "btnClear", 8200, 180, 1300, 380)
    c.Caption = U("B\u1ECF l\u1ECDc")
    c.OnClick = "=AgentClearFilter()"
    Set c = AddCtl(fn, acCommandButton, "btnImportAgent", 9700, 180, 2600, 380)
    c.Caption = U("N\u1EA1p file Excel")
    c.OnClick = "=BtnImport()"
    Set c = AddCtl(fn, acCommandButton, "btnExportAgent", 12500, 180, 2400, 380)
    c.Caption = U("Xu\u1EA5t Excel")
    c.OnClick = "=AgentExport()"
    Set c = AddCtl(fn, acLabel, "lblCount", 200, 650, 9000, 300)
    c.Caption = " "
    c.ForeColor = RGB(31, 78, 121)

    Set c = AddCtl(fn, acSubform, "sfList", 200, 1000, 15600, 7800)
    c.SourceObject = "Form.frmAgentList"
    c.HorizontalAnchor = 2
    c.VerticalAnchor = 2
    SaveNewForm f, "frmAgent"
End Sub

Public Function AgentFilter()
    Dim f As Form, s As String, t As String, w As String, n As Long
    If Not CurrentProject.AllForms("frmAgent").IsLoaded Then Exit Function
    Set f = Forms("frmAgent")
    s = Replace(Trim$(Nz(f!txtSearch, "")), "'", "''")
    t = Replace(Trim$(Nz(f!cboTitle, "")), "'", "''")
    If Len(s) > 0 Then w = "([AGENT_CODE] Like '*" & s & "*' Or [FULL_NAME] Like '*" & s & "*')"
    If Len(t) > 0 Then w = w & IIf(Len(w) > 0, " AND ", "") & "[TITLE (Code)]='" & t & "'"
    With f!sfList.Form
        .Filter = w
        .FilterOn = (Len(w) > 0)
        With .RecordsetClone
            If Not .EOF Then .MoveLast
            n = .RecordCount
        End With
    End With
    f!lblCount.Caption = U("S\u1ED1 agent: ") & n
End Function

Public Function AgentClearFilter()
    Forms("frmAgent")!txtSearch = Null
    Forms("frmAgent")!cboTitle = Null
    AgentFilter
End Function

Public Function AgentExport()
    Dim path As String
    On Error GoTo Fail
    With Application.FileDialog(2)
        .Title = U("Xu\u1EA5t danh s\u00E1ch Agent")
        .InitialFileName = CurrentProject.Path & "\Agent_" & Format$(Date, "yyyymmdd") & ".xlsx"
        If Not .Show Then Exit Function
        path = .SelectedItems(1)
    End With
    If LCase$(Right$(path, 5)) <> ".xlsx" Then path = path & ".xlsx"
    If Dir$(path) <> "" Then Kill path
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryAgent", path, True
    MsgU U("\u0110\u00E3 xu\u1EA5t:") & vbCrLf & path
    Exit Function
Fail:
    MsgU T("E_Export") & vbCrLf & Err.Description, MB_ICONERROR
End Function

Public Function BtnAgents()
    If Not FormExists("frmAgent") Then
        EnsureAgentSchema
        BuildAgentForms
    End If
    DoCmd.OpenForm "frmAgent"
End Function

Private Sub EnsureMainButton(ByVal ctlName As String, ByVal caption As String, ByVal onClick As String, _
                             ByVal l As Long, ByVal t As Long)
    Dim f As Form, c As Control, found As Boolean
    DoCmd.OpenForm FORM_MAIN, acDesign, , , , acHidden
    Set f = Forms(FORM_MAIN)
    For Each c In f.Controls
        If c.Name = ctlName Then found = True
    Next c
    If Not found Then
        Set c = AddCtl(FORM_MAIN, acCommandButton, ctlName, l, t, 3400, 460)
        c.Caption = caption
        c.OnClick = onClick
    End If
    DoCmd.Close acForm, FORM_MAIN, acSaveYes
End Sub

' Goi sau moi lan cap nhat code (nut "Cap nhat tool") va khi dung tool moi:
' bo sung cot, nhan, form, nut ma phien ban moi can. Khong xoa du lieu.
Public Function FaAfterUpdate()
    Dim i As Long
    For i = Forms.Count - 1 To 0 Step -1
        DoCmd.Close acForm, Forms(i).Name
    Next i
    EnsureTexts
    EnsureAgentSchema
    BuildAgentForms
    EnsureMainButton "btnAgents", U("Danh s\u00E1ch Agent"), "=BtnAgents()", 5000, 4440
End Function

' O Excel loi (#N/A, #VALUE!...) tra ve chuoi rong thay vi gay loi Type mismatch
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = CStr(v)
End Function

Private Function ToDateOrNull(ByVal v As Variant) As Variant
    ToDateOrNull = Null
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        ToDateOrNull = DateValue(v)
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then ToDateOrNull = CDate(CDbl(v))
    ElseIf IsDate(v) Then
        ToDateOrNull = DateValue(CDate(v))
    End If
End Function

Private Function FindCol(ByVal hdr As Object, ParamArray names() As Variant) As Long
    Dim n As Variant
    For Each n In names
        If hdr.Exists(n & "#1") Then
            FindCol = hdr(n & "#1")
            Exit Function
        End If
    Next n
End Function

' Doc sheet Agent theo ten cot, chap nhan ca file mau cua tool (AgentCode, AgentName...)
' lan file xuat tu he thong (AGENT_CODE, FULL_NAME, APPOINTMENT_DATE, AGENCY_TYPE, TITLE (Code)...).
' Ten cot duoc so khop sau khi bo khoang trang va dau gach duoi.
Private Function ImportAgents(ByVal path As String, ByRef nAgents As Long) As Boolean
    Dim xl As Object, wb As Object, ws As Object, data As Variant, hdr As Object
    Dim cCode As Long, cName As Long, cAppoint As Long, cModel As Long, cTerm As Long, cTitle As Long, cStatus As Long
    Dim r As Long, c As Long, key As String, code As String, rs As DAO.Recordset
    Dim seen As Object, spec As Variant, p As Variant, extraCol() As Long
    Dim cPeriod As Long, periods As Object, pk As Variant

    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(path, 0, True)
    On Error Resume Next
    Set ws = wb.Worksheets("Agent")
    On Error GoTo Fail
    If ws Is Nothing Then GoTo Done
    data = ws.UsedRange.Value
    If Not IsArray(data) Then GoTo Done

    Set hdr = NewDict()
    Set seen = NewDict()
    For c = 1 To UBound(data, 2)
        key = NormHeader(data(1, c))
        If Len(key) > 0 Then
            seen(key) = seen(key) + 1
            hdr(key & "#" & seen(key)) = c
        End If
    Next c
    cCode = FindCol(hdr, "AGENTCODE")
    cName = FindCol(hdr, "AGENTNAME", "FULLNAME")
    cAppoint = FindCol(hdr, "APPOINTDATE", "APPOINTMENTDATE")
    cModel = FindCol(hdr, "WORKMODEL", "AGENCYTYPE")
    cTerm = FindCol(hdr, "TERMINATEDATE", "TERMINATIONDATE")
    cTitle = FindCol(hdr, "TITLE", "TITLE(CODE)")
    cStatus = FindCol(hdr, "AGENTSTATUS")
    cPeriod = FindCol(hdr, "PERIOD")
    Set periods = NewDict()
    If cCode = 0 Or cAppoint = 0 Or cModel = 0 Then Err.Raise vbObjectError + 10, , T("E_AgentCols")

    ' Cac cot con lai cua file he thong (CCCD, ngan hang, tuyen quan ly...)
    spec = AgentSpec()
    ReDim extraCol(UBound(spec))
    For c = 0 To UBound(spec)
        p = Split(spec(c), "|")
        If Left$(p(3), 1) <> "C" Then
            key = NormHeader(p(0)) & "#" & p(1)
            If hdr.Exists(key) Then
                extraCol(c) = hdr(key)
            ElseIf hdr.Exists(NormHeader(p(4)) & "#1") Then
                extraCol(c) = hdr(NormHeader(p(4)) & "#1")
            End If
        End If
    Next c

    Set rs = CurrentDb.OpenRecordset("tblAgent", dbOpenDynaset)
    For r = 2 To UBound(data, 1)
        code = Trim$(CellText(data(r, cCode)))
        If Len(code) > 0 Then
            rs.FindFirst "AgentCode=" & Q(code)
            If rs.NoMatch Then
                rs.AddNew
                rs!AgentCode = code
            Else
                rs.Edit
            End If
            If cName > 0 Then rs!AgentName = Left$(Trim$(CellText(data(r, cName))), 100)
            rs!AppointDate = ToDateOrNull(data(r, cAppoint))
            rs!WorkModel = Left$(UCase$(Trim$(CellText(data(r, cModel)))), 2)
            If cTerm > 0 Then rs!TerminateDate = ToDateOrNull(data(r, cTerm))
            If cTitle > 0 Then rs!Title = Left$(UCase$(Trim$(CellText(data(r, cTitle)))), 20)
            If cStatus > 0 Then rs!AgentStatus = Left$(Trim$(CellText(data(r, cStatus))), 20)
            For c = 0 To UBound(spec)
                If extraCol(c) > 0 Then
                    p = Split(spec(c), "|")
                    If p(3) = "D" Then
                        rs.Fields(p(2)) = ToDateOrNull(data(r, extraCol(c)))
                    ElseIf Len(Trim$(CellText(data(r, extraCol(c))))) = 0 Then
                        rs.Fields(p(2)) = Null
                    Else
                        rs.Fields(p(2)) = Left$(Trim$(CellText(data(r, extraCol(c)))), 255)
                    End If
                End If
            Next c
            rs.Update
            nAgents = nAgents + 1
            If cPeriod > 0 Then
                key = Trim$(CellText(data(r, cPeriod)))
                If Len(key) > 0 Then periods(key) = True
            End If
        End If
    Next r
    rs.Close

    ' Luu ban chup tung thang: nap lai cung thang thi ghi de dung thang do
    For Each pk In periods.Keys
        CurrentDb.Execute "DELETE FROM tblAgentHist WHERE SNAPSHOT_PERIOD=" & Q(CStr(pk)), dbFailOnError
        CurrentDb.Execute "INSERT INTO tblAgentHist SELECT * FROM tblAgent WHERE SNAPSHOT_PERIOD=" & Q(CStr(pk)), dbFailOnError
    Next pk
Done:
    ImportAgents = True
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Exit Function
Fail:
    mErr = Err.Description
    Resume Cleanup
End Function

Public Function BtnImport()
    Dim path As String, db As DAO.Database, nIn As Long, locked As Long, nAg As Long
    On Error GoTo Fail

    path = PickFile(T("D_PickExcel"), "Excel", "*.xlsx;*.xlsm;*.xls")
    If path = "" Then Exit Function
    If Not AutoBackup() Then Exit Function
    Set db = CurrentDb
    SetStatus T("S_Importing")

    EnsureTexts
    EnsureAgentSchema
    If Not ImportAgents(path, nAg) Then
        SetStatus T("E_Import")
        MsgU T("E_Import") & vbCrLf & mErr, MB_ICONERROR
        Exit Function
    End If

    If SheetExists(path, "Holiday") Then
        db.Execute "INSERT INTO tblHoliday (HolidayDate, Description) SELECT DateValue(s.HolidayDate), s.Description " & _
            "FROM stg_Holiday AS s WHERE s.HolidayDate Is Not Null AND DateValue(s.HolidayDate) NOT IN (SELECT HolidayDate FROM tblHoliday)", dbFailOnError
        DropTable "stg_Holiday"
    End If

    If SheetExists(path, "FA_Input") Then
        ' Khong ghi de ky da chot
        Dim rsL As DAO.Recordset
        Set rsL = db.OpenRecordset("SELECT Count(*) AS N FROM stg_FA_Input AS s INNER JOIN tblPeriodStatus AS p " & _
            "ON (CStr(s.Period) = p.Period AND CLng(s.Ky) = p.Ky) WHERE p.Status='LOCKED'", dbOpenSnapshot)
        locked = rsL!N
        rsL.Close
        If locked > 0 Then
            DropTable "stg_FA_Input"
            MsgU T("E_ImportLocked"), MB_ICONERROR
            Exit Function
        End If
        db.Execute "DELETE FROM tblFA_Input WHERE EXISTS (SELECT 1 FROM stg_FA_Input AS s " & _
            "WHERE CStr(s.Period) = tblFA_Input.Period AND CLng(s.Ky) = tblFA_Input.Ky AND CStr(s.AgentCode) = tblFA_Input.AgentCode)", dbFailOnError
        db.Execute "INSERT INTO tblFA_Input (Period, Ky, AgentCode, ActualDays, FYC, Contracts, Recruits, K2Pct, MDRTPrev, TrainingDone, StatusEOM) " & _
            "SELECT CStr(s.Period), CLng(s.Ky), CStr(s.AgentCode), Nz(s.ActualDays,0), Nz(s.FYC,0), Nz(s.Contracts,0), Nz(s.Recruits,0), " & _
            "IIf(Trim(s.K2Pct & '')='', Null, Val(s.K2Pct & '')), (Val(s.MDRTPrev & '')=1), (Val(s.TrainingDone & '')=1), s.StatusEOM " & _
            "FROM stg_FA_Input AS s WHERE s.AgentCode Is Not Null", dbFailOnError
        nIn = db.RecordsAffected
        DropTable "stg_FA_Input"
    End If

    If CurrentProject.AllForms("frmAgent").IsLoaded Then
        Forms("frmAgent")!sfList.Form.Requery
        AgentFilter
    End If
    SetStatus T("S_AgentsImported") & " " & nAg & " | " & T("S_Imported") & " " & nIn
    MsgU T("S_AgentsImported") & " " & nAg & vbCrLf & T("S_Imported") & " " & nIn
    Exit Function
Fail:
    SetStatus T("E_Import")
    MsgU T("E_Import") & vbCrLf & Err.Description, MB_ICONERROR
End Function

' =====================================================================
' NUT TREN FORM
' =====================================================================

Private Function ReadForm(ByRef period As String, ByRef ky As Long) As Boolean
    period = Trim$(Nz(Forms(FORM_MAIN)!txtPeriod, ""))
    ky = CLng(Nz(Forms(FORM_MAIN)!cboKy, 0))
    If Not IsValidPeriod(period) Then
        MsgU T("E_Period"), MB_ICONERROR: Exit Function
    End If
    If ky < 1 Or ky > 4 Then
        MsgU T("E_Ky"), MB_ICONERROR: Exit Function
    End If
    ReadForm = True
End Function

Public Function BtnRun()
    Dim period As String, ky As Long
    If Not ReadForm(period, ky) Then Exit Function
    RunFA period, ky
End Function

Private Sub BuildViewQuery(ByVal name As String, ByVal period As String, ByVal ky As Long)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete name
    On Error GoTo 0
    db.CreateQueryDef name, "SELECT * FROM qryFA_BangKe WHERE Period=" & Q(period) & " AND Ky=" & ky & " ORDER BY AgentCode"
End Sub

Public Function BtnView()
    Dim period As String, ky As Long
    If Not ReadForm(period, ky) Then Exit Function
    If DCount("*", "tblFA_Result", "Period=" & Q(period) & " AND Ky=" & ky) = 0 Then
        MsgU T("E_NoResult"), MB_ICONERROR: Exit Function
    End If
    DoCmd.Close acQuery, "qryFA_Xem"
    BuildViewQuery "qryFA_Xem", period, ky
    DoCmd.OpenQuery "qryFA_Xem", acViewNormal, acReadOnly
End Function

Public Function BtnExport()
    Dim period As String, ky As Long, path As String
    On Error GoTo Fail
    If Not ReadForm(period, ky) Then Exit Function
    If DCount("*", "tblFA_Result", "Period=" & Q(period) & " AND Ky=" & ky) = 0 Then
        MsgU T("E_NoResult"), MB_ICONERROR: Exit Function
    End If
    With Application.FileDialog(2)   ' msoFileDialogSaveAs
        .Title = T("D_SaveExcel")
        .InitialFileName = CurrentProject.Path & "\FA_" & period & "_Ky" & ky & ".xlsx"
        If Not .Show Then Exit Function
        path = .SelectedItems(1)
    End With
    If LCase$(Right$(path, 5)) <> ".xlsx" Then path = path & ".xlsx"
    If Dir$(path) <> "" Then Kill path
    BuildViewQuery "qryFA_Export", period, ky
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryFA_Export", path, True
    CurrentDb.QueryDefs.Delete "qryFA_Export"
    SetStatus T("S_Exported") & " " & path
    MsgU T("S_Exported") & vbCrLf & path
    Exit Function
Fail:
    MsgU T("E_Export") & vbCrLf & Err.Description, MB_ICONERROR
End Function

Public Function BtnLock()
    Dim period As String, ky As Long, st As String
    If Not ReadForm(period, ky) Then Exit Function
    st = GetStatus(period, ky)
    If st = "LOCKED" Then
        MsgU T("E_Locked"), MB_ICONERROR: Exit Function
    End If
    If st <> "CALC" Then
        MsgU T("E_LockNotCalc"), MB_ICONERROR: Exit Function
    End If
    If MsgU(T("Q_Lock") & " " & period & " / " & ky & "?", MB_YESNO Or MB_ICONQUESTION) <> IDYES Then Exit Function
    If Not AutoBackup() Then Exit Function
    SetPeriodStatus period, ky, "LOCKED"
    SetStatus T("S_Locked") & " " & period & " / " & ky
End Function

Public Function BtnParams()
    DoCmd.OpenTable "tblParam"
    DoCmd.OpenTable "tblParam_Allowance"
    DoCmd.OpenTable "tblParam_K2"
    DoCmd.OpenTable "tblParam_Contract"
End Function

' Chon lai file Data (FA_Data.accdb) khi chuyen len thu muc chung
Public Function BtnRelink()
    Dim path As String, td As DAO.TableDef, db As DAO.Database, n As Long
    On Error GoTo Fail
    path = PickFile(T("D_PickData"), "Access", "*.accdb")
    If path = "" Then Exit Function
    Set db = CurrentDb
    For Each td In db.TableDefs
        If Len(td.Connect) > 0 And Left$(td.Name, 4) <> "MSys" Then
            td.Connect = ";DATABASE=" & path
            td.RefreshLink
            n = n + 1
        End If
    Next td
    SetStatus T("S_Relinked") & " " & path
    MsgU T("S_Relinked") & " " & n & vbCrLf & path
    Exit Function
Fail:
    MsgU T("E_Relink") & vbCrLf & Err.Description, MB_ICONERROR
End Function
