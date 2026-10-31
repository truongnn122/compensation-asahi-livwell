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

Private Const FA_VERSION As String = "2026.10.31-2"
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
            out = out & ChrW$(CLng("&H" & Mid$(s, i + 2, 4) & "&"))
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
    EnsureAgentColumns
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
        "SELECT i.*, a.AppointDate, a.WorkModel, a.TerminateDate, a.Title FROM tblFA_Input AS i " & _
        "LEFT JOIN tblAgent AS a ON i.AgentCode = a.AgentCode " & _
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
        If hdr.Exists(n) Then
            FindCol = hdr(n)
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
    For c = 1 To UBound(data, 2)
        key = UCase$(Replace(Replace(Trim$(data(1, c) & ""), " ", ""), "_", ""))
        If Len(key) > 0 And Not hdr.Exists(key) Then hdr(key) = c
    Next c
    cCode = FindCol(hdr, "AGENTCODE")
    cName = FindCol(hdr, "AGENTNAME", "FULLNAME")
    cAppoint = FindCol(hdr, "APPOINTDATE", "APPOINTMENTDATE")
    cModel = FindCol(hdr, "WORKMODEL", "AGENCYTYPE")
    cTerm = FindCol(hdr, "TERMINATEDATE", "TERMINATIONDATE")
    cTitle = FindCol(hdr, "TITLE", "TITLE(CODE)")
    cStatus = FindCol(hdr, "AGENTSTATUS")
    If cCode = 0 Or cAppoint = 0 Or cModel = 0 Then Err.Raise vbObjectError + 10, , T("E_AgentCols")

    Set rs = CurrentDb.OpenRecordset("tblAgent", dbOpenDynaset)
    For r = 2 To UBound(data, 1)
        code = Trim$(data(r, cCode) & "")
        If Len(code) > 0 Then
            rs.FindFirst "AgentCode=" & Q(code)
            If rs.NoMatch Then
                rs.AddNew
                rs!AgentCode = code
            Else
                rs.Edit
            End If
            If cName > 0 Then rs!AgentName = Left$(Trim$(data(r, cName) & ""), 100)
            rs!AppointDate = ToDateOrNull(data(r, cAppoint))
            rs!WorkModel = Left$(UCase$(Trim$(data(r, cModel) & "")), 2)
            If cTerm > 0 Then rs!TerminateDate = ToDateOrNull(data(r, cTerm))
            If cTitle > 0 Then rs!Title = Left$(UCase$(Trim$(data(r, cTitle) & "")), 20)
            If cStatus > 0 Then rs!AgentStatus = Left$(Trim$(data(r, cStatus) & ""), 20)
            rs.Update
            nAgents = nAgents + 1
        End If
    Next r
    rs.Close
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
    EnsureAgentColumns
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
