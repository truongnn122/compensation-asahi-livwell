"""Sinh modFAPM.bas (FA cho PM/PD - ALL/AD/2210-006 muc A.I) tu modFA.bas (FA cho LP).

Phan dung chung (nap agent/cham cong, form Danh sach Agent, chot ky, cap nhat tool...) giu nguyen;
chi thay phan tham so, cong thuc, bang nhap tay, bang ket qua va bang ke.
Chay: python3 gen_modFAPM.py   (moi khi modFA.bas thay doi)
"""

import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE.parents[1] / "access-fa" / "modFA.bas"
DST = HERE.parent / "modFAPM.bas"


def esc(s):
    return "".join(c if ord(c) < 128 else "\\u%04X" % ord(c) for c in s)


def U(s):
    return 'U("' + esc(s) + '")'


def span(s, header):
    i = s.index(header)
    kind = "Function" if "Function " in header else "Sub"
    j = s.index("End " + kind + "\n", i) + len("End " + kind + "\n")
    return i, j


def replace_proc(s, header, new):
    i, j = span(s, header)
    return s[:i] + new + s[j:]


def remove_proc(s, header, comment_lines=0):
    i, j = span(s, header)
    k = i
    for _ in range(comment_lines):
        k = s.rindex("\n", 0, k - 1) + 1
    if s[j:j + 1] == "\n":
        j += 1
    return s[:k] + s[j:]


def must_replace(s, old, new, count=1):
    assert old in s, old[:80]
    return s.replace(old, new, count)


s = SRC.read_text(encoding="ascii").replace("\r\n", "\n")

# ---- doi ten bang / query rieng cua khoan thuong ----
for a, b in [("tblFA_Input", "tblFAPM_Input"), ("tblFA_Result", "tblFAPM_Result"), ("qryFA_BangKe", "qryFAPM_BangKe"),
             ("qryFA_Xem", "qryFAPM_Xem"), ("qryFA_Export", "qryFAPM_Export"), ('"\\FA_" & period', '"\\FAPM_" & period'),
             ("tblParam_Contract", "tblParam_LPCount"), ('"FA_Input"', '"FAPM_Input"'), ("stg_FA_Input", "stg_FAPM_Input")]:
    s = s.replace(a, b)

# ---- dau module ----
s = must_replace(s, "' TINH PHU CAP CO DINH (FA) - ALL/AD/2210-005, Phu luc 1, Muc A.II\n' Logic doi chieu: kiem-tra/fa_reference.py",
                 "' TINH PHU CAP CO DINH (FA) CHO PM/PD - ALL/AD/2210-006, Phu luc 1, Muc A.I\n"
                 "' File nay duoc SINH TU DONG tu access-fa/modFA.bas boi src/gen_modFAPM.py - sua o do roi sinh lai.\n"
                 "' Logic doi chieu: kiem-tra/fapm_reference.py")
s = re.sub(r'Private Const FA_VERSION As String = "[^"]+"', 'Private Const FA_VERSION As String = "2026.11.01-1"', s)
s = must_replace(s, 'Private Const FA_TITLES As String = ",LP,"    \' Chuc danh duoc nhan FA (Muc A.II: LP toan thoi gian)',
                 'Private Const FA_TITLES As String = ",PM,PD,"    \' Chuc danh duoc nhan FA nhom truc tiep (Muc A)')
s = must_replace(s, "Public Function FaVersion() As String\n    FaVersion = FA_VERSION\nEnd Function\n",
                 "Public Function FaVersion() As String\n    FaVersion = FA_VERSION\nEnd Function\n\n"
                 "' Ten module de nut 'Cap nhat tool' (modUpdate) biet nap file nao\n"
                 "Public Function FaModuleName() As String\n    FaModuleName = \"modFAPM\"\nEnd Function\n")

# ---- tham so ----
s = must_replace(s, """Private pStdDays As Double
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
""", """Private pStdDays As Double
Private pMinAttPct As Double       ' 0 = khong ap dung dieu kien cham cong toi thieu
Private pMinDays As Double
Private pNewGuarAllow As Double    ' PM <= 12 thang ke tu thang tien, quy xet >= 8 luot LP hoat dong
Private pNewGuarMonths As Long
Private pNewGuarMinTurns As Double
Private pQtrGuarAllow As Double    ' PM khac, quy xet >= 6 luot LP hoat dong
Private pQtrMinTurns As Double
Private pJoinGuarAllow As Double   ' PM gia nhap den het Q2/2027: 6 thang dau
Private pJoinGuarMonths As Long
Private pJoinDeadline As String
Private pK2NACoef As Double

Private allowRows As Variant   ' (i, 0)=MinGroupFYC, (i,1)=MinActiveLP, (i,2)=AllowancePerDay
Private k2Rows As Variant      ' (i, 0)=MinK2, (i,1)=Coef - sap xep giam dan
Private lpRows As Variant      ' (i, 0)=MinTurns, (i,1)=Coef - sap xep giam dan
""")

s = replace_proc(s, "Private Sub LoadParams()", """Private Sub LoadParams()
    Dim rs As DAO.Recordset

    pStdDays = P("StdDays")
    pMinAttPct = P("MinAttendancePct")
    pMinDays = P("MinWorkDays")
    pNewGuarAllow = P("NewGuarAllowance")
    pNewGuarMonths = CLng(P("NewGuarMonths"))
    pNewGuarMinTurns = P("NewGuarMinLPTurns")
    pQtrGuarAllow = P("QtrGuarAllowance")
    pQtrMinTurns = P("QtrGuarMinLPTurns")
    pJoinGuarAllow = P("JoinGuarAllowance")
    pJoinGuarMonths = CLng(P("JoinGuarMonths"))
    pJoinDeadline = Format$(P("JoinGuarDeadline"), "0")
    pK2NACoef = P("K2_NA_Coef")

    allowRows = LoadRows("SELECT MinGroupFYC, MinActiveLP, AllowancePerDay FROM tblParam_Allowance", 3)
    k2Rows = LoadRows("SELECT MinK2, Coef FROM tblParam_K2 ORDER BY MinK2 DESC", 2)
    lpRows = LoadRows("SELECT MinTurns, Coef FROM tblParam_LPCount ORDER BY MinTurns DESC", 2)

    Set holidays = NewDict()
    Set rs = CurrentDb.OpenRecordset("SELECT HolidayDate FROM tblHoliday", dbOpenSnapshot)
    Do Until rs.EOF
        holidays(CLng(DateValue(rs!HolidayDate))) = True
        rs.MoveNext
    Loop
    rs.Close
End Sub
""")

s = replace_proc(s, "Private Function AllowanceFor(", """Private Function AllowanceFor(ByVal groupFyc As Double, ByVal activeLP As Double) As Double
    Dim i As Long, best As Double
    For i = 0 To UBound(allowRows, 1)
        If groupFyc >= allowRows(i, 0) And activeLP >= allowRows(i, 1) Then
            If allowRows(i, 2) > best Then best = allowRows(i, 2)
        End If
    Next i
    AllowanceFor = best
End Function
""")
s = must_replace(s, "' Muc A: lay muc phu cap cao nhat ma LP thoa (Lua chon 1 hoac Lua chon 2)",
                 "' Phu cap/ngay: muc cao nhat ma nhom truc tiep thoa (FYC nhom va so LP hoat dong trong thang)")
s = must_replace(s, "' Muc B: %He so K2. K2 = Null nghia la N/A (chua co K2)", "' %He so K2 nhom truc tiep. K2 = Null nghia la N/A (chua co K2)")
s = replace_proc(s, "Private Function ContractCoefFor(", """Private Function LPCountCoefFor(ByVal turns As Double) As Double
    Dim i As Long
    For i = 0 To UBound(lpRows, 1)
        If turns >= lpRows(i, 0) Then
            LPCountCoefFor = lpRows(i, 1)
            Exit Function
        End If
    Next i
End Function
""")
s = must_replace(s, "' Muc D: %He so hop dong", "' %He so Luot LP nhom truc tiep hoat dong (cong don ke tu khi thang tien/bo nhiem PM)")

# ---- du lieu cac thang truoc ----
s = replace_proc(s, "Private Function LoadLatestResults(", """Private Function LoadLatestResults(ByVal period As String) As Object
    Dim d As Object, rs As DAO.Recordset
    Set d = NewDict()
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT r.AgentCode, r.FA_Month, r.QtrGuarMonth " & _
        "FROM tblFAPM_Result AS r WHERE r.Period=" & Q(period) & " AND r.Ky = " & _
        "(SELECT Max(x.Ky) FROM tblFAPM_Result AS x WHERE x.Period=r.Period AND x.AgentCode=r.AgentCode)", dbOpenSnapshot)
    Do Until rs.EOF
        d(CStr(rs!AgentCode)) = Array(Nz(rs!FA_Month, 0), Nz(rs!QtrGuarMonth, 0))
        rs.MoveNext
    Loop
    rs.Close
    Set LoadLatestResults = d
End Function
""")
s = must_replace(s, "' Ket qua cua 1 thang, lay theo ky chi tra lon nhat da tinh cua tung LP", "' Ket qua cua 1 thang, lay theo ky chi tra lon nhat da tinh cua tung PM")
s = replace_proc(s, "Private Function LoadQtrQualified(", """Private Function LoadQtrTurns(ByVal qtrEnd As String) As Object
    Dim d As Object, rs As DAO.Recordset, inList As String
    Set d = NewDict()
    inList = Q(PeriodAdd(qtrEnd, -2)) & "," & Q(PeriodAdd(qtrEnd, -1)) & "," & Q(qtrEnd)
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT i.AgentCode, Sum(i.ActiveLP) AS SumLP " & _
        "FROM tblFAPM_Input AS i WHERE i.Period IN (" & inList & ") AND i.Ky = " & _
        "(SELECT Max(x.Ky) FROM tblFAPM_Input AS x WHERE x.Period=i.Period AND x.AgentCode=i.AgentCode) " & _
        "GROUP BY i.AgentCode", dbOpenSnapshot)
    Do Until rs.EOF
        d(CStr(rs!AgentCode)) = Nz(rs!SumLP, 0)
        rs.MoveNext
    Loop
    rs.Close
    Set LoadQtrTurns = d
End Function

' So LP hoat dong tung thang truoc thang tinh (ky moi nhat cua moi thang): code -> Dictionary(period -> ActiveLP)
Private Function LoadPriorTurns(ByVal period As String) As Object
    Dim d As Object, rs As DAO.Recordset, code As String
    Set d = NewDict()
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT i.AgentCode, i.Period, i.ActiveLP FROM tblFAPM_Input AS i WHERE i.Period < " & Q(period) & " AND i.Ky = " & _
        "(SELECT Max(x.Ky) FROM tblFAPM_Input AS x WHERE x.Period=i.Period AND x.AgentCode=i.AgentCode)", dbOpenSnapshot)
    Do Until rs.EOF
        code = CStr(rs!AgentCode)
        If Not d.Exists(code) Then Set d(code) = NewDict()
        d(code)(CStr(rs!Period)) = Nz(rs!ActiveLP, 0)
        rs.MoveNext
    Loop
    rs.Close
    Set LoadPriorTurns = d
End Function
""")
s = must_replace(s, "' LP dat dieu kien bao dam FA quy ke tiep: >= 6 HD va >= 24 trieu FYC phat sinh trong quy",
                 "' Tong luot LP hoat dong trong quy xet (3 thang) cua tung PM")

# ---- RunFA ----
s = replace_proc(s, "Public Function RunFA(", """Public Function RunFA(ByVal period As String, ByVal ky As Long) As Boolean
    Dim db As DAO.Database, ws As DAO.Workspace
    Dim rsIn As DAO.Recordset, rsOut As DAO.Recordset
    Dim st As String, stUser As String, stAt As Variant
    Dim priorTurns As Object, qtrTurns As Object, paidPrev As Object, m1Qtr As Object, m2Qtr As Object
    Dim monthStart As Date, monthEnd As Date, monthInQtr As Long, prevQtrEnd As String, titleList As String
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
    If DCount("*", "tblFAPM_Result", "Period=" & Q(period) & " AND Ky>" & ky) > 0 Then
        MsgU T("E_LaterKy"), MB_ICONERROR: Exit Function
    End If
    EnsureTexts
    EnsureDataSchema
    If DCount("*", "Tbl_Agent", "[Period]<=" & Q(period)) = 0 Then
        MsgU """ + U("Chưa có danh sách agent (Tbl_Agent) cho tháng này hoặc trước đó. Nạp file Tbl_Agent trước.") + """, MB_ICONERROR: Exit Function
    End If
    If DCount("*", "tblFAPM_Result", "Period=" & Q(period) & " AND Ky=" & ky) > 0 Then
        If MsgU(T("Q_Recalc"), MB_YESNO Or MB_ICONQUESTION) <> IDYES Then Exit Function
    End If

    SetPeriodStatus period, ky, "RUNNING"
    SetStatus T("S_Loading")
    LoadParams

    monthStart = PeriodStart(period)
    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
    monthInQtr = ((Month(monthStart) - 1) Mod 3) + 1
    prevQtrEnd = PeriodAdd(period, -monthInQtr)

    Set priorTurns = LoadPriorTurns(period)
    Set qtrTurns = LoadQtrTurns(prevQtrEnd)
    Set paidPrev = LoadPaidPrevKy(period, ky)
    If monthInQtr = 3 And ky = 4 Then
        Set m1Qtr = LoadLatestResults(PeriodAdd(period, -2))
        Set m2Qtr = LoadLatestResults(PeriodAdd(period, -1))
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True

    db.Execute "DELETE FROM tblFAPM_Result WHERE Period=" & Q(period) & " AND Ky=" & ky, dbFailOnError

    ' Agent: ban chup cua thang tinh, neu chua co thi ban gan nhat truoc do (Data chung); chi lay PM/PD
    ' Cham cong: tong so cong trong thang (Data chung). FYC nhom, so LP hoat dong, K2: tblFAPM_Input
    Dim sqlA As String, sqlT As String, sqlI As String
    titleList = "'" & Replace(Mid$(FA_TITLES, 2, Len(FA_TITLES) - 2), ",", "','") & "'"
    sqlA = "SELECT x.* FROM Tbl_Agent AS x WHERE x.[Period] = (SELECT Max(y.[Period]) FROM Tbl_Agent AS y " & _
           "WHERE y.[AGENT_CODE] = x.[AGENT_CODE] AND y.[Period] <= " & Q(period) & ")"
    sqlT = "SELECT [" & AttCol("CODE") & "] AS TCode, Count(*) AS TRows, " & _
           "Sum(IIf([" & AttCol("CONG") & "] Is Null, Nz([" & AttCol("AM") & "],0) + Nz([" & AttCol("PM") & "],0), [" & AttCol("CONG") & "])) AS TDays " & _
           "FROM Tbl_Attendance WHERE [Period]=" & Q(period) & " GROUP BY [" & AttCol("CODE") & "]"
    sqlI = "SELECT * FROM tblFAPM_Input WHERE Period=" & Q(period) & " AND Ky=" & ky
    Set rsIn = db.OpenRecordset( _
        "SELECT a.[AGENT_CODE] AS AgentCode, a.[FULL_NAME] AS AgentName, a.[Period] AS AgentPeriod, " & _
        "a.[APPOINTMENT_DATE] AS AppointDate, a.[PROMOTION_DATE] AS PromoDate, a.[AGENCY_TYPE] AS WorkModel, " & _
        "a.[TERMINATION_DATE] AS TerminateDate, a.[TITLE (Code)] AS Title, a.[AGENT_STATUS] AS AgentStatus, " & _
        "t.TDays, t.TRows, i.AgentCode AS InCode, i.GroupFYC, i.ActiveLP, i.K2Pct, i.StatusEOM " & _
        "FROM ((" & sqlA & ") AS a LEFT JOIN (" & sqlT & ") AS t ON a.[AGENT_CODE] = t.TCode) " & _
        "LEFT JOIN (" & sqlI & ") AS i ON a.[AGENT_CODE] = i.AgentCode " & _
        "WHERE UCase(Trim(a.[TITLE (Code)] & '')) IN (" & titleList & ") OR i.AgentCode Is Not Null " & _
        "ORDER BY a.[AGENT_CODE]", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset("tblFAPM_Result", dbOpenDynaset, dbAppendOnly)

    Do Until rsIn.EOF
        Dim code As String, reason As String, eligible As Boolean, title As String
        Dim wm As Long, pmMonths As Long, baseDays As Long, actual As Double, att As Double, attPct As Double, dayF As Double
        Dim gFyc As Double, activeLP As Double, cumTurns As Double, allow As Double, k2c As Double, lpc As Double
        Dim faFormula As Double, joinGuar As Double, faMonth As Double
        Dim turns As Double, level As Double, qtrMonth As Double, topup As Double, reviewMonths As Long
        Dim faTotal As Double, paid As Double, promo As Variant, promoPeriod As String, pk As Variant
        Dim hasInput As Boolean, statusEOM As String

        code = CStr(rsIn!AgentCode)
        reason = ""
        actual = Nz(rsIn!TDays, 0)
        hasInput = Not IsNull(rsIn!InCode)
        If hasInput Then statusEOM = Nz(rsIn!StatusEOM, Nz(rsIn!AgentStatus, "")) Else statusEOM = Nz(rsIn!AgentStatus, "")
        gFyc = Nz(rsIn!GroupFYC, 0)
        activeLP = Nz(rsIn!ActiveLP, 0)
        promo = rsIn!PromoDate
        If IsNull(promo) Then promo = rsIn!AppointDate
        joinGuar = 0: level = 0: qtrMonth = 0: topup = 0: cumTurns = 0
        wm = 0: pmMonths = 0: baseDays = 0: attPct = 0: dayF = 0: promoPeriod = ""

        ' ---- Thang lam viec & ty le cham cong (005 Muc D.k, D.z) ----
        If IsNull(rsIn!AppointDate) Then
            reason = T("R_NoAgent")
        ElseIf rsIn!AppointDate > monthEnd Then
            reason = T("R_NotAppointed")
        Else
            wm = DateDiff("m", rsIn!AppointDate, monthStart) + 1
            pmMonths = DateDiff("m", promo, monthStart) + 1
            promoPeriod = Format$(promo, "yyyymm")
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

        ' ---- Dieu kien ----
        If reason = "" Then
            title = UCase$(Trim$(Nz(rsIn!Title, "")))
            If InStr(FA_TITLES, "," & title & ",") = 0 Then
                reason = T("R_NotPM")
            ElseIf Not IsNull(rsIn!TerminateDate) Then
                If rsIn!TerminateDate <= monthEnd Then reason = T("R_Terminated")
            End If
        End If
        If reason = "" Then
            If UCase$(Trim$(statusEOM)) = "SE" Then
                reason = T("R_SE")
            ElseIf attPct < pMinAttPct And actual < pMinDays Then
                reason = T("R_Attendance")
            End If
        End If
        eligible = (reason = "")

        ' ---- Luot LP hoat dong cong don ke tu thang thang tien ----
        cumTurns = activeLP
        If priorTurns.Exists(code) And Len(promoPeriod) > 0 Then
            For Each pk In priorTurns(code).Keys
                If CStr(pk) >= promoPeriod And CStr(pk) < period Then cumTurns = cumTurns + priorTurns(code)(pk)
            Next pk
        End If

        ' ---- Cong thuc FA ----
        allow = AllowanceFor(gFyc, activeLP)
        k2c = K2CoefFor(rsIn!K2Pct)
        lpc = LPCountCoefFor(cumTurns)
        faFormula = RoundHU(allow * dayF * k2c * lpc, 0)

        ' ---- Bao dam 6 thang dau cho PM gia nhap den het Q2/2027 ----
        If eligible And Len(promoPeriod) > 0 Then
            If promoPeriod <= pJoinDeadline And pmMonths >= 1 And pmMonths <= pJoinGuarMonths Then joinGuar = RoundHU(pJoinGuarAllow * dayF, 0)
        End If
        If eligible Then faMonth = MaxD(faFormula, joinGuar) Else faMonth = 0

        ' ---- Bao dam theo quy: xet luot LP hoat dong quy truoc va tham nien PM tai cuoi quy do ----
        If qtrTurns.Exists(code) Then turns = qtrTurns(code) Else turns = 0
        If Len(promoPeriod) > 0 Then
            reviewMonths = DateDiff("m", promo, PeriodStart(prevQtrEnd)) + 1
            If reviewMonths <= pNewGuarMonths And turns >= pNewGuarMinTurns Then
                level = pNewGuarAllow
            ElseIf turns >= pQtrMinTurns Then
                level = pQtrGuarAllow
            End If
        End If
        If level > 0 And eligible Then qtrMonth = RoundHU(level * dayF, 0)
        If level > 0 And monthInQtr = 3 And ky = 4 Then
            Dim sumGuar As Double, sumFA As Double
            sumGuar = qtrMonth: sumFA = faMonth
            If m1Qtr.Exists(code) Then sumGuar = sumGuar + m1Qtr(code)(1): sumFA = sumFA + m1Qtr(code)(0)
            If m2Qtr.Exists(code) Then sumGuar = sumGuar + m2Qtr(code)(1): sumFA = sumFA + m2Qtr(code)(0)
            topup = MaxD(0, sumGuar - sumFA)
        End If

        faTotal = faMonth + topup
        If paidPrev.Exists(code) Then paid = paidPrev(code) Else paid = 0

        ' ---- Ghi ket qua ----
        rsOut.AddNew
        rsOut!Period = period
        rsOut!Ky = ky
        rsOut!AgentCode = code
        rsOut!WorkMonthNo = wm
        rsOut!PromoMonthNo = pmMonths
        rsOut!WorkDaysBase = baseDays
        rsOut!ActualDays = actual
        rsOut!AttendancePct = attPct
        rsOut!DayFactor = dayF
        rsOut!Eligible = eligible
        rsOut!Reason = reason
        rsOut!GroupFYC = gFyc
        rsOut!ActiveLP = activeLP
        rsOut!CumLPTurns = cumTurns
        rsOut!AllowancePerDay = allow
        rsOut!K2Pct = rsIn!K2Pct
        rsOut!K2Coef = k2c
        rsOut!LPCountCoef = lpc
        rsOut!FA_Formula = faFormula
        rsOut!JoinGuar = joinGuar
        rsOut!FA_Month = faMonth
        rsOut!QtrLPTurns = turns
        rsOut!QtrGuarLevel = level
        rsOut!QtrGuarMonth = qtrMonth
        rsOut!QtrGuarTopup = topup
        rsOut!FA_Total = faTotal
        rsOut!PaidPrevKy = paid
        rsOut!Payable = faTotal - paid
        rsOut!CalcAt = Now()
        rsOut!CalcBy = GetUserId()
        rsOut!AgentName = rsIn!AgentName
        rsOut!Title = rsIn!Title
        rsOut!WorkModel = rsIn!WorkModel
        rsOut!AppointDate = rsIn!AppointDate
        rsOut!PromoDate = promo
        rsOut!AgentPeriod = rsIn!AgentPeriod
        rsOut!AttRows = Nz(rsIn!TRows, 0)
        If hasInput Then rsOut!InputSource = "FAPM_Input" Else rsOut!InputSource = "Data chung"
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

    SetStatus T("S_Done") & " " & n & " PM/PD | " & Format$(total, "#,##0")
    MsgU T("S_Done") & " " & n & " PM/PD" & vbCrLf & T("S_TotalPayable") & " " & Format$(total, "#,##0")
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
""")

# ---- nhan ----
s = replace_proc(s, "Private Sub EnsureTexts()", "Private Sub EnsureTexts()\n"
                 f'    EnsureText "R_NotPM", {U("Chức danh không phải PM/PD (FA nhóm trực tiếp chỉ áp dụng PM/PD)")}\n'
                 f'    EnsureText "R_Attendance", {U("Không đạt chấm công tối thiểu")}\n'
                 "End Sub\n")

# ---- bang ket qua: cot bo sung cho ban cu ----
s = replace_proc(s, "Private Sub EnsureResultColumns()", """Private Sub EnsureResultColumns()
    Dim cols As Variant, c As Variant
    If LocalConnect("tblFAPM_Result") = "-" Then Exit Sub
    For Each c In Array("PromoMonthNo|INTEGER", "GroupFYC|CURRENCY", "ActiveLP|DOUBLE", "CumLPTurns|DOUBLE", _
                        "LPCountCoef|DOUBLE", "JoinGuar|CURRENCY", "QtrLPTurns|DOUBLE", "QtrGuarLevel|CURRENCY", _
                        "AgentName|TEXT(255)", "Title|TEXT(20)", "WorkModel|TEXT(10)", "AppointDate|DATETIME", _
                        "PromoDate|DATETIME", "AgentPeriod|TEXT(6)", "AttRows|INTEGER", "InputSource|TEXT(20)")
        cols = Split(c, "|")
        If Not LocalFieldExists("tblFAPM_Result", cols(0)) Then
            CurrentDb.Execute "ALTER TABLE tblFAPM_Result ADD COLUMN [" & cols(0) & "] " & cols(1), dbFailOnError
            CurrentDb.TableDefs.Refresh
        End If
    Next c
End Sub
""")

# ---- bang ke ----
bk = [("r.Period", None), ("r.Ky", None), ("r.AgentCode", None), ("r.AgentName", "Họ tên"), ("r.Title", "Chức danh"),
      ("r.PromoDate", "Ngày thăng tiến/bổ nhiệm"), ("r.PromoMonthNo", "Tháng thứ (kể từ thăng tiến)"),
      ("r.WorkMonthNo", "Tháng làm việc thứ"), ("r.WorkDaysBase", "Số ngày làm việc của tháng"),
      ("r.ActualDays", "Số ngày đi làm thực tế"), ("r.AttendancePct", "Tỷ lệ chấm công"), ("r.DayFactor", "Hệ số ngày làm việc"),
      ("r.GroupFYC", "FYC nhóm trực tiếp"), ("r.ActiveLP", "Số LP hoạt động trong tháng"), ("r.AllowancePerDay", "Phụ cấp mỗi ngày"),
      ("r.K2Pct", "K2 nhóm"), ("r.K2Coef", "Hệ số K2"), ("r.CumLPTurns", "Lượt LP hoạt động lũy kế"), ("r.LPCountCoef", "Hệ số lượt LP"),
      ("r.FA_Formula", "FA theo công thức"), ("r.JoinGuar", "Bảo đảm 6 tháng đầu"), ("r.FA_Month", "FA tháng"),
      ("r.QtrLPTurns", "Lượt LP hoạt động quý trước"), ("r.QtrGuarLevel", "Mức bảo đảm quý (đ/ngày)"),
      ("r.QtrGuarMonth", "Mức bảo đảm quý của tháng"), ("r.QtrGuarTopup", "Bù bảo đảm quý"), ("r.FA_Total", "Tổng FA lũy kế"),
      ("r.PaidPrevKy", "Đã chi các kỳ trước"), ("r.Payable", "Chi kỳ này"), ("r.Eligible", "Đủ điều kiện"), ("r.Reason", "Ghi chú"),
      ("r.CalcAt", "Thời điểm tính"), ("r.CalcBy", "Người tính"), ("r.AppointDate", "Ngày bổ nhiệm"),
      ("r.AgentPeriod", "Tháng dữ liệu agent"), ("r.AttRows", "Số dòng chấm công"), ("r.InputSource", "Nguồn số liệu")]
lines = "".join(f'    s = s & {U(("" if i == 0 else ", ") + f + (f" AS [{a}]" if a else ""))}\n' for i, (f, a) in enumerate(bk))
s = replace_proc(s, "Private Sub BuildBangKeQuery()", "Private Sub BuildBangKeQuery()\n    Dim s As String\n    s = \"SELECT \"\n" + lines +
                 '    s = s & " FROM tblFAPM_Result AS r"\n    On Error Resume Next\n    CurrentDb.QueryDefs.Delete "qryFAPM_BangKe"\n'
                 '    On Error GoTo 0\n    CurrentDb.CreateQueryDef "qryFAPM_BangKe", s\nEnd Sub\n')

# ---- nhap sheet FAPM_Input ----
i = s.index('        db.Execute "DELETE FROM tblFAPM_Input WHERE EXISTS')
j = s.index('        nIn = db.RecordsAffected', i)
s = s[:i] + '''        db.Execute "DELETE FROM tblFAPM_Input WHERE EXISTS (SELECT 1 FROM stg_FAPM_Input AS s " & _
            "WHERE CStr(s.Period) = tblFAPM_Input.Period AND CLng(s.Ky) = tblFAPM_Input.Ky AND CStr(s.AgentCode) = tblFAPM_Input.AgentCode)", dbFailOnError
        db.Execute "INSERT INTO tblFAPM_Input (Period, Ky, AgentCode, GroupFYC, ActiveLP, K2Pct, StatusEOM) " & _
            "SELECT CStr(s.Period), CLng(s.Ky), CStr(s.AgentCode), Nz(s.GroupFYC,0), Nz(s.ActiveLP,0), " & _
            "IIf(Trim(s.K2Pct & '')='', Null, Val(s.K2Pct & '')), s.StatusEOM " & _
            "FROM stg_FAPM_Input AS s WHERE s.AgentCode Is Not Null", dbFailOnError
''' + s[j:]
s = must_replace(s, "' NHAP DU LIEU TU EXCEL (sheet Agent, Holiday, FAPM_Input)" if "(sheet Agent, Holiday, FAPM_Input)" in s
                 else "' NHAP DU LIEU TU EXCEL (sheet Agent, Holiday, FA_Input)", "' NHAP DU LIEU TU EXCEL (sheet Agent, Holiday, FAPM_Input)")

# ---- tham so: nut mo bang ----
s = replace_proc(s, "Public Function BtnParams()", 'Public Function BtnParams()\n    DoCmd.OpenTable "tblParam"\n'
                 '    DoCmd.OpenTable "tblParam_Allowance"\n    DoCmd.OpenTable "tblParam_K2"\n    DoCmd.OpenTable "tblParam_LPCount"\nEnd Function\n')

# ---- kiem tra ----
assert all(ord(c) < 128 for c in s), "con ky tu ngoai ASCII"
for bad in ["ContractCoefFor", "MDRT", "Recruits", "pQtrMinCt", "pNewGuarMinCt", "LoadQtrQualified", "R_NotLP", "R_NotFT", "R_Training", "TrainingDone"]:
    assert bad not in s, bad
defs = set(re.findall(r"(?:Function|Sub) ([A-Za-z_]+)\(", s))
for fn in ["LPCountCoefFor", "LoadQtrTurns", "LoadPriorTurns", "FaModuleName", "RunFA", "EnsureResultColumns"]:
    assert fn in defs, fn
procs_open = len(re.findall(r"^\s*(Public |Private )?(Function|Sub) ", s, re.M))
procs_close = len(re.findall(r"^\s*End (Function|Sub)\b", s, re.M))
assert procs_open == procs_close, (procs_open, procs_close)
blk = 0
for ln in s.splitlines():
    t = ln.strip()
    if re.match(r"^If\b.*\bThen$", t):
        blk += 1
    if t == "End If":
        blk -= 1
assert blk == 0, "If/End If lech"
DST.write_bytes(s.replace("\n", "\r\n").encode("ascii"))
print("OK", DST, len(s.splitlines()), "dong")
