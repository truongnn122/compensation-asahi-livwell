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

Private Const FA_VERSION As String = "2026.11.01-5"
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

' Ten module de nut 'Cap nhat tool' (modUpdate) biet nap file nao
Public Function FaModuleName() As String
    FaModuleName = "modFA"
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
    Dim db As DAO.Database, i As Long, c As String, nm As String
    Set db = CurrentDb
    For i = 0 To db.TableDefs.Count - 1
        c = db.TableDefs(i).Connect
        nm = db.TableDefs(i).Name
        If InStr(1, c, "DATABASE=", vbTextCompare) > 0 Then
            BackEndPath = Mid$(c, InStr(1, c, "DATABASE=", vbTextCompare) + 9)
            If nm = "Tbl_Agent" Then Exit Function
        End If
    Next i
    If BackEndPath = "" Then Err.Raise vbObjectError + 20, , U("Ch\u01B0a k\u1EBFt n\u1ED1i file Data chung. B\u1EA5m 'K\u1EBFt n\u1ED1i file d\u1EEF li\u1EC7u'.")
End Function

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
    EnsureTexts
    EnsureDataSchema
    If DCount("*", "Tbl_Agent", "[Period]<=" & Q(period)) = 0 Then
        MsgU U("Ch\u01B0a c\u00F3 danh s\u00E1ch agent (Tbl_Agent) cho th\u00E1ng n\u00E0y ho\u1EB7c tr\u01B0\u1EDBc \u0111\u00F3. N\u1EA1p file Tbl_Agent tr\u01B0\u1EDBc."), MB_ICONERROR: Exit Function
    End If
    If DCount("*", "tblFA_Result", "Period=" & Q(period) & " AND Ky=" & ky) > 0 Then
        If MsgU(T("Q_Recalc"), MB_YESNO Or MB_ICONQUESTION) <> IDYES Then Exit Function
    End If

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

    ' Agent: ban chup cua thang tinh, neu chua co thi ban gan nhat truoc do (Data chung); chi lay chuc danh LP
    ' Cham cong: tong so cong trong thang (Data chung). Doanh so/K2/MDRT: tblFA_Input (tam thoi, se thay bang bang Data chung)
    Dim sqlA As String, sqlT As String, sqlI As String, titleList As String
    titleList = "'" & Replace(Mid$(FA_TITLES, 2, Len(FA_TITLES) - 2), ",", "','") & "'"
    sqlA = "SELECT x.* FROM Tbl_Agent AS x WHERE x.[Period] = (SELECT Max(y.[Period]) FROM Tbl_Agent AS y " & _
           "WHERE y.[AGENT_CODE] = x.[AGENT_CODE] AND y.[Period] <= " & Q(period) & ")"
    sqlT = "SELECT [" & AttCol("CODE") & "] AS TCode, Count(*) AS TRows, " & _
           "Sum(IIf([" & AttCol("CONG") & "] Is Null, Nz([" & AttCol("AM") & "],0) + Nz([" & AttCol("PM") & "],0), [" & AttCol("CONG") & "])) AS TDays " & _
           "FROM Tbl_Attendance WHERE [Period]=" & Q(period) & " GROUP BY [" & AttCol("CODE") & "]"
    sqlI = "SELECT * FROM tblFA_Input WHERE Period=" & Q(period) & " AND Ky=" & ky
    Set rsIn = db.OpenRecordset( _
        "SELECT a.[AGENT_CODE] AS AgentCode, a.[FULL_NAME] AS AgentName, a.[Period] AS AgentPeriod, " & _
        "a.[APPOINTMENT_DATE] AS AppointDate, a.[AGENCY_TYPE] AS WorkModel, a.[TERMINATION_DATE] AS TerminateDate, " & _
        "a.[TITLE (Code)] AS Title, a.[AGENT_STATUS] AS AgentStatus, a.[TRAINING_COMPLETED] AS AgentTraining, " & _
        "t.TDays, t.TRows, i.AgentCode AS InCode, i.FYC, i.Contracts, i.Recruits, i.K2Pct, i.MDRTPrev, i.TrainingDone, i.StatusEOM " & _
        "FROM ((" & sqlA & ") AS a LEFT JOIN (" & sqlT & ") AS t ON a.[AGENT_CODE] = t.TCode) " & _
        "LEFT JOIN (" & sqlI & ") AS i ON a.[AGENT_CODE] = i.AgentCode " & _
        "WHERE UCase(Trim(a.[TITLE (Code)] & '')) IN (" & titleList & ") OR i.AgentCode Is Not Null " & _
        "ORDER BY a.[AGENT_CODE]", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset("tblFA_Result", dbOpenDynaset, dbAppendOnly)

    Do Until rsIn.EOF
        Dim code As String, reason As String, eligible As Boolean
        Dim wm As Long, baseDays As Long, actual As Double, att As Double, attPct As Double, dayF As Double
        Dim fyc As Double, ct As Double, rec As Double, allow As Double
        Dim k2c As Double, mdrtc As Double, ctc As Double, faFormula As Double
        Dim newGuar As Double, pending As Double, retro As Double, guarAmt As Double, cumCt As Double
        Dim faMonth As Double, qtrOK As Boolean, qtrMonth As Double, topup As Double
        Dim faTotal As Double, paid As Double, m1 As Variant
        Dim hasInput As Boolean, trained As Boolean, statusEOM As String

        code = CStr(rsIn!AgentCode)
        reason = ""
        actual = Nz(rsIn!TDays, 0)
        hasInput = Not IsNull(rsIn!InCode)
        If hasInput Then
            trained = Nz(rsIn!TrainingDone, False)
            statusEOM = Nz(rsIn!StatusEOM, Nz(rsIn!AgentStatus, ""))
        Else
            trained = IsYes(rsIn!AgentTraining)
            statusEOM = Nz(rsIn!AgentStatus, "")
        End If
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
            If UCase$(Trim$(statusEOM)) = "SE" Then
                reason = T("R_SE")
            ElseIf Not trained Then
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
        rsOut!AgentName = rsIn!AgentName
        rsOut!Title = rsIn!Title
        rsOut!WorkModel = rsIn!WorkModel
        rsOut!AppointDate = rsIn!AppointDate
        rsOut!AgentPeriod = rsIn!AgentPeriod
        rsOut!AttRows = Nz(rsIn!TRows, 0)
        If hasInput Then rsOut!InputSource = "FA_Input" Else rsOut!InputSource = "Data chung"
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

' Ten cot bang cham cong Tbl_Attendance (giu nguyen tieu de file Excel)
Private Function AttCol(ByVal key As String) As String
    Select Case key
        Case "DATE": AttCol = U("Ng\u00E0y ch\u1EA5m c\u00F4ng")
        Case "CODE": AttCol = U("M\u00E3 s\u1ED1 \u0110L")
        Case "NAME": AttCol = U("H\u1ECD v\u00E0 t\u00EAn")
        Case "AM": AttCol = U("Bu\u1ED5i s\u00E1ng")
        Case "PM": AttCol = U("Bu\u1ED5i chi\u1EC1u")
        Case "CONG": AttCol = U("S\u1ED1 c\u00F4ng")
    End Select
End Function

' File Data chung chi giu Tbl_Agent + Tbl_Attendance (dung chung cho moi khoan thuong).
' Bang rieng cua FA (tham so, nhap tay, ket qua, chot ky, ngay le) nam trong file tool nay.
' Ham chay lai nhieu lan khong sao: tu chuyen cau truc cu sang moi, khong mat du lieu.
Private Sub EnsureDataSchema()
    Dim bePath As String, be As DAO.Database, t As Variant
    bePath = BackEndPath()
    For Each t In Array("tblFA_Input", "tblFA_Result", "tblPeriodStatus", "tblParam", "tblParam_Allowance", _
                        "tblParam_K2", "tblParam_Contract", "tblHoliday")
        MoveToLocal CStr(t), bePath
    Next t
    EnsureResultColumns

    Set be = DBEngine.OpenDatabase(bePath)
    EnsureAgentTable be
    EnsureAttendanceTable be
    MigrateOldAgent be, "tblAgentHist"
    MigrateOldAgent be, "tblAgent"
    be.Close
    DropLink "tblAgent"
    DropLink "tblAgentHist"
    LinkTable "Tbl_Agent", bePath
    LinkTable "Tbl_Attendance", bePath
    CurrentDb.TableDefs.Refresh
End Sub

' Chuoi ket noi cua bang trong file tool: "" = bang cuc bo, "-" = khong co bang, con lai = bang lien ket
Private Function LocalConnect(ByVal name As String) As String
    Dim db As DAO.Database, i As Long
    Set db = CurrentDb
    db.TableDefs.Refresh
    For i = 0 To db.TableDefs.Count - 1
        If StrComp(db.TableDefs(i).Name, name, vbTextCompare) = 0 Then
            LocalConnect = db.TableDefs(i).Connect
            Exit Function
        End If
    Next i
    LocalConnect = "-"
End Function

Private Function LocalFieldExists(ByVal tbl As String, ByVal fld As String) As Boolean
    Dim db As DAO.Database, f As DAO.Field
    Set db = CurrentDb
    If LocalConnect(tbl) = "-" Then Exit Function
    For Each f In db.TableDefs(tbl).Fields
        If StrComp(f.Name, fld, vbTextCompare) = 0 Then LocalFieldExists = True
    Next f
End Function

' Bang dang lien ket sang file Data -> chep ve file tool roi xoa ben Data
Private Sub MoveToLocal(ByVal name As String, ByVal bePath As String)
    Dim be As DAO.Database
    If Len(LocalConnect(name)) = 0 Or LocalConnect(name) = "-" Then Exit Sub
    DoCmd.DeleteObject acTable, name
    DoCmd.TransferDatabase acImport, "Microsoft Access", bePath, acTable, name, name
    Set be = DBEngine.OpenDatabase(bePath)
    be.Execute "DROP TABLE [" & name & "]", dbFailOnError
    be.Close
End Sub

Private Sub DropLink(ByVal name As String)
    Dim c As String
    c = LocalConnect(name)
    If Len(c) > 0 And c <> "-" Then DoCmd.DeleteObject acTable, name
End Sub

Private Sub EnsureAgentTable(ByVal be As DAO.Database)
    Dim spec As Variant, p As Variant, i As Long, ddl As String
    spec = AgentSpec()
    If TableExists(be, "Tbl_Agent") Then
        For i = 0 To UBound(spec)
            p = Split(spec(i), "|")
            If Not HasField(be, "Tbl_Agent", p(4)) Then
                be.Execute "ALTER TABLE [Tbl_Agent] ADD COLUMN " & FieldDDL(p), dbFailOnError
                be.TableDefs.Refresh
            End If
        Next i
        Exit Sub
    End If
    For i = 0 To UBound(spec)
        p = Split(spec(i), "|")
        ddl = ddl & IIf(i = 0, "", ", ") & FieldDDL(p)
    Next i
    be.Execute "CREATE TABLE [Tbl_Agent] (" & ddl & ", CONSTRAINT pkTblAgent PRIMARY KEY ([Period], [AGENT_CODE]))", dbFailOnError
    be.TableDefs.Refresh
End Sub

Private Sub EnsureAttendanceTable(ByVal be As DAO.Database)
    If TableExists(be, "Tbl_Attendance") Then Exit Sub
    be.Execute "CREATE TABLE [Tbl_Attendance] ([Period] TEXT(6) NOT NULL, [" & AttCol("DATE") & "] DATETIME NOT NULL, " & _
               "[" & AttCol("CODE") & "] TEXT(20) NOT NULL, [" & AttCol("NAME") & "] TEXT(255), " & _
               "[" & AttCol("AM") & "] DOUBLE, [" & AttCol("PM") & "] DOUBLE, [" & AttCol("CONG") & "] DOUBLE, " & _
               "CONSTRAINT pkTblAttendance PRIMARY KEY ([Period], [" & AttCol("DATE") & "], [" & AttCol("CODE") & "]))", dbFailOnError
End Sub

' Chep du lieu agent tu bang cu (tblAgent / tblAgentHist, ten cot cu hoac moi) sang Tbl_Agent roi xoa bang cu
Private Sub MigrateOldAgent(ByVal be As DAO.Database, ByVal oldName As String)
    Dim spec As Variant, p As Variant, i As Long, cols As String, src As String, per As String
    If Not TableExists(be, oldName) Then Exit Sub
    spec = AgentSpec()
    For i = 0 To UBound(spec)
        p = Split(spec(i), "|")
        If HasField(be, oldName, p(4)) Then
            cols = cols & ", [" & p(4) & "]": src = src & ", [" & p(4) & "]"
            If p(4) = "Period" Then per = p(4)
        ElseIf HasField(be, oldName, p(2)) Then
            cols = cols & ", [" & p(4) & "]": src = src & ", [" & p(2) & "]"
            If p(4) = "Period" Then per = p(2)
        End If
    Next i
    ' Dong khong co Period khong xac dinh duoc thang -> bo qua; dong trung khoa bi bo qua (khong dung dbFailOnError)
    If Len(per) > 0 Then
        be.Execute "INSERT INTO [Tbl_Agent] (" & Mid$(cols, 3) & ") SELECT " & Mid$(src, 3) & _
                   " FROM [" & oldName & "] WHERE [" & per & "] Is Not Null"
    End If
    be.Execute "DROP TABLE [" & oldName & "]", dbFailOnError
    be.TableDefs.Refresh
End Sub

' Luu kem thong tin dau vao da dung vao bang ket qua de tra lai ky cu khong phu thuoc Data chung
Private Sub EnsureResultColumns()
    Dim cols As Variant, c As Variant
    If LocalConnect("tblFA_Result") = "-" Then Exit Sub
    For Each c In Array("AgentName|TEXT(255)", "Title|TEXT(20)", "WorkModel|TEXT(10)", "AppointDate|DATETIME", _
                        "AgentPeriod|TEXT(6)", "AttRows|INTEGER", "InputSource|TEXT(20)")
        cols = Split(c, "|")
        If Not LocalFieldExists("tblFA_Result", cols(0)) Then
            CurrentDb.Execute "ALTER TABLE tblFA_Result ADD COLUMN [" & cols(0) & "] " & cols(1), dbFailOnError
            CurrentDb.TableDefs.Refresh
        End If
    Next c
End Sub

Private Function TableExists(ByVal db As DAO.Database, ByVal name As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If td.Name = name Then TableExists = True
    Next td
End Function

Private Function HasField(ByVal db As DAO.Database, ByVal tbl As String, ByVal name As String) As Boolean
    Dim f As DAO.Field
    For Each f In db.TableDefs(tbl).Fields
        If StrComp(f.Name, name, vbTextCompare) = 0 Then HasField = True
    Next f
End Function

Private Function FieldDDL(ByVal p As Variant) As String
    If p(4) = "AGENT_CODE" Then
        FieldDDL = "[AGENT_CODE] TEXT(20) NOT NULL"
    ElseIf p(3) = "D" Or p(3) = "CD" Then
        FieldDDL = "[" & p(4) & "] DATETIME"
    Else
        FieldDDL = "[" & p(4) & "] TEXT(255)"
    End If
End Function

Private Sub LinkTable(ByVal name As String, ByVal bePath As String)
    Dim db As DAO.Database, c As String
    c = LocalConnect(name)
    If c = "-" Then
        DoCmd.TransferDatabase acLink, "Microsoft Access", bePath, acTable, name, name
    ElseIf Len(c) > 0 Then
        Set db = CurrentDb
        db.TableDefs(name).Connect = ";DATABASE=" & bePath
        db.TableDefs(name).RefreshLink
    End If
End Sub

' Query hien thi Tbl_Agent theo dung thu tu va tieu de cot cua file he thong
Private Sub BuildAgentQuery()
    Dim spec As Variant, p As Variant, i As Long, sql As String
    spec = AgentSpec()
    For i = 0 To UBound(spec)
        p = Split(spec(i), "|")
        sql = sql & IIf(i = 0, "", ", ") & "[" & p(4) & "]"
    Next i
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "qryAgent"
    On Error GoTo 0
    CurrentDb.CreateQueryDef "qryAgent", "SELECT " & sql & " FROM Tbl_Agent ORDER BY [Period] DESC, [AGENT_CODE]"
    BuildBangKeQuery
End Sub

' Bang ke FA (ten cot tieng Viet), lay tu bang ket qua da luu - khong phu thuoc Data chung
Private Sub BuildBangKeQuery()
    Dim s As String
    s = "SELECT "
    s = s & U("r.Period")
    s = s & U(", r.Ky")
    s = s & U(", r.AgentCode")
    s = s & U(", r.AgentName AS [H\u1ECD t\u00EAn]")
    s = s & U(", r.Title AS [Ch\u1EE9c danh]")
    s = s & U(", r.WorkModel AS [M\u00F4 h\u00ECnh]")
    s = s & U(", r.WorkMonthNo AS [Th\u00E1ng l\u00E0m vi\u1EC7c th\u1EE9]")
    s = s & U(", r.WorkDaysBase AS [S\u1ED1 ng\u00E0y l\u00E0m vi\u1EC7c c\u1EE7a th\u00E1ng]")
    s = s & U(", r.ActualDays AS [S\u1ED1 ng\u00E0y \u0111i l\u00E0m th\u1EF1c t\u1EBF]")
    s = s & U(", r.AttendancePct AS [T\u1EF7 l\u1EC7 ch\u1EA5m c\u00F4ng]")
    s = s & U(", r.DayFactor AS [H\u1EC7 s\u1ED1 ng\u00E0y l\u00E0m vi\u1EC7c]")
    s = s & U(", r.FYC AS [FYC ph\u00E1t sinh]")
    s = s & U(", r.Contracts AS [S\u1ED1 H\u0110 ph\u00E1t sinh]")
    s = s & U(", r.Recruits AS [S\u1ED1 LP gi\u1EDBi thi\u1EC7u]")
    s = s & U(", r.AllowancePerDay AS [Ph\u1EE5 c\u1EA5p m\u1ED7i ng\u00E0y]")
    s = s & U(", r.K2Pct AS [K2]")
    s = s & U(", r.K2Coef AS [H\u1EC7 s\u1ED1 K2]")
    s = s & U(", r.MDRTCoef AS [H\u1EC7 s\u1ED1 MDRT]")
    s = s & U(", r.ContractCoef AS [H\u1EC7 s\u1ED1 h\u1EE3p \u0111\u1ED3ng]")
    s = s & U(", r.FA_Formula AS [FA theo c\u00F4ng th\u1EE9c]")
    s = s & U(", r.NewGuar AS [B\u1EA3o \u0111\u1EA3m LP m\u1EDBi]")
    s = s & U(", r.NewGuarPending AS [B\u1EA3o \u0111\u1EA3m LP m\u1EDBi \u0111ang ch\u1EDD]")
    s = s & U(", r.NewGuarRetro AS [B\u00F9 b\u1EA3o \u0111\u1EA3m th\u00E1ng 1]")
    s = s & U(", r.FA_Month AS [FA th\u00E1ng]")
    s = s & U(", r.QtrGuarFlag AS [\u0110\u01B0\u1EE3c b\u1EA3o \u0111\u1EA3m qu\u00FD]")
    s = s & U(", r.QtrGuarMonth AS [M\u1EE9c b\u1EA3o \u0111\u1EA3m qu\u00FD c\u1EE7a th\u00E1ng]")
    s = s & U(", r.QtrGuarTopup AS [B\u00F9 b\u1EA3o \u0111\u1EA3m qu\u00FD]")
    s = s & U(", r.FA_Total AS [T\u1ED5ng FA l\u0169y k\u1EBF]")
    s = s & U(", r.PaidPrevKy AS [\u0110\u00E3 chi c\u00E1c k\u1EF3 tr\u01B0\u1EDBc]")
    s = s & U(", r.Payable AS [Chi k\u1EF3 n\u00E0y]")
    s = s & U(", r.Eligible AS [\u0110\u1EE7 \u0111i\u1EC1u ki\u1EC7n]")
    s = s & U(", r.Reason AS [Ghi ch\u00FA]")
    s = s & U(", r.CalcAt AS [Th\u1EDDi \u0111i\u1EC3m t\u00EDnh]")
    s = s & U(", r.CalcBy AS [Ng\u01B0\u1EDDi t\u00EDnh]")
    s = s & U(", r.AppointDate AS [Ng\u00E0y b\u1ED5 nhi\u1EC7m], r.AgentPeriod AS [Th\u00E1ng d\u1EEF li\u1EC7u agent], r.AttRows AS [S\u1ED1 d\u00F2ng ch\u1EA5m c\u00F4ng], r.InputSource AS [Ngu\u1ED3n s\u1ED1 li\u1EC7u]")
    s = s & " FROM tblFA_Result AS r"
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "qryFA_BangKe"
    On Error GoTo 0
    CurrentDb.CreateQueryDef "qryFA_BangKe", s
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
        If p(3) <> "D" And p(3) <> "CD" And InStr(UCase$(p(4)), "NAME") > 0 Then c.ColumnWidth = 2800 Else c.ColumnWidth = 1700
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
    f.OnLoad = "=AgentInit()"

    Set c = AddCtl(fn, acLabel, "lblSearch", 200, 200, 2400, 320)
    c.Caption = U("T\u00ECm m\u00E3 / h\u1ECD t\u00EAn")
    Set c = AddCtl(fn, acTextBox, "txtSearch", 2600, 200, 2600, 320)
    c.AfterUpdate = "=AgentFilter()"
    Set c = AddCtl(fn, acLabel, "lblTitle", 5500, 200, 1200, 320)
    c.Caption = U("Ch\u1EE9c danh")
    Set c = AddCtl(fn, acComboBox, "cboTitle", 6700, 200, 1300, 320)
    c.RowSourceType = "Table/Query"
    c.RowSource = "SELECT DISTINCT [TITLE (Code)] FROM Tbl_Agent WHERE [TITLE (Code)] Is Not Null ORDER BY [TITLE (Code)]"
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
    Set c = AddCtl(fn, acLabel, "lblPeriod", 200, 680, 2400, 320)
    c.Caption = U("Th\u00E1ng (Period)")
    Set c = AddCtl(fn, acComboBox, "cboPeriod", 2600, 680, 1500, 320)
    c.RowSourceType = "Table/Query"
    c.RowSource = "SELECT DISTINCT [Period] FROM Tbl_Agent ORDER BY [Period] DESC"
    c.AfterUpdate = "=AgentFilter()"
    Set c = AddCtl(fn, acLabel, "lblCount", 4400, 680, 5000, 320)
    c.Caption = " "
    c.ForeColor = RGB(31, 78, 121)
    Set c = AddCtl(fn, acCommandButton, "btnAttendance", 9700, 640, 2600, 380)
    c.Caption = U("Xem ch\u1EA5m c\u00F4ng")
    c.OnClick = "=BtnAttendance()"

    Set c = AddCtl(fn, acSubform, "sfList", 200, 1150, 15600, 7650)
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
    If Len(Nz(f!cboPeriod, "")) > 0 Then w = "[Period]=" & Q(f!cboPeriod)
    If Len(s) > 0 Then w = w & IIf(Len(w) > 0, " AND ", "") & "([AGENT_CODE] Like '*" & s & "*' Or [FULL_NAME] Like '*" & s & "*')"
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

Public Function AgentInit()
    Forms("frmAgent")!cboPeriod = DMax("[Period]", "Tbl_Agent")
    AgentFilter
End Function

' Tong hop cham cong theo agent (Data chung), loc theo thang dang chon tren form Danh sach Agent
Public Function BtnAttendance()
    Dim per As String, sql As String
    If CurrentProject.AllForms("frmAgent").IsLoaded Then per = Nz(Forms("frmAgent")!cboPeriod, "")
    sql = "SELECT [Period], [" & AttCol("CODE") & "], First([" & AttCol("NAME") & "]) AS [" & AttCol("NAME") & "], " & _
          "Count(*) AS [S\u1ED1 ng\u00E0y ch\u1EA5m], Sum([" & AttCol("CONG") & "]) AS [T\u1ED5ng c\u00F4ng] " & _
          "FROM Tbl_Attendance " & IIf(Len(per) > 0, "WHERE [Period]=" & Q(per) & " ", "") & _
          "GROUP BY [Period], [" & AttCol("CODE") & "] ORDER BY [Period] DESC, [" & AttCol("CODE") & "]"
    sql = U(sql)
    DoCmd.Close acQuery, "qryChamCong"
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "qryChamCong"
    On Error GoTo 0
    CurrentDb.CreateQueryDef "qryChamCong", sql
    DoCmd.OpenQuery "qryChamCong", acViewNormal, acReadOnly
End Function

Public Function AgentClearFilter()
    Forms("frmAgent")!txtSearch = Null
    Forms("frmAgent")!cboTitle = Null
    Forms("frmAgent")!cboPeriod = DMax("[Period]", "Tbl_Agent")
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
        EnsureDataSchema
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
    EnsureDataSchema
    BuildAgentForms
    EnsureMainButton "btnAgents", U("Danh s\u00E1ch Agent"), "=BtnAgents()", 5000, 4440
End Function

' O Excel loi (#N/A, #VALUE!...) tra ve chuoi rong thay vi gay loi Type mismatch
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = CStr(v)
End Function

Private Function IsYes(ByVal v As Variant) As Boolean
    Select Case UCase$(Trim$(CellText(v)))
        Case "Y", "YES", "1", "TRUE", "X", "-1": IsYes = True
    End Select
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

' Doc sheet Agent theo ten cot: file xuat he thong 98 cot (chuan) hoac mau 5 cot cu (AgentCode, AgentName...)
' lan file xuat tu he thong (AGENT_CODE, FULL_NAME, APPOINTMENT_DATE, AGENCY_TYPE, TITLE (Code)...).
' Ten cot duoc so khop sau khi bo khoang trang va dau gach duoi.
Private Function ImportAgents(ByVal path As String, ByRef nAgents As Long) As Boolean
    Dim xl As Object, wb As Object, ws As Object, data As Variant, hdr As Object, seen As Object
    Dim spec As Variant, p As Variant, col() As Long, r As Long, c As Long, key As String, code As String, v As String
    Dim cCode As Long, cAppoint As Long, cModel As Long, cPeriod As Long
    Dim rs As DAO.Recordset, periods As Object, pk As Variant

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

    ' Tieu de -> vi tri cot; cot trung ten duoc danh so theo lan xuat hien (#1, #2)
    Set hdr = NewDict()
    Set seen = NewDict()
    For c = 1 To UBound(data, 2)
        key = NormHeader(data(1, c))
        If Len(key) > 0 Then
            seen(key) = seen(key) + 1
            hdr(key & "#" & seen(key)) = c
        End If
    Next c

    spec = AgentSpec()
    ReDim col(UBound(spec))
    For c = 0 To UBound(spec)
        p = Split(spec(c), "|")
        key = NormHeader(p(0)) & "#" & p(1)
        If hdr.Exists(key) Then
            col(c) = hdr(key)
        ElseIf hdr.Exists(NormHeader(p(4)) & "#1") Then
            col(c) = hdr(NormHeader(p(4)) & "#1")
        Else
            col(c) = FindCol(hdr, OldTemplateHeader(p(4)))
        End If
        Select Case p(4)
            Case "AGENT_CODE": cCode = col(c)
            Case "APPOINTMENT_DATE": cAppoint = col(c)
            Case "AGENCY_TYPE": cModel = col(c)
            Case "Period": cPeriod = col(c)
        End Select
    Next c
    If cCode = 0 Or cAppoint = 0 Or cModel = 0 Then Err.Raise vbObjectError + 10, , T("E_AgentCols")
    If cPeriod = 0 Then Err.Raise vbObjectError + 11, , U("Sheet Agent thi\u1EBFu c\u1ED9t Period (th\u00E1ng d\u1EEF li\u1EC7u, d\u1EA1ng YYYYMM).")

    ' Nap lai thang nao thi thay toan bo du lieu thang do; cac thang khac giu nguyen
    Set periods = NewDict()
    For r = 2 To UBound(data, 1)
        key = Trim$(CellText(data(r, cPeriod)))
        If Len(key) > 0 And Len(Trim$(CellText(data(r, cCode)))) > 0 Then periods(key) = True
    Next r
    For Each pk In periods.Keys
        CurrentDb.Execute "DELETE FROM Tbl_Agent WHERE [Period]=" & Q(CStr(pk)), dbFailOnError
    Next pk

    Set seen = NewDict()
    Set rs = CurrentDb.OpenRecordset("Tbl_Agent", dbOpenDynaset, dbAppendOnly)
    For r = 2 To UBound(data, 1)
        code = Trim$(CellText(data(r, cCode)))
        key = Trim$(CellText(data(r, cPeriod)))
        If Len(code) > 0 And Len(key) > 0 And Not seen.Exists(key & "|" & code) Then
            seen(key & "|" & code) = True
            rs.AddNew
            For c = 0 To UBound(spec)
                If col(c) > 0 Then
                    p = Split(spec(c), "|")
                    If p(3) = "D" Or p(3) = "CD" Then
                        rs.Fields(p(4)) = ToDateOrNull(data(r, col(c)))
                    Else
                        v = Trim$(CellText(data(r, col(c))))
                        If p(4) = "AGENCY_TYPE" Or p(4) = "TITLE (Code)" Then v = UCase$(v)
                        If Len(v) = 0 Then
                            rs.Fields(p(4)) = Null
                        Else
                            rs.Fields(p(4)) = Left$(v, 255)
                        End If
                    End If
                End If
            Next c
            rs.Fields("AGENT_CODE") = code
            rs.Fields("Period") = key
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

' Doc bang cham cong (sheet bat ky co cot Ma so DL + Ngay cham cong) vao Tbl_Attendance.
' Nap lai thang nao thi thay du lieu thang do. Dong trung (cung ngay, cung ma) chi giu 1 dong.
Private Function ImportAttendance(ByVal path As String, ByRef nRows As Long, ByRef nDup As Long) As Boolean
    Dim xl As Object, wb As Object, ws As Object, data As Variant, hdr As Object, seen As Object
    Dim cPer As Long, cNgay As Long, cCode As Long, cName As Long, cAM As Long, cPM As Long, cCong As Long
    Dim r As Long, c As Long, key As String, code As String, per As String, d As Variant
    Dim rs As DAO.Recordset, periods As Object, pk As Variant, am As Double, pm As Double, cong As Variant

    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(path, 0, True)
    For Each ws In wb.Worksheets
        data = ws.UsedRange.Value
        If IsArray(data) Then
            Set hdr = NewDict()
            For c = 1 To UBound(data, 2)
                key = NormHeader(data(1, c))
                If Len(key) > 0 And Not hdr.Exists(key) Then hdr(key) = c
            Next c
            If hdr.Exists(NormHeader(AttCol("CODE"))) And hdr.Exists(NormHeader(AttCol("DATE"))) Then Exit For
        End If
        Set hdr = Nothing
    Next ws
    If hdr Is Nothing Then GoTo Done

    cCode = hdr(NormHeader(AttCol("CODE")))
    cNgay = hdr(NormHeader(AttCol("DATE")))
    If hdr.Exists("PERIOD") Then cPer = hdr("PERIOD")
    If hdr.Exists(NormHeader(AttCol("NAME"))) Then cName = hdr(NormHeader(AttCol("NAME")))
    If hdr.Exists(NormHeader(AttCol("AM"))) Then cAM = hdr(NormHeader(AttCol("AM")))
    If hdr.Exists(NormHeader(AttCol("PM"))) Then cPM = hdr(NormHeader(AttCol("PM")))
    If hdr.Exists(NormHeader(AttCol("CONG"))) Then cCong = hdr(NormHeader(AttCol("CONG")))

    Set periods = NewDict()
    Set seen = NewDict()
    For r = 2 To UBound(data, 1)
        code = Trim$(CellText(data(r, cCode)))
        d = ToDateOrNull(data(r, cNgay))
        If Len(code) > 0 And Not IsNull(d) Then
            per = ""
            If cPer > 0 Then per = Trim$(CellText(data(r, cPer)))
            If Len(per) = 0 Then per = Format$(d, "yyyymm")
            key = per & "|" & Format$(d, "yyyymmdd") & "|" & code
            If seen.Exists(key) Then
                nDup = nDup + 1
            Else
                seen(key) = r
                periods(per) = True
            End If
        End If
    Next r

    For Each pk In periods.Keys
        CurrentDb.Execute "DELETE FROM Tbl_Attendance WHERE [Period]=" & Q(CStr(pk)), dbFailOnError
    Next pk
    Set rs = CurrentDb.OpenRecordset("Tbl_Attendance", dbOpenDynaset, dbAppendOnly)
    For Each pk In seen.Keys
        r = seen(pk)
        am = 0: pm = 0
        If cAM > 0 Then am = Val(CellText(data(r, cAM)))
        If cPM > 0 Then pm = Val(CellText(data(r, cPM)))
        cong = Null
        If cCong > 0 Then
            If IsNumeric(CellText(data(r, cCong))) And Len(CellText(data(r, cCong))) > 0 Then cong = CDbl(data(r, cCong))
        End If
        If IsNull(cong) Then cong = am + pm
        rs.AddNew
        rs.Fields("Period") = Split(pk, "|")(0)
        rs.Fields(AttCol("DATE")) = ToDateOrNull(data(r, cNgay))
        rs.Fields(AttCol("CODE")) = Trim$(CellText(data(r, cCode)))
        If cName > 0 Then rs.Fields(AttCol("NAME")) = Left$(Trim$(CellText(data(r, cName))), 255)
        rs.Fields(AttCol("AM")) = am
        rs.Fields(AttCol("PM")) = pm
        rs.Fields(AttCol("CONG")) = cong
        rs.Update
        nRows = nRows + 1
    Next pk
    rs.Close
Done:
    ImportAttendance = True
Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Exit Function
Fail:
    mErr = Err.Description
    Resume Cleanup
End Function

' Ten cot cua mau 5 cot cu (truoc khi dung file he thong)
Private Function OldTemplateHeader(ByVal field As String) As String
    Select Case field
        Case "AGENT_CODE": OldTemplateHeader = "AGENTCODE"
        Case "FULL_NAME": OldTemplateHeader = "AGENTNAME"
        Case "APPOINTMENT_DATE": OldTemplateHeader = "APPOINTDATE"
        Case "AGENCY_TYPE": OldTemplateHeader = "WORKMODEL"
        Case "TERMINATION_DATE": OldTemplateHeader = "TERMINATEDATE"
        Case "TITLE (Code)": OldTemplateHeader = "TITLE"
        Case Else: OldTemplateHeader = "-"
    End Select
End Function

Public Function BtnImport()
    Dim path As String, db As DAO.Database, nIn As Long, locked As Long, nAg As Long, nAtt As Long, nDup As Long, msg As String
    On Error GoTo Fail

    path = PickFile(T("D_PickExcel"), "Excel", "*.xlsx;*.xlsm;*.xls")
    If path = "" Then Exit Function
    If Not AutoBackup() Then Exit Function
    Set db = CurrentDb
    SetStatus T("S_Importing")

    EnsureTexts
    EnsureDataSchema
    If Not ImportAgents(path, nAg) Then
        SetStatus T("E_Import")
        MsgU T("E_Import") & vbCrLf & mErr, MB_ICONERROR
        Exit Function
    End If
    If Not ImportAttendance(path, nAtt, nDup) Then
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
    msg = T("S_AgentsImported") & " " & nAg & vbCrLf & _
          U("\u0110\u00E3 n\u1EA1p ch\u1EA5m c\u00F4ng, s\u1ED1 d\u00F2ng:") & " " & nAtt
    If nDup > 0 Then msg = msg & vbCrLf & U("B\u1ECF qua d\u00F2ng ch\u1EA5m c\u00F4ng tr\u00F9ng (c\u00F9ng ng\u00E0y, c\u00F9ng m\u00E3):") & " " & nDup
    msg = msg & vbCrLf & T("S_Imported") & " " & nIn
    SetStatus Replace(msg, vbCrLf, " | ")
    MsgU msg
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

' Chon lai file Data (Data.accdb) khi chuyen len thu muc chung
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
