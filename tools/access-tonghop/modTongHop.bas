Option Compare Database
Option Explicit

' =====================================================================
' TONG HOP CAC KHOAN THUONG CUA TUNG AGENT
' Doc bang ket qua tu cac file Access khoan thuong (FA, thuong thang, thuong quy...)
' theo khai bao trong tblBonusSource; luu ket qua vao tblTongHop (tra lai bat ky luc nao).
' Khong ghi gi vao file khoan thuong (mo o che do chi doc).
' Lan dau: tao file Access trong -> Alt+F11 -> Import modTongHop.bas -> Ctrl+G -> THSetup
' Chu tieng Viet viet dang \uXXXX vi module chi chua ky tu ASCII.
' =====================================================================

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Const TH_VERSION As String = "2026.11.01-2"
Private Const FORM_TH As String = "frmTongHop"
Private Const MB_YESNO As Long = 4
Private Const MB_ICONERROR As Long = 16
Private Const MB_ICONQUESTION As Long = 32
Private Const MB_ICONINFO As Long = 64
Private Const IDYES As Long = 6

Public Function THVersion() As String
    THVersion = TH_VERSION
End Function

' =====================================================================
' Tien ich
' =====================================================================

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

Private Function MsgU(ByVal text As String, Optional ByVal flags As Long = MB_ICONINFO) As Long
    MsgU = MessageBoxW(Application.hWndAccessApp, StrPtr(text), StrPtr(TITLE_APP()), flags)
End Function

Private Function TITLE_APP() As String
    TITLE_APP = U("T\u1ED5ng h\u1EE3p thu nh\u1EADp c\u00E1c kho\u1EA3n th\u01B0\u1EDFng")
End Function

Private Function Q(ByVal s As String) As String
    Q = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function B(ByVal name As String) As String
    B = "[" & name & "]"
End Function

Private Sub SetStatus(ByVal text As String)
    If CurrentProject.AllForms(FORM_TH).IsLoaded Then
        Forms(FORM_TH)!lblStatus.Caption = text
        DoEvents
    End If
End Sub

Private Function GetUserId() As String
    GetUserId = Environ$("USERNAME") & "@" & Environ$("COMPUTERNAME")
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal name As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, name, vbTextCompare) = 0 Then TableExists = True
    Next td
End Function

Private Function FieldExists(ByVal db As DAO.Database, ByVal tbl As String, ByVal fld As String) As Boolean
    Dim f As DAO.Field
    If Not TableExists(db, tbl) Then Exit Function
    For Each f In db.TableDefs(tbl).Fields
        If StrComp(f.Name, fld, vbTextCompare) = 0 Then FieldExists = True
    Next f
End Function

Private Function IsValidPeriod(ByVal period As String) As Boolean
    If Len(period) <> 6 Or Not IsNumeric(period) Then Exit Function
    If CInt(Right$(period, 2)) < 1 Or CInt(Right$(period, 2)) > 12 Then Exit Function
    IsValidPeriod = True
End Function

Private Function PickFile(ByVal title As String) As String
    With Application.FileDialog(3)
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Access", "*.accdb;*.mdb"
        If .Show Then PickFile = .SelectedItems(1)
    End With
End Function

' =====================================================================
' CAI DAT: bang, form, khoi dong
' =====================================================================

Public Sub THSetup()
    CreateTables
    BuildForm
    SetStartup
    MsgU U("\u0110\u00E3 c\u00E0i \u0111\u1EB7t xong. \u0110\u00F3ng v\u00E0 m\u1EDF l\u1EA1i file \u0111\u1EC3 d\u00F9ng form T\u1ED5ng h\u1EE3p.")
End Sub

Private Sub CreateTables()
    Dim db As DAO.Database
    Set db = CurrentDb
    If Not TableExists(db, "tblBonusSource") Then
        db.Execute "CREATE TABLE tblBonusSource (BonusCode TEXT(20) CONSTRAINT pkBonusSource PRIMARY KEY, " & _
            "BonusName TEXT(100), FilePath TEXT(255), ResultTable TEXT(64), AgentField TEXT(64), NameField TEXT(64), " & _
            "PeriodField TEXT(64), KyField TEXT(64), AmountField TEXT(64), StatusTable TEXT(64), " & _
            "SortOrder INTEGER, Active YESNO, Note TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "tblTongHop") Then
        db.Execute "CREATE TABLE tblTongHop (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, AgentCode TEXT(20) NOT NULL, " & _
            "BonusCode TEXT(20) NOT NULL, AgentName TEXT(255), Amount CURRENCY, PeriodLocked YESNO, SourceFile TEXT(255), " & _
            "CollectedAt DATETIME, CollectedBy TEXT(100), CONSTRAINT pkTongHop PRIMARY KEY (Period, Ky, AgentCode, BonusCode))", dbFailOnError
    End If
    If Not TableExists(db, "tblCollectLog") Then
        db.Execute "CREATE TABLE tblCollectLog (ID COUNTER CONSTRAINT pkCollectLog PRIMARY KEY, Period TEXT(6), Ky INTEGER, " & _
            "BonusCode TEXT(20), RowCount LONG, Total CURRENCY, PeriodLocked YESNO, Note TEXT(255), " & _
            "CollectedAt DATETIME, CollectedBy TEXT(100))", dbFailOnError
    End If
End Sub

Private Function AddCtl(ByVal formName As String, ByVal ctlType As Long, ByVal ctlName As String, _
                        ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long) As Control
    Set AddCtl = CreateControl(formName, ctlType, acDetail, , , l, t, w, h)
    AddCtl.Name = ctlName
    AddCtl.FontName = "Segoe UI"
End Function

Private Sub BuildForm()
    Dim f As Form, c As Control, fn As String, i As Long, btns As Variant, ao As AccessObject
    For Each ao In CurrentProject.AllForms
        If ao.Name = FORM_TH Then
            If ao.IsLoaded Then DoCmd.Close acForm, FORM_TH
            DoCmd.DeleteObject acForm, FORM_TH
        End If
    Next ao

    Set f = CreateForm()
    fn = f.Name
    f.Caption = TITLE_APP()
    f.RecordSelectors = False
    f.NavigationButtons = False
    f.DividingLines = False
    f.ScrollBars = 0
    f.AutoCenter = True
    f.Width = 8800
    f.Section(acDetail).Height = 7800

    Set c = AddCtl(fn, acLabel, "lblTitle", 400, 300, 8000, 500)
    c.Caption = U("T\u1ED4NG H\u1EE2P THU NH\u1EACP C\u00C1C KHO\u1EA2N TH\u01AF\u1EDENG")
    c.FontSize = 12: c.FontBold = True
    Set c = AddCtl(fn, acLabel, "lblPeriod", 400, 1000, 2600, 340)
    c.Caption = U("Th\u00E1ng (YYYYMM)")
    Set c = AddCtl(fn, acTextBox, "txtPeriod", 3100, 1000, 1600, 340)
    c.DefaultValue = "=Format(DateAdd(""m"",-1,Date()),""yyyymm"")"
    Set c = AddCtl(fn, acLabel, "lblKy", 400, 1450, 2600, 340)
    c.Caption = U("K\u1EF3 chi tr\u1EA3 (0 = c\u1EA3 th\u00E1ng)")
    Set c = AddCtl(fn, acComboBox, "cboKy", 3100, 1450, 1600, 340)
    c.RowSourceType = "Value List"
    c.RowSource = "0;1;2;3;4"
    c.LimitToList = True
    c.DefaultValue = "0"
    Set c = AddCtl(fn, acCheckBox, "chkLocked", 3100, 1920, 300, 300)
    c.DefaultValue = "True"
    Set c = AddCtl(fn, acLabel, "lblLocked", 3450, 1880, 5000, 340)
    c.Caption = U("Ch\u1EC9 l\u1EA5y k\u1EF3 \u0111\u00E3 CH\u1ED0T \u1EDF file kho\u1EA3n th\u01B0\u1EDFng")

    btns = Array( _
        Array("btnCollect", U("1. T\u1ED5ng h\u1EE3p"), "=BtnCollect()"), _
        Array("btnPivot", U("2. Xem b\u1EA3ng t\u1ED5ng h\u1EE3p"), "=BtnPivot()"), _
        Array("btnDetail", U("3. Xem chi ti\u1EBFt t\u1EEBng kho\u1EA3n"), "=BtnDetail()"), _
        Array("btnExport", U("4. Xu\u1EA5t Excel"), "=BtnExport()"), _
        Array("btnAddSource", U("Th\u00EAm file kho\u1EA3n th\u01B0\u1EDFng"), "=BtnAddSource()"), _
        Array("btnSources", U("Danh s\u00E1ch ngu\u1ED3n"), "=BtnSources()"), _
        Array("btnLog", U("Nh\u1EADt k\u00FD t\u1ED5ng h\u1EE3p"), "=BtnLog()"), _
        Array("btnBackup", U("Sao l\u01B0u"), "=BtnBackup()"))
    For i = 0 To UBound(btns)
        If i < 4 Then
            Set c = AddCtl(fn, acCommandButton, btns(i)(0), 400, 2500 + i * 560, 4300, 460)
        Else
            Set c = AddCtl(fn, acCommandButton, btns(i)(0), 5000, 2500 + (i - 4) * 560, 3400, 460)
        End If
        c.Caption = btns(i)(1)
        c.OnClick = btns(i)(2)
    Next i

    Set c = AddCtl(fn, acLabel, "lblStatus", 400, 4900, 8000, 2400)
    c.Caption = U("S\u1EB5n s\u00E0ng. L\u1EA7n \u0111\u1EA7u: b\u1EA5m 'Th\u00EAm file kho\u1EA3n th\u01B0\u1EDFng' \u0111\u1EC3 ch\u1ECDn file FA.")
    c.ForeColor = RGB(31, 78, 121)

    DoCmd.Close acForm, fn, acSaveYes
    DoCmd.Rename FORM_TH, acForm, fn
End Sub

Private Sub SetProp(ByVal name As String, ByVal value As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(name) = value
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(name, dbText, value)
    End If
End Sub

Private Sub SetStartup()
    SetProp "StartupForm", FORM_TH
    SetProp "AppTitle", TITLE_APP()
    Application.RefreshTitleBar
End Sub

' =====================================================================
' NGUON: cac file khoan thuong
' =====================================================================

' Chon file khoan thuong. File FA duoc nhan dien tu dong; file khac: dien them cot trong "Danh sach nguon".
Public Function BtnAddSource()
    Dim path As String, src As DAO.Database, code As String, rs As DAO.Recordset, isFA As Boolean, isFAPM As Boolean
    On Error GoTo Fail
    path = PickFile(U("Ch\u1ECDn file Access c\u1EE7a kho\u1EA3n th\u01B0\u1EDFng"))
    If path = "" Then Exit Function
    Set src = DBEngine.OpenDatabase(path, False, True)
    isFA = TableExists(src, "tblFA_Result")
    isFAPM = TableExists(src, "tblFAPM_Result")
    src.Close
    If isFA Then
        code = "FA"
    ElseIf isFAPM Then
        code = "FA_PM"
    Else
        code = UCase$(Trim$(InputBox("Bonus code (VD: TTHANG, TQUY, MDRT):", "Tong hop")))
        If code = "" Then Exit Function
    End If

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblBonusSource WHERE BonusCode=" & Q(code), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!BonusCode = code
        rs!Active = True
        rs!SortOrder = DCount("*", "tblBonusSource") + 1
    Else
        rs.Edit
    End If
    rs!FilePath = path
    If isFA Then
        rs!BonusName = U("Ph\u1EE5 c\u1EA5p c\u1ED1 \u0111\u1ECBnh (FA)")
        rs!ResultTable = "tblFA_Result"
        rs!AgentField = "AgentCode"
        rs!NameField = "AgentName"
        rs!PeriodField = "Period"
        rs!KyField = "Ky"
        rs!AmountField = "Payable"
        rs!StatusTable = "tblPeriodStatus"
    ElseIf isFAPM Then
        rs!BonusName = U("Ph\u1EE5 c\u1EA5p c\u1ED1 \u0111\u1ECBnh (FA) \u2013 PM/PD")
        rs!ResultTable = "tblFAPM_Result"
        rs!AgentField = "AgentCode"
        rs!NameField = "AgentName"
        rs!PeriodField = "Period"
        rs!KyField = "Ky"
        rs!AmountField = "Payable"
        rs!StatusTable = "tblPeriodStatus"
    ElseIf IsNull(rs!BonusName) Then
        rs!BonusName = code
    End If
    rs.Update
    rs.Close
    If isFA Or isFAPM Then
        MsgU U("\u0110\u00E3 th\u00EAm ngu\u1ED3n FA:") & vbCrLf & path
    Else
        MsgU U("\u0110\u00E3 th\u00EAm ngu\u1ED3n. \u0110i\u1EC1n t\u00EAn b\u1EA3ng k\u1EBFt qu\u1EA3 v\u00E0 t\u00EAn c\u00E1c c\u1ED9t trong b\u1EA3ng ngu\u1ED3n v\u1EEBa m\u1EDF.")
        BtnSources
    End If
    Exit Function
Fail:
    MsgU U("L\u1ED7i:") & vbCrLf & Err.Description, MB_ICONERROR
End Function

Public Function BtnSources()
    DoCmd.OpenTable "tblBonusSource"
End Function

Public Function BtnLog()
    DoCmd.OpenTable "tblCollectLog", acViewNormal, acReadOnly
End Function

' =====================================================================
' TONG HOP
' =====================================================================

Private Function ReadForm(ByRef period As String, ByRef ky As Long, ByRef lockedOnly As Boolean) As Boolean
    period = Trim$(Nz(Forms(FORM_TH)!txtPeriod, ""))
    ky = CLng(Nz(Forms(FORM_TH)!cboKy, 0))
    lockedOnly = Nz(Forms(FORM_TH)!chkLocked, True)
    If Not IsValidPeriod(period) Then
        MsgU U("Th\u00E1ng kh\u00F4ng h\u1EE3p l\u1EC7. Nh\u1EADp d\u1EA1ng YYYYMM, v\u00ED d\u1EE5 202610."), MB_ICONERROR
        Exit Function
    End If
    ReadForm = True
End Function

' Ky = 0: tong hop ca 4 ky chi tra cua thang
Public Function BtnCollect()
    Dim period As String, ky As Long, lockedOnly As Boolean, k As Long, kFrom As Long, kTo As Long
    Dim rsSrc As DAO.Recordset, report As String, ws As DAO.Workspace, inTrans As Boolean
    On Error GoTo Fail
    If Not ReadForm(period, ky, lockedOnly) Then Exit Function
    If MsgU(U("T\u1ED5ng h\u1EE3p l\u1EA1i (ghi \u0111\u00E8 k\u1EBFt qu\u1EA3 t\u1ED5ng h\u1EE3p c\u0169 c\u1EE7a) th\u00E1ng/k\u1EF3") & " " & period & IIf(ky = 0, "", " / " & ky) & "?", MB_YESNO Or MB_ICONQUESTION) <> IDYES Then Exit Function
    If ky = 0 Then kFrom = 1: kTo = 4 Else kFrom = ky: kTo = ky

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    inTrans = True
    For k = kFrom To kTo
        CurrentDb.Execute "DELETE FROM tblTongHop WHERE Period=" & Q(period) & " AND Ky=" & k, dbFailOnError
        Set rsSrc = CurrentDb.OpenRecordset("SELECT * FROM tblBonusSource WHERE Active=True ORDER BY SortOrder, BonusCode", dbOpenSnapshot)
        Do Until rsSrc.EOF
            report = report & CollectOne(rsSrc, period, k, lockedOnly) & vbCrLf
            rsSrc.MoveNext
        Loop
        rsSrc.Close
    Next k
    ws.CommitTrans
    inTrans = False
    SetStatus report
    MsgU U("\u0110\u00E3 t\u1ED5ng h\u1EE3p xong:") & vbCrLf & vbCrLf & report
    Exit Function
Fail:
    If inTrans Then ws.Rollback
    MsgU U("L\u1ED7i:") & vbCrLf & Err.Description, MB_ICONERROR
End Function

' Lay ket qua 1 khoan thuong cho 1 thang/ky; tra ve 1 dong mo ta de bao cao
Private Function CollectOne(ByVal s As DAO.Recordset, ByVal period As String, ByVal ky As Long, ByVal lockedOnly As Boolean) As String
    Dim code As String, path As String, src As DAO.Database, rs As DAO.Recordset, out As DAO.Recordset
    Dim sql As String, locked As Boolean, n As Long, total As Double, note As String, kyF As String

    code = s!BonusCode
    path = Nz(s!FilePath, "")
    kyF = Nz(s!KyField, "")
    CollectOne = code & " / " & ky & ": "

    If Len(path) = 0 Or Dir$(path) = "" Then
        note = U("kh\u00F4ng t\u00ECm th\u1EA5y file")
        GoTo LogIt
    End If
    Set src = DBEngine.OpenDatabase(path, False, True)
    If Not TableExists(src, Nz(s!ResultTable, "")) Then
        note = U("kh\u00F4ng c\u00F3 b\u1EA3ng") & " " & Nz(s!ResultTable, "")
        GoTo LogIt
    End If

    ' Trang thai chot ky ben file khoan thuong (khong khai bao bang trang thai -> coi nhu da chot)
    locked = True
    If Len(Nz(s!StatusTable, "")) > 0 Then
        locked = False
        If TableExists(src, s!StatusTable) Then
            sql = "SELECT Status FROM " & B(s!StatusTable) & " WHERE " & B(s!PeriodField) & "=" & Q(period)
            If Len(kyF) > 0 Then sql = sql & " AND " & B(kyF) & "=" & ky
            Set rs = src.OpenRecordset(sql, dbOpenSnapshot)
            If Not rs.EOF Then locked = (Nz(rs!Status, "") = "LOCKED")
            rs.Close
        End If
    End If
    If lockedOnly And Not locked Then
        note = U("k\u1EF3 ch\u01B0a ch\u1ED1t - b\u1ECF qua")
        GoTo LogIt
    End If

    sql = "SELECT " & B(s!AgentField) & " AS A, " & _
          IIf(Len(Nz(s!NameField, "")) > 0, "First(" & B(Nz(s!NameField, "")) & ")", "Null") & " AS N, " & _
          "Sum(" & B(s!AmountField) & ") AS S FROM " & B(s!ResultTable) & _
          " WHERE " & B(s!PeriodField) & "=" & Q(period)
    If Len(kyF) > 0 Then sql = sql & " AND " & B(kyF) & "=" & ky
    sql = sql & " GROUP BY " & B(s!AgentField)
    Set rs = src.OpenRecordset(sql, dbOpenSnapshot)
    Set out = CurrentDb.OpenRecordset("tblTongHop", dbOpenDynaset, dbAppendOnly)
    Do Until rs.EOF
        If Not IsNull(rs!A) Then
            out.AddNew
            out!Period = period
            out!Ky = ky
            out!AgentCode = CStr(rs!A)
            out!BonusCode = code
            out!AgentName = rs!N
            out!Amount = Nz(rs!S, 0)
            out!PeriodLocked = locked
            out!SourceFile = path
            out!CollectedAt = Now()
            out!CollectedBy = GetUserId()
            out.Update
            n = n + 1
            total = total + Nz(rs!S, 0)
        End If
        rs.MoveNext
    Loop
    rs.Close
    out.Close
    If Not locked Then note = U("k\u1EF3 CH\u01AFA ch\u1ED1t - s\u1ED1 t\u1EA1m")

LogIt:
    If Not src Is Nothing Then src.Close
    CurrentDb.Execute "INSERT INTO tblCollectLog (Period, Ky, BonusCode, RowCount, Total, PeriodLocked, Note, CollectedAt, CollectedBy) VALUES (" & _
        Q(period) & ", " & ky & ", " & Q(code) & ", " & n & ", " & Replace(CStr(total), ",", ".") & ", " & _
        IIf(locked, "True", "False") & ", " & Q(note) & ", Now(), " & Q(GetUserId()) & ")", dbFailOnError
    CollectOne = CollectOne & n & " agent, " & Format$(total, "#,##0") & IIf(Len(note) > 0, " (" & note & ")", "")
End Function

' =====================================================================
' XEM / XUAT
' =====================================================================

' Bang ngang: moi agent 1 dong, moi khoan thuong 1 cot, cot tong. Ky = 0: cong ca 4 ky cua thang.
Private Function PivotSQL(ByVal period As String, ByVal ky As Long) As String
    Dim w As String
    w = "WHERE t.Period=" & Q(period) & IIf(ky = 0, "", " AND t.Ky=" & ky)
    PivotSQL = "TRANSFORM Sum(t.Amount) SELECT t.Period, " & IIf(ky = 0, "", "t.Ky, ") & "t.AgentCode, " & _
        "First(t.AgentName) AS " & B(U("H\u1ECD t\u00EAn")) & ", Sum(t.Amount) AS " & B(U("T\u1ED5ng thu nh\u1EADp")) & _
        " FROM tblTongHop AS t " & w & " GROUP BY t.Period, " & IIf(ky = 0, "", "t.Ky, ") & "t.AgentCode PIVOT t.BonusCode"
End Function

Private Sub ShowQuery(ByVal name As String, ByVal sql As String)
    DoCmd.Close acQuery, name
    On Error Resume Next
    CurrentDb.QueryDefs.Delete name
    On Error GoTo 0
    CurrentDb.CreateQueryDef name, sql
    DoCmd.OpenQuery name, acViewNormal, acReadOnly
End Sub

Private Function HasData(ByVal period As String, ByVal ky As Long) As Boolean
    HasData = DCount("*", "tblTongHop", "Period=" & Q(period) & IIf(ky = 0, "", " AND Ky=" & ky)) > 0
    If Not HasData Then MsgU U("Ch\u01B0a c\u00F3 d\u1EEF li\u1EC7u t\u1ED5ng h\u1EE3p cho th\u00E1ng/k\u1EF3 n\u00E0y. B\u1EA5m 'T\u1ED5ng h\u1EE3p' tr\u01B0\u1EDBc."), MB_ICONERROR
End Function

Public Function BtnPivot()
    Dim period As String, ky As Long, lk As Boolean
    If Not ReadForm(period, ky, lk) Then Exit Function
    If Not HasData(period, ky) Then Exit Function
    ShowQuery "qryTongHop_Xem", PivotSQL(period, ky)
End Function

Public Function BtnDetail()
    Dim period As String, ky As Long, lk As Boolean
    If Not ReadForm(period, ky, lk) Then Exit Function
    If Not HasData(period, ky) Then Exit Function
    ShowQuery "qryTongHop_ChiTiet", "SELECT t.Period, t.Ky, t.AgentCode, t.AgentName, t.BonusCode, b.BonusName, t.Amount, " & _
        "t.PeriodLocked, t.SourceFile, t.CollectedAt, t.CollectedBy FROM tblTongHop AS t LEFT JOIN tblBonusSource AS b " & _
        "ON t.BonusCode = b.BonusCode WHERE t.Period=" & Q(period) & IIf(ky = 0, "", " AND t.Ky=" & ky) & _
        " ORDER BY t.AgentCode, t.Ky, b.SortOrder"
End Function

Public Function BtnExport()
    Dim period As String, ky As Long, lk As Boolean, path As String
    On Error GoTo Fail
    If Not ReadForm(period, ky, lk) Then Exit Function
    If Not HasData(period, ky) Then Exit Function
    With Application.FileDialog(2)
        .Title = U("L\u01B0u b\u1EA3ng t\u1ED5ng h\u1EE3p")
        .InitialFileName = CurrentProject.Path & "\TongHop_" & period & IIf(ky = 0, "", "_Ky" & ky) & ".xlsx"
        If Not .Show Then Exit Function
        path = .SelectedItems(1)
    End With
    If LCase$(Right$(path, 5)) <> ".xlsx" Then path = path & ".xlsx"
    If Dir$(path) <> "" Then Kill path
    DoCmd.Close acQuery, "qryTongHop_Export"
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "qryTongHop_Export"
    On Error GoTo Fail
    CurrentDb.CreateQueryDef "qryTongHop_Export", PivotSQL(period, ky)
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryTongHop_Export", path, True
    CurrentDb.QueryDefs.Delete "qryTongHop_Export"
    MsgU U("\u0110\u00E3 xu\u1EA5t:") & vbCrLf & path
    Exit Function
Fail:
    MsgU U("L\u1ED7i:") & vbCrLf & Err.Description, MB_ICONERROR
End Function

Public Function BtnBackup()
    Dim fso As Object, folder As String, dst As String
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = CurrentProject.Path & "\Backup"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    dst = folder & "\" & fso.GetBaseName(CurrentProject.FullName) & "_" & Format$(Now(), "yyyymmdd_hhnnss") & ".accdb"
    fso.CopyFile CurrentProject.FullName, dst, False
    MsgU U("\u0110\u00E3 sao l\u01B0u:") & vbCrLf & dst
    Exit Function
Fail:
    MsgU U("L\u1ED7i:") & vbCrLf & Err.Description, MB_ICONERROR
End Function
