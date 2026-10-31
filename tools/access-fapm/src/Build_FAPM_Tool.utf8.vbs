' =====================================================================
' DỰNG TOOL TÍNH PHỤ CẤP CỐ ĐỊNH (FA) CHO PM/PD – ALL/AD/2210-006 mục A.I
' Chạy: double-click Build_FAPM_Tool.vbs (cùng thư mục với modFAPM.bas, modUpdate.bas)
' Kết quả: FA_PM_App.accdb (tool + dữ liệu riêng của FA PM). Dùng chung Data.accdb với tool FA LP:
'   - nếu thư mục đã có Data.accdb thì kết nối vào đó (không tạo mới, không sửa dữ liệu)
'   - nếu chưa có thì tạo Data.accdb mới
' Yêu cầu: máy Windows có Microsoft Access 2010 trở lên (bản đầy đủ)
' =====================================================================
Option Explicit

Const acLink = 2, acTable = 0, acModule = 5, acForm = 2, acSaveYes = 1, acDetail = 0
Const acLabel = 100, acTextBox = 109, acComboBox = 111, acCommandButton = 104
Const dbText = 10

Dim fso, here, beFile, feFile, modFile, acc, dataExisted
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
beFile = here & "\Data.accdb"
feFile = here & "\FA_PM_App.accdb"
modFile = here & "\modFAPM.bas"

If Not fso.FileExists(modFile) Or Not fso.FileExists(here & "\modUpdate.bas") Then
    MsgBox "Không tìm thấy modFAPM.bas / modUpdate.bas cùng thư mục với script.", 16, "Dựng tool FA PM"
    WScript.Quit 1
End If
If fso.FileExists(feFile) Then
    MsgBox "Đã có FA_PM_App.accdb trong thư mục này." & vbCrLf & _
           "Đổi tên/di chuyển file cũ rồi chạy lại để tránh ghi đè dữ liệu.", 48, "Dựng tool FA PM"
    WScript.Quit 1
End If
dataExisted = fso.FileExists(beFile)

On Error Resume Next
Set acc = CreateObject("Access.Application")
Check "Mở Microsoft Access"

If Not dataExisted Then BuildBackEnd
BuildFrontEnd

acc.Quit
MsgBox "Hoàn tất." & vbCrLf & vbCrLf & _
       "• FA_PM_App.accdb: mở file này để sử dụng" & vbCrLf & _
       "• Data.accdb: " & IIfS(dataExisted, "dùng chung với tool FA LP (đã có sẵn)", "mới tạo – nạp Tbl_Agent và Tbl_Attendance") & vbCrLf & vbCrLf & _
       "Lần đầu mở, bấm 'Enable Content' hoặc thêm thư mục vào Trusted Locations.", 64, "Dựng tool FA PM"

' =====================================================================
Function IIfS(cond, a, b)
    If cond Then IIfS = a Else IIfS = b
End Function

Sub Check(stepName)
    If Err.Number <> 0 Then
        MsgBox "Lỗi ở bước: " & stepName & vbCrLf & Err.Number & " – " & Err.Description, 16, "Dựng tool FA PM"
        Err.Clear
        On Error Resume Next
        acc.CloseCurrentDatabase
        acc.Quit
        WScript.Quit 1
    End If
End Sub

Sub Exec(db, sql)
    On Error Resume Next
    db.Execute sql, 128 ' dbFailOnError
    Check "SQL: " & Left(sql, 120)
End Sub

' Số → chuỗi SQL, luôn dùng dấu chấm thập phân (máy đặt vùng Việt Nam dùng dấu phẩy)
Function N(x)
    N = Replace(CStr(x), ",", ".")
End Function

Function Q(s)
    Q = "'" & Replace(s, "'", "''") & "'"
End Function

' =====================================================================
' DATA CHUNG (chỉ khi chưa có): bảng Tbl_Agent tối thiểu, FaAfterUpdate sẽ bổ sung đủ 98 cột + Tbl_Attendance
' =====================================================================
Sub BuildBackEnd()
    Dim db
    On Error Resume Next
    acc.NewCurrentDatabase beFile
    Check "Tạo Data.accdb"
    Set db = acc.CurrentDb
    Exec db, "CREATE TABLE [Tbl_Agent] ([Period] TEXT(255), [AGENT_CODE] TEXT(20) NOT NULL, " & _
             "CONSTRAINT pkTblAgent PRIMARY KEY ([Period], [AGENT_CODE]))"
    acc.CloseCurrentDatabase
    Check "Đóng Data.accdb"
End Sub

' =====================================================================
' TOOL: bảng riêng của FA PM (cục bộ), liên kết Data chung, nhãn, module, form
' =====================================================================
Sub BuildFrontEnd()
    Dim db, r, t
    On Error Resume Next
    acc.NewCurrentDatabase feFile
    Check "Tạo FA_PM_App.accdb"
    Set db = acc.CurrentDb

    Exec db, "CREATE TABLE tblHoliday (HolidayDate DATETIME CONSTRAINT pkHoliday PRIMARY KEY, Description TEXT(100))"
    Exec db, "CREATE TABLE tblFAPM_Input (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, AgentCode TEXT(20) NOT NULL, " & _
             "GroupFYC CURRENCY, ActiveLP DOUBLE, K2Pct DOUBLE, StatusEOM TEXT(20), " & _
             "CONSTRAINT pkFAPMInput PRIMARY KEY (Period, Ky, AgentCode))"
    Exec db, "CREATE TABLE tblFAPM_Result (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, AgentCode TEXT(20) NOT NULL, " & _
             "WorkMonthNo INTEGER, PromoMonthNo INTEGER, WorkDaysBase INTEGER, ActualDays DOUBLE, AttendancePct DOUBLE, DayFactor DOUBLE, " & _
             "Eligible YESNO, Reason TEXT(255), GroupFYC CURRENCY, ActiveLP DOUBLE, CumLPTurns DOUBLE, " & _
             "AllowancePerDay CURRENCY, K2Pct DOUBLE, K2Coef DOUBLE, LPCountCoef DOUBLE, " & _
             "FA_Formula CURRENCY, JoinGuar CURRENCY, FA_Month CURRENCY, " & _
             "QtrLPTurns DOUBLE, QtrGuarLevel CURRENCY, QtrGuarMonth CURRENCY, QtrGuarTopup CURRENCY, FA_Total CURRENCY, " & _
             "PaidPrevKy CURRENCY, Payable CURRENCY, CalcAt DATETIME, CalcBy TEXT(100), " & _
             "AgentName TEXT(255), Title TEXT(20), WorkModel TEXT(10), AppointDate DATETIME, PromoDate DATETIME, " & _
             "AgentPeriod TEXT(6), AttRows INTEGER, InputSource TEXT(20), " & _
             "CONSTRAINT pkFAPMResult PRIMARY KEY (Period, Ky, AgentCode))"
    Exec db, "CREATE TABLE tblPeriodStatus (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, Status TEXT(10), " & _
             "UpdatedBy TEXT(100), UpdatedAt DATETIME, CONSTRAINT pkPeriodStatus PRIMARY KEY (Period, Ky))"
    Exec db, "CREATE TABLE tblParam (ParamKey TEXT(50) CONSTRAINT pkParam PRIMARY KEY, ParamValue DOUBLE, Description TEXT(255))"
    Exec db, "CREATE TABLE tblParam_Allowance (ID COUNTER CONSTRAINT pkAllow PRIMARY KEY, " & _
             "MinGroupFYC CURRENCY, MinActiveLP DOUBLE, AllowancePerDay CURRENCY, Description TEXT(255))"
    Exec db, "CREATE TABLE tblParam_K2 (MinK2 DOUBLE CONSTRAINT pkK2 PRIMARY KEY, Coef DOUBLE, Description TEXT(255))"
    Exec db, "CREATE TABLE tblParam_LPCount (MinTurns DOUBLE CONSTRAINT pkLPCount PRIMARY KEY, Coef DOUBLE, Description TEXT(255))"

    ' ---- Tham số chung (mục A.I văn bản 006) ----
    For Each r In Array( _
        Array("StdDays", 22, "Hệ số ngày làm việc = %attendance × 22 (005 mục D.z)"), _
        Array("MinAttendancePct", 0, "Điều kiện chấm công tối thiểu (%); 0 = không áp dụng (006 không quy định cho PM)"), _
        Array("MinWorkDays", 0, "... hoặc số ngày đi làm tối thiểu; 0 = không áp dụng"), _
        Array("NewGuarAllowance", 500000, "Bảo đảm 0,5 triệu/ngày: PM ≤ 12 tháng kể từ thăng tiến, quý xét ≥ 8 lượt LP hoạt động"), _
        Array("NewGuarMonths", 12, "Số tháng kể từ thăng tiến/bổ nhiệm PM để áp mức bảo đảm 0,5 triệu"), _
        Array("NewGuarMinLPTurns", 8, "Lượt LP hoạt động tối thiểu trong quý xét (mức 0,5 triệu)"), _
        Array("QtrGuarAllowance", 150000, "Bảo đảm 0,15 triệu/ngày: PM khác, quý xét ≥ 6 lượt LP hoạt động"), _
        Array("QtrGuarMinLPTurns", 6, "Lượt LP hoạt động tối thiểu trong quý xét (mức 0,15 triệu)"), _
        Array("JoinGuarAllowance", 500000, "PM gia nhập đến hết Q2/2027: bảo đảm 0,5 triệu/ngày trong 6 tháng đầu"), _
        Array("JoinGuarMonths", 6, "Số tháng đầu kể từ gia nhập được bảo đảm"), _
        Array("JoinGuarDeadline", 202706, "Tháng gia nhập muộn nhất được hưởng bảo đảm 6 tháng (YYYYMM)"), _
        Array("K2_NA_Coef", 1, "%Hệ số K2 khi chưa có K2 (N/A) = 100%"))
        Exec db, "INSERT INTO tblParam VALUES (" & Q(r(0)) & ", " & N(r(1)) & ", " & Q(r(2)) & ")"
    Next

    ' ---- Phụ cấp/ngày làm việc theo FYC nhóm trực tiếp và số LP hoạt động ----
    For Each r In Array( _
        Array(10000000, 2, 300000, "FYC nhóm ≥ 10 triệu và ≥ 2 LP hoạt động"), _
        Array(30000000, 2, 500000, "FYC nhóm ≥ 30 triệu và ≥ 2 LP hoạt động"), _
        Array(60000000, 2, 800000, "FYC nhóm ≥ 60 triệu và ≥ 2 LP hoạt động"))
        Exec db, "INSERT INTO tblParam_Allowance (MinGroupFYC, MinActiveLP, AllowancePerDay, Description) VALUES (" & _
                 N(r(0)) & ", " & N(r(1)) & ", " & N(r(2)) & ", " & Q(r(3)) & ")"
    Next

    ' ---- %Hệ số K2 ----
    For Each r In Array(Array(80, 1.2, "K2 ≥ 80%"), Array(70, 1.1, "K2 ≥ 70%"), _
                        Array(60, 1, "K2 ≥ 60% (N/A dùng tham số K2_NA_Coef)"), Array(0, 0.5, "K2 < 60%"))
        Exec db, "INSERT INTO tblParam_K2 VALUES (" & N(r(0)) & ", " & N(r(1)) & ", " & Q(r(2)) & ")"
    Next

    ' ---- %Hệ số Lượt LP nhóm trực tiếp hoạt động (cộng dồn từ khi thăng tiến) ----
    For Each r In Array(Array(100, 1.15, "≥ 100 lượt"), Array(50, 1.05, "≥ 50 lượt"), Array(0, 1, "< 50 lượt"))
        Exec db, "INSERT INTO tblParam_LPCount VALUES (" & N(r(0)) & ", " & N(r(1)) & ", " & Q(r(2)) & ")"
    Next

    ' ---- Liên kết Data chung (Tbl_Agent bắt buộc; Tbl_Attendance nếu đã có) ----
    acc.DoCmd.TransferDatabase acLink, "Microsoft Access", beFile, acTable, "Tbl_Agent", "Tbl_Agent"
    Check "Liên kết Tbl_Agent từ Data.accdb"
    acc.DoCmd.TransferDatabase acLink, "Microsoft Access", beFile, acTable, "Tbl_Attendance", "Tbl_Attendance"
    Err.Clear ' chưa có thì FaAfterUpdate tạo và liên kết

    ' ---- Chữ hiển thị (Unicode) ----
    Exec db, "CREATE TABLE tblText (TextKey TEXT(50) CONSTRAINT pkText PRIMARY KEY, TextValue TEXT(255))"
    For Each r In Array( _
        Array("AppTitle", "Tính Phụ cấp cố định (FA) – PM/PD"), _
        Array("E_ParamMissing", "Thiếu tham số:"), _
        Array("E_Locked", "Kỳ này đã CHỐT, không thể tính lại hoặc ghi đè."), _
        Array("E_Running", "Kỳ này đang được tính bởi:"), _
        Array("E_LaterKy", "Đã có kết quả kỳ chi trả sau của tháng này. Không tính lại kỳ trước để tránh lệch số đã chi."), _
        Array("Q_Recalc", "Tháng/kỳ này đã có kết quả. Tính lại và ghi đè?"), _
        Array("S_Loading", "Đang đọc tham số và dữ liệu..."), _
        Array("S_Calculating", "Đang tính... đã xong"), _
        Array("S_Done", "Đã tính xong:"), _
        Array("S_TotalPayable", "Tổng chi kỳ này (VND):"), _
        Array("E_Calc", "Lỗi khi tính FA. Đã hoàn tác, dữ liệu không thay đổi."), _
        Array("R_NoAgent", "Không có trong danh sách Agent"), _
        Array("R_NotAppointed", "Ngày bổ nhiệm sau tháng tính"), _
        Array("R_NotPM", "Chức danh không phải PM/PD (FA nhóm trực tiếp chỉ áp dụng PM/PD)"), _
        Array("E_AgentCols", "Sheet Agent thiếu cột bắt buộc: mã đại lý (AGENT_CODE), ngày bổ nhiệm (APPOINTMENT_DATE), mô hình (AGENCY_TYPE)."), _
        Array("S_AgentsImported", "Đã nạp danh sách agent, số dòng:"), _
        Array("R_Terminated", "HĐ đại lý chấm dứt trong/trước tháng tính"), _
        Array("R_SE", "Tình trạng SE cuối tháng"), _
        Array("R_Attendance", "Không đạt chấm công tối thiểu"), _
        Array("D_PickExcel", "Chọn file Excel dữ liệu (sheet Agent, Holiday, FAPM_Input hoặc file chấm công)"), _
        Array("S_Importing", "Đang nhập dữ liệu..."), _
        Array("E_ImportLocked", "File có dữ liệu thuộc tháng/kỳ đã CHỐT. Không nhập."), _
        Array("S_Imported", "Đã nhập FAPM_Input, số dòng:"), _
        Array("E_Import", "Lỗi khi nhập dữ liệu. Kiểm tra tên sheet và tên cột theo file mẫu."), _
        Array("E_Period", "Tháng tính không hợp lệ. Nhập dạng YYYYMM, ví dụ 202610."), _
        Array("E_Ky", "Chọn kỳ chi trả từ 1 đến 4."), _
        Array("E_NoResult", "Chưa có kết quả cho tháng/kỳ này. Bấm 'Chạy tính FA' trước."), _
        Array("D_SaveExcel", "Lưu bảng kê FA PM/PD"), _
        Array("S_Exported", "Đã xuất bảng kê:"), _
        Array("E_Export", "Lỗi khi xuất Excel."), _
        Array("E_LockNotCalc", "Chỉ chốt được kỳ đã tính xong."), _
        Array("Q_Lock", "Sau khi chốt sẽ không tính lại hay nhập đè được. Chốt kỳ"), _
        Array("S_Locked", "Đã chốt kỳ"), _
        Array("D_PickData", "Chọn file dữ liệu Data.accdb"), _
        Array("S_Relinked", "Đã kết nối lại bảng:"), _
        Array("E_Relink", "Lỗi khi kết nối file dữ liệu."))
        Exec db, "INSERT INTO tblText VALUES (" & Q(r(0)) & ", " & Q(r(1)) & ")"
    Next

    db.CreateQueryDef "qryFAPM_BangKe", "SELECT * FROM tblFAPM_Result"
    Check "Tạo query qryFAPM_BangKe"

    acc.LoadFromText acModule, "modFAPM", modFile
    Check "Nạp module modFAPM.bas"
    acc.LoadFromText acModule, "modUpdate", here & "\modUpdate.bas"
    Check "Nạp module modUpdate.bas"

    BuildForm

    SetProp db, "StartupForm", "frmMain"
    SetProp db, "AppTitle", "Tính Phụ cấp cố định (FA) – PM/PD"

    acc.RunCommand 126 ' acCmdCompileAndSaveAllModules
    If Err.Number <> 0 Then
        MsgBox "Cảnh báo: chưa biên dịch được module (" & Err.Description & ")." & vbCrLf & _
               "Mở FA_PM_App.accdb, vào VBA (Alt+F11) > Debug > Compile để xem dòng lỗi.", 48, "Dựng tool FA PM"
        Err.Clear
    End If

    acc.Run "FaAfterUpdate"
    If Err.Number <> 0 Then
        MsgBox "Dựng xong nhưng bước tạo bảng Tbl_Agent/Tbl_Attendance và form Danh sách Agent bị lỗi:" & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & "Mở FA_PM_App.accdb, bấm nút 'Cập nhật tool' để chạy lại bước này.", 48, "Dựng tool FA PM"
        Err.Clear
    End If

    acc.CloseCurrentDatabase
    Check "Đóng FA_PM_App.accdb"
End Sub

Sub SetProp(db, name, value)
    On Error Resume Next
    db.Properties(name) = value
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(name, dbText, value)
    End If
    Check "Thuộc tính " & name
End Sub

' =====================================================================
' FORM CHÍNH
' =====================================================================
Sub BuildForm()
    Dim frm, fName, c, top, i, btns
    On Error Resume Next
    Set frm = acc.CreateForm()
    Check "Tạo form"
    fName = frm.Name
    frm.Caption = "Tính Phụ cấp cố định (FA) – PM/PD"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.AutoCenter = True
    frm.Width = 8800
    frm.Section(acDetail).Height = 7600

    Set c = NewCtl(fName, acLabel, "lblTitle", 400, 300, 8000, 500)
    c.Caption = "PHỤ CẤP CỐ ĐỊNH (FA) NHÓM TRỰC TIẾP PM/PD – ALL/AD/2210-006 mục A.I"
    c.FontSize = 12: c.FontBold = True

    Set c = NewCtl(fName, acLabel, "lblPeriod", 400, 1000, 2600, 340)
    c.Caption = "Tháng tính (YYYYMM)"
    Set c = NewCtl(fName, acTextBox, "txtPeriod", 3100, 1000, 1600, 340)
    c.DefaultValue = "=Format(DateAdd(""m"",-1,Date()),""yyyymm"")"

    Set c = NewCtl(fName, acLabel, "lblKy", 400, 1500, 2600, 340)
    c.Caption = "Kỳ chi trả (1–4)"
    Set c = NewCtl(fName, acComboBox, "cboKy", 3100, 1500, 1600, 340)
    c.RowSourceType = "Value List"
    c.RowSource = "1;2;3;4"
    c.LimitToList = True
    c.DefaultValue = "1"

    btns = Array( _
        Array("btnImport", "1. Nhập dữ liệu từ Excel", "=BtnImport()"), _
        Array("btnRun", "2. Chạy tính FA", "=BtnRun()"), _
        Array("btnView", "3. Xem kết quả", "=BtnView()"), _
        Array("btnExport", "4. Xuất bảng kê Excel", "=BtnExport()"), _
        Array("btnLock", "5. Chốt kỳ", "=BtnLock()"), _
        Array("btnParams", "Bảng tham số chính sách", "=BtnParams()"), _
        Array("btnRelink", "Kết nối file dữ liệu", "=BtnRelink()"), _
        Array("btnUpdate", "Cập nhật tool", "=BtnUpdate()"), _
        Array("btnBackup", "Sao lưu dữ liệu", "=BtnBackup()"))
    top = 2200
    For i = 0 To UBound(btns)
        If i < 5 Then
            Set c = NewCtl(fName, acCommandButton, btns(i)(0), 400, top + i * 560, 4300, 460)
        Else
            Set c = NewCtl(fName, acCommandButton, btns(i)(0), 5000, top + (i - 5) * 560, 3400, 460)
        End If
        c.Caption = btns(i)(1)
        c.OnClick = btns(i)(2)
    Next

    Set c = NewCtl(fName, acLabel, "lblStatus", 400, 5300, 8000, 1600)
    c.Caption = "Sẵn sàng. Thứ tự: Nhập dữ liệu → Chạy tính → Xem/Xuất → Chốt kỳ."
    c.ForeColor = RGB(31, 78, 121)

    acc.DoCmd.Close acForm, fName, acSaveYes
    Check "Lưu form"
    acc.DoCmd.Rename "frmMain", acForm, fName
    Check "Đổi tên form"
End Sub

Function NewCtl(fName, ctlType, ctlName, l, t, w, h)
    Dim c
    On Error Resume Next
    Set c = acc.CreateControl(fName, ctlType, acDetail, "", "", l, t, w, h)
    Check "Tạo control " & ctlName
    c.Name = ctlName
    c.FontName = "Segoe UI"
    If ctlType <> acCommandButton Then c.FontSize = 10
    Set NewCtl = c
End Function
