' =====================================================================
' DỰNG TOOL TÍNH PHỤ CẤP CỐ ĐỊNH (FA) – ALL/AD/2210-005 mục A.II
' Chạy: double-click Build_FA_Tool.vbs (cùng thư mục với modFA.bas)
' Kết quả: Data.accdb (dữ liệu dùng chung) + FA_App.accdb (tool FA kèm dữ liệu riêng của FA)
' Yêu cầu: máy Windows có Microsoft Access 2010 trở lên (hoặc Access Runtime không đủ – cần bản đầy đủ để dựng)
' =====================================================================
Option Explicit

Const acLink = 2, acTable = 0, acModule = 5, acForm = 2, acSaveYes = 1, acDetail = 0
Const acLabel = 100, acTextBox = 109, acComboBox = 111, acCommandButton = 104
Const dbText = 10

Dim fso, here, beFile, feFile, modFile, acc, step
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
beFile = here & "\Data.accdb"
feFile = here & "\FA_App.accdb"
modFile = here & "\modFA.bas"

If Not fso.FileExists(modFile) Then
    MsgBox "Không tìm thấy modFA.bas cùng thư mục với script.", 16, "Dựng tool FA"
    WScript.Quit 1
End If
If fso.FileExists(beFile) Or fso.FileExists(feFile) Then
    MsgBox "Đã có Data.accdb hoặc FA_App.accdb trong thư mục này." & vbCrLf & _
           "Đổi tên/di chuyển file cũ rồi chạy lại để tránh ghi đè dữ liệu.", 48, "Dựng tool FA"
    WScript.Quit 1
End If

On Error Resume Next
Set acc = CreateObject("Access.Application")
Check "Mở Microsoft Access"

BuildBackEnd
BuildFrontEnd

acc.Quit
MsgBox "Hoàn tất." & vbCrLf & vbCrLf & _
       "• Data.accdb: dữ liệu dùng chung cho mọi khoản thưởng (Tbl_Agent, Tbl_Attendance)" & vbCrLf & _
       "• FA_App.accdb: mở file này để sử dụng" & vbCrLf & vbCrLf & _
       "Lần đầu mở, bấm 'Enable Content' hoặc thêm thư mục vào Trusted Locations.", 64, "Dựng tool FA"

' =====================================================================
Sub Check(stepName)
    If Err.Number <> 0 Then
        MsgBox "Lỗi ở bước: " & stepName & vbCrLf & Err.Number & " – " & Err.Description, 16, "Dựng tool FA"
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
' BACK-END: bảng dữ liệu + tham số chính sách
' =====================================================================
Sub BuildBackEnd()
    Dim db, r
    On Error Resume Next
    acc.NewCurrentDatabase beFile
    Check "Tạo Data.accdb"
    Set db = acc.CurrentDb

    Exec db, "CREATE TABLE tblHoliday (HolidayDate DATETIME CONSTRAINT pkHoliday PRIMARY KEY, Description TEXT(100))"
    Exec db, "CREATE TABLE tblFA_Input (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, AgentCode TEXT(20) NOT NULL, " & _
             "ActualDays DOUBLE, FYC CURRENCY, Contracts DOUBLE, Recruits INTEGER, K2Pct DOUBLE, " & _
             "MDRTPrev YESNO, TrainingDone YESNO, StatusEOM TEXT(20), " & _
             "CONSTRAINT pkFAInput PRIMARY KEY (Period, Ky, AgentCode))"
    Exec db, "CREATE TABLE tblFA_Result (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, AgentCode TEXT(20) NOT NULL, " & _
             "WorkMonthNo INTEGER, WorkDaysBase INTEGER, ActualDays DOUBLE, AttendancePct DOUBLE, DayFactor DOUBLE, " & _
             "Eligible YESNO, Reason TEXT(255), FYC CURRENCY, Contracts DOUBLE, Recruits INTEGER, " & _
             "AllowancePerDay CURRENCY, K2Pct DOUBLE, K2Coef DOUBLE, MDRTCoef DOUBLE, ContractCoef DOUBLE, " & _
             "FA_Formula CURRENCY, NewGuar CURRENCY, NewGuarPending CURRENCY, NewGuarRetro CURRENCY, FA_Month CURRENCY, " & _
             "QtrGuarFlag YESNO, QtrGuarMonth CURRENCY, QtrGuarTopup CURRENCY, FA_Total CURRENCY, " & _
             "PaidPrevKy CURRENCY, Payable CURRENCY, CalcAt DATETIME, CalcBy TEXT(100), " & _
             "CONSTRAINT pkFAResult PRIMARY KEY (Period, Ky, AgentCode))"
    Exec db, "CREATE TABLE tblPeriodStatus (Period TEXT(6) NOT NULL, Ky INTEGER NOT NULL, Status TEXT(10), " & _
             "UpdatedBy TEXT(100), UpdatedAt DATETIME, CONSTRAINT pkPeriodStatus PRIMARY KEY (Period, Ky))"
    Exec db, "CREATE TABLE tblParam (ParamKey TEXT(50) CONSTRAINT pkParam PRIMARY KEY, ParamValue DOUBLE, Description TEXT(255))"
    Exec db, "CREATE TABLE tblParam_Allowance (ID COUNTER CONSTRAINT pkAllow PRIMARY KEY, OptionNo INTEGER, " & _
             "MinFYC CURRENCY, MinRecruits INTEGER, AllowancePerDay CURRENCY, Description TEXT(255))"
    Exec db, "CREATE TABLE tblParam_K2 (MinK2 DOUBLE CONSTRAINT pkK2 PRIMARY KEY, Coef DOUBLE, Description TEXT(255))"
    Exec db, "CREATE TABLE tblParam_Contract (MinContracts DOUBLE CONSTRAINT pkCt PRIMARY KEY, Coef DOUBLE, Description TEXT(255))"

    ' ---- Tham số chung ----
    For Each r In Array( _
        Array("StdDays", 22, "Hệ số ngày làm việc = %attendance × 22 (mục D.z)"), _
        Array("MinAttendancePct", 70, "Điều kiện tối thiểu: %attendance ≥ 70% ... (mục A.II)"), _
        Array("MinWorkDays", 15, "... hoặc ≥ 15 ngày làm việc trong tháng"), _
        Array("NewGuarAllowance", 300000, "Bảo đảm 0,3 triệu/ngày cho LP mới"), _
        Array("NewGuarMonths", 2, "Số tháng làm việc đầu được bảo đảm"), _
        Array("NewGuarMinContracts", 0.5, "Số HĐ tối thiểu trong 2 tháng đầu để nhận bảo đảm"), _
        Array("QtrGuarAllowance", 150000, "Bảo đảm FA quý kế tiếp: 0,15 triệu/ngày"), _
        Array("QtrGuarMinContracts", 6, "Điều kiện bảo đảm quý: ≥ 6 HĐ phát sinh trong quý"), _
        Array("QtrGuarMinFYC", 24000000, "Điều kiện bảo đảm quý: ≥ 24 triệu FYC phát sinh trong quý"), _
        Array("K2_NA_Coef", 1, "%Hệ số K2 khi chưa có K2 (N/A) = 100%"), _
        Array("MDRT_Coef", 1.2, "%Hệ số MDRT khi đạt MDRT năm liền kề trước = 120%"))
        Exec db, "INSERT INTO tblParam VALUES (" & Q(r(0)) & ", " & N(r(1)) & ", " & Q(r(2)) & ")"
    Next

    ' ---- Mục A: Phụ cấp/ngày làm việc (lấy mức cao nhất LP thỏa) ----
    For Each r In Array( _
        Array(1, 0, 0, 200000, "Lựa chọn 1: FYC < 15 triệu"), _
        Array(1, 15000000, 0, 300000, "Lựa chọn 1: FYC ≥ 15 triệu"), _
        Array(1, 25000000, 0, 500000, "Lựa chọn 1: FYC ≥ 25 triệu"), _
        Array(2, 0, 1, 300000, "Lựa chọn 2: FYC < 15 triệu và ≥ 1 LP giới thiệu"), _
        Array(2, 15000000, 1, 500000, "Lựa chọn 2: FYC ≥ 15 triệu và ≥ 1 LP giới thiệu"), _
        Array(2, 25000000, 1, 700000, "Lựa chọn 2: FYC ≥ 25 triệu và ≥ 1 LP giới thiệu"))
        Exec db, "INSERT INTO tblParam_Allowance (OptionNo, MinFYC, MinRecruits, AllowancePerDay, Description) VALUES (" & _
                 N(r(0)) & ", " & N(r(1)) & ", " & N(r(2)) & ", " & N(r(3)) & ", " & Q(r(4)) & ")"
    Next

    ' ---- Mục B: %Hệ số K2 ----
    For Each r In Array(Array(80, 1.2, "K2 ≥ 80%"), Array(70, 1.1, "K2 ≥ 70%"), _
                        Array(60, 1, "K2 ≥ 60% (N/A dùng tham số K2_NA_Coef)"), Array(0, 0.5, "K2 < 60%"))
        Exec db, "INSERT INTO tblParam_K2 VALUES (" & N(r(0)) & ", " & N(r(1)) & ", " & Q(r(2)) & ")"
    Next

    ' ---- Mục D: %Hệ số hợp đồng ----
    For Each r In Array(Array(2, 1, "≥ 2 HĐ"), Array(1, 0.7, "≥ 1 HĐ"), _
                        Array(0.5, 0.5, "≥ 0,5 HĐ"), Array(0, 0, "< 0,5 HĐ"))
        Exec db, "INSERT INTO tblParam_Contract VALUES (" & N(r(0)) & ", " & N(r(1)) & ", " & Q(r(2)) & ")"
    Next

    acc.CloseCurrentDatabase
    Check "Đóng Data.accdb"
End Sub

' =====================================================================
' FRONT-END: liên kết bảng, nhãn tiếng Việt, query, module, form
' =====================================================================
Sub BuildFrontEnd()
    Dim db, t, r
    On Error Resume Next
    acc.NewCurrentDatabase feFile
    Check "Tạo FA_App.accdb"
    Set db = acc.CurrentDb

    ' Lien ket tam cac bang FA; FaAfterUpdate chuyen chung ve file nay, Data.accdb chi con Tbl_Agent + Tbl_Attendance
    For Each t In Array("tblHoliday", "tblFA_Input", "tblFA_Result", "tblPeriodStatus", _
                        "tblParam", "tblParam_Allowance", "tblParam_K2", "tblParam_Contract")
        acc.DoCmd.TransferDatabase acLink, "Microsoft Access", beFile, acTable, t, t
        Check "Liên kết bảng " & t
    Next

    ' Chữ hiển thị (Unicode) – VBA đọc qua hàm T()
    Exec db, "CREATE TABLE tblText (TextKey TEXT(50) CONSTRAINT pkText PRIMARY KEY, TextValue TEXT(255))"
    For Each r In Array( _
        Array("AppTitle", "Tính Phụ cấp cố định (FA) – LP"), _
        Array("E_ParamMissing", "Thiếu tham số:"), _
        Array("E_Locked", "Kỳ này đã CHỐT, không thể tính lại hoặc ghi đè."), _
        Array("E_Running", "Kỳ này đang được tính bởi:"), _
        Array("E_LaterKy", "Đã có kết quả kỳ chi trả sau của tháng này. Không tính lại kỳ trước để tránh lệch số đã chi."), _
        Array("E_NoInput", "Chưa có dữ liệu FA_Input cho tháng/kỳ này."), _
        Array("Q_Recalc", "Tháng/kỳ này đã có kết quả. Tính lại và ghi đè?"), _
        Array("S_Loading", "Đang đọc tham số và dữ liệu..."), _
        Array("S_Calculating", "Đang tính... đã xong"), _
        Array("S_Done", "Đã tính xong:"), _
        Array("S_TotalPayable", "Tổng chi kỳ này (VND):"), _
        Array("E_Calc", "Lỗi khi tính FA. Đã hoàn tác, dữ liệu không thay đổi."), _
        Array("R_NoAgent", "Không có trong danh sách Agent"), _
        Array("R_NotAppointed", "Ngày bổ nhiệm sau tháng tính"), _
        Array("R_NotFT", "Không phải LP toàn thời gian"), _
        Array("R_NotLP", "Chức danh không phải LP (FA chỉ áp dụng LP toàn thời gian)"), _
        Array("E_AgentCols", "Sheet Agent thiếu cột bắt buộc: mã đại lý (AgentCode/AGENT_CODE), ngày bổ nhiệm (AppointDate/APPOINTMENT_DATE), mô hình (WorkModel/AGENCY_TYPE)."), _
        Array("S_AgentsImported", "Đã nạp danh sách agent, số dòng:"), _
        Array("R_Terminated", "HĐ đại lý chấm dứt trong/trước tháng tính"), _
        Array("R_SE", "Tình trạng SE cuối tháng"), _
        Array("R_Training", "Chưa hoàn tất huấn luyện bắt buộc"), _
        Array("R_Attendance", "Chấm công < 70% và < 15 ngày làm việc"), _
        Array("R_GuarPending", "LP mới chưa đủ 0,5 HĐ – bảo đảm chờ đến khi đạt (trong 2 tháng đầu)"), _
        Array("D_PickExcel", "Chọn file Excel dữ liệu (sheet Agent, Holiday, FA_Input)"), _
        Array("S_Importing", "Đang nhập dữ liệu..."), _
        Array("E_ImportLocked", "File có dữ liệu thuộc tháng/kỳ đã CHỐT. Không nhập."), _
        Array("S_Imported", "Đã nhập FA_Input, số dòng:"), _
        Array("E_Import", "Lỗi khi nhập dữ liệu. Kiểm tra tên sheet và tên cột theo file mẫu."), _
        Array("E_Period", "Tháng tính không hợp lệ. Nhập dạng YYYYMM, ví dụ 202610."), _
        Array("E_Ky", "Chọn kỳ chi trả từ 1 đến 4."), _
        Array("E_NoResult", "Chưa có kết quả cho tháng/kỳ này. Bấm 'Chạy tính FA' trước."), _
        Array("D_SaveExcel", "Lưu bảng kê FA"), _
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

    db.CreateQueryDef "qryFA_BangKe", "SELECT * FROM tblFA_Result"
    Check "Tạo query qryFA_BangKe"

    acc.LoadFromText acModule, "modFA", modFile
    Check "Nạp module modFA.bas"
    acc.LoadFromText acModule, "modUpdate", here & "\modUpdate.bas"
    Check "Nạp module modUpdate.bas"

    BuildForm

    SetProp db, "StartupForm", "frmMain"
    SetProp db, "AppTitle", "Tính Phụ cấp cố định (FA) – LP"

    acc.RunCommand 126 ' acCmdCompileAndSaveAllModules
    If Err.Number <> 0 Then
        MsgBox "Cảnh báo: chưa biên dịch được module (" & Err.Description & ")." & vbCrLf & _
               "Mở FA_App.accdb, vào VBA (Alt+F11) > Debug > Compile để xem dòng lỗi.", 48, "Dựng tool FA"
        Err.Clear
    End If

    acc.Run "FaAfterUpdate"
    Check "Tạo bảng/form Danh sách Agent"

    acc.CloseCurrentDatabase
    Check "Đóng FA_App.accdb"
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
    frm.Caption = "Tính Phụ cấp cố định (FA) – LP"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.AutoCenter = True
    frm.Width = 8800
    frm.Section(acDetail).Height = 7600

    Set c = NewCtl(fName, acLabel, "lblTitle", 400, 300, 8000, 500)
    c.Caption = "PHỤ CẤP CỐ ĐỊNH HÀNG THÁNG (FA) – ALL/AD/2210-005 mục A.II"
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
