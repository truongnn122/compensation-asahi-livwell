Option Compare Database
Option Explicit

' =====================================================================
' CAP NHAT TOOL & SAO LUU DU LIEU (file Data chung + file tool)
' Module nay it thay doi. Moi lan nang cap chi thay modFA.bas:
'   chep modFA.bas moi vao cung thu muc voi FA_App.accdb -> bam "Cap nhat tool".
' Lan dau cai dat: Alt+F11 > File > Import File > modUpdate.bas,
'   dat con tro trong Sub CaiDatCapNhat va bam F5.
' Chu tieng Viet viet dang \uXXXX vi module chi chua ky tu ASCII.
' =====================================================================

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Const M_UPD_MISSING As String = "Ch\u01B0a c\u00F3 file modFA.bas m\u1EDBi trong th\u01B0 m\u1EE5c:"
Private Const M_UPD_CONFIRM As String = "C\u1EADp nh\u1EADt tool t\u1EEB file modFA.bas? D\u1EEF li\u1EC7u s\u1EBD \u0111\u01B0\u1EE3c sao l\u01B0u tr\u01B0\u1EDBc, kh\u00F4ng b\u1ECB thay \u0111\u1ED5i."
Private Const M_UPD_DONE As String = "\u0110\u00E3 c\u1EADp nh\u1EADt tool l\u00EAn phi\u00EAn b\u1EA3n:"
Private Const M_UPD_ERR As String = "L\u1ED7i khi c\u1EADp nh\u1EADt tool. D\u1EEF li\u1EC7u kh\u00F4ng b\u1ECB \u1EA3nh h\u01B0\u1EDFng."
Private Const M_BK_DONE As String = "\u0110\u00E3 sao l\u01B0u d\u1EEF li\u1EC7u v\u00E0o:"
Private Const M_BK_ERR As String = "Kh\u00F4ng sao l\u01B0u \u0111\u01B0\u1EE3c d\u1EEF li\u1EC7u:"
Private Const M_BTN_UPDATE As String = "C\u1EADp nh\u1EADt tool"
Private Const M_BTN_BACKUP As String = "Sao l\u01B0u d\u1EEF li\u1EC7u"
Private Const M_SETUP_DONE As String = "\u0110\u00E3 th\u00EAm n\u00FAt 'C\u1EADp nh\u1EADt tool' v\u00E0 'Sao l\u01B0u d\u1EEF li\u1EC7u' v\u00E0o form ch\u00EDnh."
Private Const M_TITLE As String = "Tool FA"

Private Function UX(ByVal s As String) As String
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
    UX = out
End Function

Private Function Msg(ByVal text As String, Optional ByVal flags As Long = 64) As Long
    Msg = MessageBoxW(Application.hWndAccessApp, StrPtr(text), StrPtr(UX(M_TITLE)), flags)
End Function

Private Function DataFilePath() As String
    Dim td As DAO.TableDef, c As String
    For Each td In CurrentDb.TableDefs
        c = td.Connect
        If InStr(1, c, "DATABASE=", vbTextCompare) > 0 Then
            DataFilePath = Mid$(c, InStr(1, c, "DATABASE=", vbTextCompare) + 9)
            Exit Function
        End If
    Next td
End Function

' Sao luu file Data chung va file tool (chua tham so + ket qua thuong) vao thu muc Backup canh file Data.
' Tra ve duong dan ban sao file Data ("" neu loi).
Public Function BackupData(Optional ByVal silent As Boolean = False) As String
    Dim fso As Object, src As String, folder As String, stamp As String, dst As String
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    src = DataFilePath()
    If src = "" Then src = CurrentProject.FullName
    folder = fso.GetParentFolderName(src) & "\Backup"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    stamp = Format$(Now(), "yyyymmdd_hhnnss")
    dst = folder & "\" & fso.GetBaseName(src) & "_" & stamp & ".accdb"
    fso.CopyFile src, dst, False
    If src <> CurrentProject.FullName Then
        fso.CopyFile CurrentProject.FullName, folder & "\" & fso.GetBaseName(CurrentProject.FullName) & "_" & stamp & ".accdb", False
    End If
    BackupData = dst
    If Not silent Then Msg UX(M_BK_DONE) & vbCrLf & folder
    Exit Function
Fail:
    Msg UX(M_BK_ERR) & vbCrLf & Err.Description, 16
End Function

Public Function BtnBackup()
    BackupData False
End Function

Public Function BtnUpdate()
    Dim path As String, ver As String
    On Error GoTo Fail
    path = CurrentProject.Path & "\modFA.bas"
    If Dir$(path) = "" Then
        Msg UX(M_UPD_MISSING) & vbCrLf & CurrentProject.Path, 48
        Exit Function
    End If
    If Msg(UX(M_UPD_CONFIRM), 4 Or 32) <> 6 Then Exit Function
    If BackupData(True) = "" Then Exit Function

    DoCmd.Close acForm, "frmMain"
    Application.LoadFromText acModule, "modFA", path
    RunCommand acCmdCompileAndSaveAllModules
    Name path As path & ".da_cap_nhat_" & Format$(Now(), "yyyymmdd_hhnnss")

    On Error Resume Next
    Application.Run "FaAfterUpdate"
    ver = Application.Run("FaVersion")
    On Error GoTo 0
    DoCmd.OpenForm "frmMain"
    Msg UX(M_UPD_DONE) & " " & ver
    Exit Function
Fail:
    Msg UX(M_UPD_ERR) & vbCrLf & Err.Description, 16
    On Error Resume Next
    DoCmd.OpenForm "frmMain"
End Function

' Chay 1 lan (F5 trong cua so VBA) de them 2 nut vao form chinh cua ban da dung truoc do.
Public Sub CaiDatCapNhat()
    Dim frm As Form, c As Control, has1 As Boolean, has2 As Boolean
    DoCmd.OpenForm "frmMain", acDesign, , , , acHidden
    Set frm = Forms("frmMain")
    For Each c In frm.Controls
        If c.Name = "btnUpdate" Then has1 = True
        If c.Name = "btnBackup" Then has2 = True
    Next c
    If Not has1 Then
        Set c = CreateControl("frmMain", acCommandButton, acDetail, , , 5000, 3320, 3400, 460)
        c.Name = "btnUpdate": c.Caption = UX(M_BTN_UPDATE): c.OnClick = "=BtnUpdate()": c.FontName = "Segoe UI"
    End If
    If Not has2 Then
        Set c = CreateControl("frmMain", acCommandButton, acDetail, , , 5000, 3880, 3400, 460)
        c.Name = "btnBackup": c.Caption = UX(M_BTN_BACKUP): c.OnClick = "=BtnBackup()": c.FontName = "Segoe UI"
    End If
    DoCmd.Close acForm, "frmMain", acSaveYes
    Msg UX(M_SETUP_DONE)
End Sub
