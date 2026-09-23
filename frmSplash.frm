VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ch­¬ng tr×nh kÕ to¸n Sao ViÖt"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   1605
   ClientWidth     =   8280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":57E2
   ScaleHeight     =   341.25
   ScaleMode       =   0  'User
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "3"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2760
      Top             =   3600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   7800
      TabIndex        =   0
      Top             =   6960
      Width           =   3855
      Begin VB.Label LbAbout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Warning:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Tag             =   "Warning"
         Top             =   0
         Width           =   675
      End
      Begin VB.Label LbAbout 
         BackColor       =   &H80000013&
         Caption         =   $"frmSplash.frx":151AC
         Height          =   735
         Index           =   7
         Left            =   750
         TabIndex        =   5
         Tag             =   "Copyright"
         Top             =   2955
         Width           =   5895
      End
      Begin VB.Label LbAbout 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Accounting Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Index           =   2
         Left            =   3075
         TabIndex        =   4
         Tag             =   "Product"
         Top             =   435
         Width           =   3585
      End
      Begin VB.Label LbAbout 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "SAO VIET Software Center"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3195
         TabIndex        =   3
         Tag             =   "CompanyProduct"
         Top             =   0
         Width           =   3750
      End
      Begin VB.Label LbAbout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Version 6.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   5295
         TabIndex        =   2
         Tag             =   "Version"
         Top             =   1335
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "WWW.SAOVIET.COM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Top             =   1545
         Width           =   2010
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "dlgCommonDialog"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   300
      Width           =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   366
      X2              =   366
      Y1              =   -1000
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   0
      X2              =   366
      Y1              =   -1000
      Y2              =   -1000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   228
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   20
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type versionInfo
    Version As String
    fileId As String
    fileName As String
    ReleaseDate As String
    Description As String
End Type
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                           (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                                            ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private LastVersion As String
Public Ready As Boolean
Public Sub StartSAS()

'SaveSetting "MyApp", "Settings", "FirstRun", "True"
    pDataPath = GetSetting(IniPath, "Environment", "Path", pCurDir + "DATA\KETOAN.MDB")

    Dim folderPath As String
    Dim FilePath As String

    folderPath = App.path & "\Hoadon"
    FilePath = folderPath & "\vbPath.txt"

    ' N?u chua có file vbPath.txt thì t?o
    If Dir$(FilePath) = "" Then
        Open FilePath For Output As #1
        Close #1
    End If

    ' §Æt c¸c format
    'pDataPath = "C:\Tao moi\DATA\KT.MDB"
    Mask_0 = GetSetting(IniPath, "Environment", "IntMask", "###,###,###,###")
    If Cdbl5("1,5") <> 1.5 Then sDecimal = "." Else sDecimal = ","
    pThang = GetSetting(IniPath, "Environment", "NDecimal", 2)
    Select Case pThang
    Case 0: Mask_2 = Mask_0
    Case 1: Mask_2 = "###0.0"
    Case 2: Mask_2 = "Standard"
    Case 3: Mask_2 = "###0.000"
    Case 4: Mask_2 = "###0.0000"
    Case Else: Mask_2 = GetSetting(IniPath, "Environment", "DblMask", "Standard")
    End Select
    Mask_N = 10 ^ pThang
    Mask_D = GetShortDateFormat
    ' §äc File INI
    pThangDauKy = GetSetting(IniPath, "Environment", "StartMonth", "1")
    pTKTrunggian = GetSetting(IniPath, "Environment", "TmpAccount", "5")
    SHCT_Len = GetSetting(IniPath, "Environment", "InvCodeLen", 2)

    ShTkSPDo = GetSetting(IniPath, "ProductCost", "TK_SPDO", "154")
    ShTkTP = GetSetting(IniPath, "ProductCost", "TK_SPDO", "155")
    ShTkKQ = GetSetting(IniPath, "ProductCost", "TK_XDKQ", "911")

    Set WSpace = DBEngine.CreateWorkspace(CStr(Time), "Admin", "", dbUseJet)
    Workspaces.Append WSpace
  
    Do While OpenDB(pDataPath) <> 0
        dlgCommonDialog.Flags = &H4&
        dlgCommonDialog.fileName = "*.MDB"
        dlgCommonDialog.InitDir = App.path & "\Data"

        On Error GoTo QuitSAS
        dlgCommonDialog.ShowOpen
        On Error GoTo 0
        pDataPath = dlgCommonDialog.fileName
        Open FilePath For Output As #1
        Print #1, pDataPath;
        Close #1
    Loop

    pThang = month(Date)
    pWinDir = GetWinDir

    Select Case UCase(App.ExeName)
    Case "SERVER": pProcessMode = 2
    Case "CLIENT": pProcessMode = 1
    Case Else: pProcessMode = 0
    End Select

    Exit Sub
QuitSAS:
    WSpace.Close
    End
End Sub
Private Sub KiemtraversionDrive()
    Dim fileId As String
    fileId = "1fXXVQCw2XHusJ9y5zwIt0jPJTCL4dxc2"

    Dim fileName As String
    fileName = "version.config"

    Dim exeDir As String
    exeDir = App.path  ' D:\DA3\DA3
    Dim extractFolder As String
    extractFolder = exeDir & "\Tools\Debug"

    Dim zipPath As String
    zipPath = extractFolder & "\" & fileName  ' D:\DA3\DA3\Tools\Debug\VietstarDriver.zip
    Dim retryCount As Integer
    Dim downloadUrl As String
    Dim Result As Long
    ' Xóa file cu
    If Dir(zipPath) <> "" Then Kill zipPath

    Dim urls(2) As String
    urls(0) = "https://drive.google.com/uc?export=download&id=" & fileId
    urls(1) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=1"
    urls(2) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=t"

    Dim success As Boolean
    success = False

    For retryCount = 0 To 2
        downloadUrl = urls(retryCount)
        Result = URLDownloadToFile(0, downloadUrl, zipPath, 0, 0)

        If Result = 0 Then
            success = True
            Exit For
        End If

        Sleep 1000
    Next retryCount

    If Not success Then
        'MsgBox "Không th? t?i file!", vbExclamation
        Exit Sub
    End If


    'Load menu
    Dim json As String
    Dim versions As Collection
    Dim i As Long


    'Lay version hien tai

    Dim txtPath As String
    txtPath = App.path & "\" & "HoaDon\version.txt"
    Dim content As String
    content = ReadTxt(txtPath)

    'Kiem tra version hien tai có trong danh sách không co thi tai ban cap nhat moi nhat
    Dim versionData As versionInfo
    Docfilejson content, versionData
    If versionData.Version = "" Then
        'Lay version moi nhat
        DownloadAndRun
    End If
    Docfilejson "1.1", versionData
    LastVersion = Replace(versionData.fileName, """", "")
    Ready = True
End Sub
Private Sub DownloadAndRun()
    On Error GoTo ErrorHandler

    Dim exeDir As String
    Dim fileId As String
    Dim fileName As String
    Dim zipPath As String
    Dim extractFolder As String
    Dim exePath As String
    Dim downloadUrl As String
    Dim Result As Long
    Dim retryCount As Integer

    ' ===== THI?T L?P ÐU?NG D?N =====
    exeDir = App.path  ' D:\DA3\DA3
    fileId = "1qj3mjQkA1o6QEsellIQ09_vrzq1ETCOL"
    fileName = "VietstarDriver.zip"

    ' Thu m?c gi?i nén (Tools\Debug)
    extractFolder = exeDir & "\Tools\Debug"
    zipPath = extractFolder & "\" & fileName  ' D:\DA3\DA3\Tools\Debug\VietstarDriver.zip
    exePath = extractFolder & "\VietstarDriver.exe"

    ' Xóa file cu
    If Dir(zipPath) <> "" Then Kill zipPath
    ' T?o thu m?c n?u chua có
    If Dir(extractFolder, vbDirectory) = "" Then
        MkDir extractFolder
    End If
    ' ===== T?I FILE =====
    Dim urls(2) As String
    urls(0) = "https://drive.google.com/uc?export=download&id=" & fileId
    urls(1) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=1"
    urls(2) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=t"

    Dim success As Boolean
    success = False

    For retryCount = 0 To 2
        downloadUrl = urls(retryCount)
        Result = URLDownloadToFile(0, downloadUrl, zipPath, 0, 0)

        If Result = 0 Then
            If FileLen(zipPath) > 1024 Then
                success = True
                Exit For
            Else
                Kill zipPath
            End If
        End If

        Sleep 1000
    Next retryCount

    If Not success Then
        MsgBox "Không th? t?i file!", vbExclamation
        Exit Sub
    End If

    ' ===== GI?I NÉN VÀO TOOLS\DEBUG =====
    If Dir(zipPath) <> "" Then
        UnzipFile zipPath, extractFolder  ' Gi?i nén vào Tools\Debug
        If Dir(zipPath) <> "" Then Kill zipPath
    End If

    ' ===== CH?Y FILE EXE =====
    If Dir(exePath) <> "" Then
        Shell exePath, vbNormalFocus
        Unload Me
        End
    Else
        MsgBox "Không tìm th?y file EXE!", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "L?i: " & Err.Description, vbCritical
End Sub
Private Sub Docfilejson(ByVal versionid As String, ByRef versionData As versionInfo)

    Dim json As String
    Dim pVersion As Long

    json = ReadJsonFile(App.path & "\Tools\Debug\version.config")

    If json = "" Then
        MsgBox "Không d?c du?c file JSON"
        Exit Sub
    End If

    'Tìm dúng version du?c truy?n vào
    pVersion = InStr(1, json, """" & versionid & """", vbTextCompare)

    If pVersion = 0 Then
        'MsgBox "Không tìm th?y version " & versionid
        Exit Sub
    End If

    versionData.Version = versionid

    versionData.fileId = GetJsonValue(json, "fileId", pVersion)
    versionData.fileName = GetJsonValue(json, "fileName", pVersion)
    versionData.ReleaseDate = GetJsonValue(json, "releaseDate", pVersion)
    versionData.Description = GetJsonValue(json, "description", pVersion)

End Sub
Private Function ReadTxt(FilePath As String) As String
    Dim fileNumber As Integer
    Dim content As String

    On Error GoTo ErrorHandler

    ' Ki?m tra file t?n t?i
    If Dir(FilePath) = "" Then
        ReadTxt = "File không t?n t?i: " & FilePath
        Exit Function
    End If

    ' L?y file number
    fileNumber = FreeFile

    ' M? file d? d?c
    Open FilePath For Input As #fileNumber

    ' Ð?c toàn b? n?i dung
    If LOF(fileNumber) > 0 Then
        content = Input$(LOF(fileNumber), #fileNumber)
    End If

    ' Ðóng file
    Close #fileNumber

    ' Tr? v? n?i dung
    ReadTxt = content
    Exit Function

ErrorHandler:
    ReadTxt = "L?i d?c file: " & Err.Description
    On Error Resume Next
    Close #fileNumber
End Function
Private Sub Form_Load()
    Ready = False
    'Timer1.Enabled = True
    Dim sCmdLine As String
    sCmdLine = Command$    ' L?y toàn b? tham s?

    'Doc file mod txt
    Dim FilePath As String
    FilePath = App.path & "\Hoadon\Mode.txt"
    If Dir(FilePath) <> "" Then
        ' Ð?c file
        Dim FileNum As Integer
        Dim noiDung As String

        FileNum = FreeFile
        Open FilePath For Input As #FileNum
        noiDung = Input$(LOF(FileNum), FileNum)
        Close #FileNum
    Else
        noiDung = 0
    End If

    If InStr(1, sCmdLine, "/AutoMode", vbTextCompare) > 0 Or noiDung = 2 Then
        modStatic = 2
        Me.Height = 0
        Me.Width = 0
    Else
        modStatic = 1
    End If

    ' Me.Height = 5280
    'Me.Width = 8250
    dlgCommonDialog.InitDir = pCurDir + "DATA"
    If DEMO = 1 Then LbAbout(4).Caption = "Training Version"
End Sub
Private Sub UnzipFile(zipPath As String, extractFolder As String)
    On Error GoTo ErrorHandler

    Dim cmd As String
    Dim wsh As Object
    Dim Result As Long
 

    ' Ki?m tra file ZIP
    If Dir(zipPath) = "" Then
        MsgBox "Không tìm th?y file ZIP: " & zipPath, vbExclamation
        Exit Sub
    End If

    ' T?o thu m?c dích
    If Dir(extractFolder, vbDirectory) = "" Then
        MkDir extractFolder
    End If

    ' Ki?m tra file ZIP
    If FileLen(zipPath) < 1024 Then
        MsgBox "File ZIP b? l?i ho?c quá nh?!", vbExclamation
        Kill zipPath
        Exit Sub
    End If

    ' PowerShell
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
          """Expand-Archive -LiteralPath '" & Replace(zipPath, "'", "''") & _
          "' -DestinationPath '" & Replace(extractFolder, "'", "''") & _
          "' -Force"""

    Set wsh = CreateObject("WScript.Shell")

    ' 0 = ?n c?a s?
    ' True = CH? PowerShell ch?y xong
    Result = wsh.Run(cmd, 0, True)

    If Result <> 0 Then
        MsgBox "Gi?i nén th?t b?i. Mã l?i: " & Result, vbExclamation
        Exit Sub
    End If
 

    Set wsh = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "L?i gi?i nén: " & Err.number & " - " & Err.Description, vbExclamation
End Sub
Private Function ReadJsonFile(ByVal FilePath As String) As String
    Dim f As Integer
    Dim s As String

    If Dir(FilePath) = "" Then Exit Function

    f = FreeFile

    Open FilePath For Input As #f
        s = Input$(LOF(f), f)
    Close #f

    ReadJsonFile = s
End Function
Private Function GetJsonValue(ByVal json As String, _
                              ByVal PropertyName As String, _
                              ByVal startPos As Long) As String

    Dim p As Long
    Dim pColon As Long
    Dim pStart As Long
    Dim pEnd As Long
    Dim SearchText As String

    SearchText = """" & PropertyName & """"

    p = InStr(startPos, json, SearchText, vbTextCompare)

    If p = 0 Then Exit Function

    pColon = InStr(p + Len(SearchText), json, ":")

    If pColon = 0 Then Exit Function

    ' B? kho?ng tr?ng và d?u "
    pStart = pColon + 1

    Do While pStart <= Len(json)
        If Mid$(json, pStart, 1) <> " " _
           And Mid$(json, pStart, 1) <> vbTab _
           And Mid$(json, pStart, 1) <> vbCr _
           And Mid$(json, pStart, 1) <> vbLf _
           And Mid$(json, pStart, 1) <> """" Then
            Exit Do
        End If

        pStart = pStart + 1
    Loop

    ' N?u giá tr? b?t d?u b?ng "
    If Mid$(json, pColon + 1, 1) = """" Then

        pStart = pColon + 2

        pEnd = InStr(pStart, json, """")

        If pEnd = 0 Then Exit Function

        GetJsonValue = Mid$(json, pStart, pEnd - pStart)

    Else

        pEnd = InStr(pStart, json, ",")

        If pEnd = 0 Then
            pEnd = InStr(pStart, json, "}")
        End If

        If pEnd = 0 Then Exit Function

        GetJsonValue = Trim$(Mid$(json, pStart, pEnd - pStart))

    End If

End Function
Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
  
    KiemtraversionDrive
End Sub
