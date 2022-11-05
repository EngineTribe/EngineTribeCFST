VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "引擎部落接入点选择工具"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8550
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新接入点列表"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   6330
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存更改"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   2055
   End
   Begin VB.OptionButton OptionMasterServer 
      BackColor       =   &H80000005&
      Caption         =   "手动选择"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.OptionButton OptionMasterServer 
      BackColor       =   &H80000005&
      Caption         =   "直连"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.OptionButton OptionMasterServer 
      BackColor       =   &H80000005&
      Caption         =   "默认"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   240
      Y2              =   6600
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "从右侧 CDN 接入点中手动选择"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3.4 之前使用的连接方式，不使用 CDN"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "由 CDN 选择接入点，可能会连接失败"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "服务器接入点选择"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const ET_START = "#EngineTribeStart"
Private Const ET_END = "#EngineTribeEnd"
Private Const MASTER_HOST = "juego.enginetribe.gq"
Private Const STORAGE_HOST = "storage.enginetribe.gq"
Private Const UNSET = "Unset"
Dim MASTER_IP As String


Private Sub cmdUpdate_Click()
    Kill App.Path & "\result.csv"
    DoCFST
    LoadResults
End Sub
Private Sub Form_Load()
    MASTER_IP = Chr(49) & Chr(51) & Chr(55) & Chr(46) & Chr(49) & Chr(56) & Chr(52) & Chr(46) & Chr(50) & Chr(51) & Chr(53) & Chr(46) & Chr(57) & Chr(56)
    If Dir(App.Path & "\result.csv", vbNormal) = "" Then
        DoCFST
    End If
    LoadResults
    InitHosts
    ParseHosts
End Sub

Private Sub DoCFST()
    List1.Clear
    List1.AddItem "正在更新接入点列表，"
    List1.AddItem "请稍后 ..."
    List1.Enabled = False
    frmMain.Show
    DoEvents
    Shell "cmd /c cd """ & App.Path & """ && echo . | cfst.exe -p 0 -n 500 -dd -tp 80", vbMinimizedNoFocus
    Do Until Dir(App.Path & "\result.csv", vbNormal) <> ""
        Sleep 100
        DoEvents
    Loop
    List1.Clear
    List1.Enabled = True
    Shell "taskkill /f /im cfst.exe"
End Sub

Private Sub LoadResults()
    List1.Clear
    List1.Enabled = False
    frmMain.Show
    DoEvents
    Dim SingleLine As Variant, SingleArr() As String, Counter As Integer
    Counter = 0
    For Each SingleLine In Split(ReadTextFile(App.Path & "\result.csv"), vbLf)
        SingleArr = Split(SingleLine, ",")
        Counter = Counter + 1
        If Counter = 101 Then Exit For
        If Left$(SingleLine, 2) <> "IP" Then List1.AddItem SingleArr(0) & " (" & SingleArr(4) & "ms)"
    Next
    List1.RemoveItem 0
    List1.Enabled = True
    List1.Selected(0) = True
End Sub

Private Sub InitHosts()
    On Error GoTo InitHostsErrorHandler
    Dim HostsString As String, HostsPath As String
    HostsPath = Environ("SystemRoot") & "\System32\drivers\etc\hosts"
    HostsString = ReadTextFile(HostsPath)
    If InStr(HostsString, ET_START) = 0 Then
        Open HostsPath For Output As #2
        Print #2, HostsString
        Print #2, ""
        Print #2, ET_START
        Print #2, ET_END;
        Close #2
    End If
    Exit Sub
    
InitHostsErrorHandler:
    MsgBox "请以管理员权限运行本程序!" & vbCrLf & "如果杀毒软件拦截了本程序修改 Hosts，请允许之!", vbCritical
    End
End Sub

Private Sub ParseHosts()
    Dim HostsString As String, HostsPath As String
    HostsPath = Environ("SystemRoot") & "\System32\drivers\etc\hosts"
    HostsString = ReadTextFile(HostsPath)
    
    Dim MasterServerHost As String
    Dim EngineTribeHostsSection As String
    EngineTribeHostsSection = Trim(Split(Split(HostsString, ET_START)(1), ET_END)(0))
    
    MasterServerHost = UNSET
    
    Dim SingleLine As Variant
    For Each SingleLine In Split(EngineTribeHostsSection, vbCrLf)
        If Right(SingleLine, 20) = MASTER_HOST Then MasterServerHost = Trim(Split(SingleLine, MASTER_HOST)(0))
    Next
    
    If MasterServerHost = UNSET Then
        OptionMasterServer(0).Value = True
    ElseIf MasterServerHost = MASTER_IP Then
        OptionMasterServer(1).Value = True
    Else
        OptionMasterServer(2).Value = True
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo SaveHostsErrorHandler
    Dim HostsString As String, HostsPath As String
    HostsPath = Environ("SystemRoot") & "\System32\drivers\etc\hosts"
    HostsString = ReadTextFile(HostsPath)
    
    Dim SectionToWrite As String, SelectedIP As String
    Dim EngineTribeHostsSection As String
    EngineTribeHostsSection = Trim(Split(Split(HostsString, ET_START)(1), ET_END)(0))
    
    If OptionMasterServer(0).Value Then
        SectionToWrite = vbCrLf
    ElseIf OptionMasterServer(1).Value Then
        SectionToWrite = vbCrLf & MASTER_IP & " " & MASTER_HOST & vbCrLf
    Else
        If List1.Text = "" Then
            MsgBox "你还没选择自定义接入点。"
            Exit Sub
        Else
            SelectedIP = Split(List1.Text, " (")(0)
            SectionToWrite = vbCrLf & SelectedIP & " " & MASTER_HOST & vbCrLf & SelectedIP & " " & STORAGE_HOST & vbCrLf
        End If
    End If
    
    Open HostsPath For Output As #3
    Print #3, Replace(HostsString, EngineTribeHostsSection, SectionToWrite);
    Close #3
    
    MsgBox "保存完成！", vbInformation
    Exit Sub
    
SaveHostsErrorHandler:
    MsgBox "请以管理员权限运行本程序!" & vbCrLf & "如果杀毒软件拦截了本程序修改 Hosts，请允许之!", vbCritical
    End
End Sub

Private Function ReadTextFile(sFilePath As String) As String
    On Error Resume Next
    Dim handle As Integer
    If LenB(Dir$(sFilePath)) > 0 Then
        handle = FreeFile
        Open sFilePath For Binary As #handle
        ReadTextFile = Space$(LOF(handle))
        Get #handle, , ReadTextFile
        Close #handle
    End If
End Function

Private Sub List1_Click()
    OptionMasterServer(2).Value = True
End Sub
