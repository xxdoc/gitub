VERSION 5.00
Begin VB.Form Mainform 
   Caption         =   "Gitub代码格式化工具"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "开始转换"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1260
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Text            =   "C:\vb6db"
      Top             =   795
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   435
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "支持拖到文件夹"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label FileInCopy 
      Height          =   180
      Left            =   1260
      TabIndex        =   6
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "正在备份："
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "备份位置："
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "源代码位置："
      Height          =   180
      Left            =   300
      TabIndex        =   1
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CDRomRoot As String
Dim Path
Dim fs As New FileSystemObject
Dim Source As String, Target As String

Private Sub Command1_Click()

    Dim SourceFd As Folder, TargetFd As Folder

    Dim f        As File, I As Integer

   
    
    If fs.FolderExists(Source) And Source <> "" And Target <> "" Then
        Command1.Enabled = False
        fs.CopyFolder Source, Target, True
        fs.DeleteFolder Source
        MyMkDir Source
    Else
        MsgBox "只支持文件夹批量转换"
        Exit Sub

    End If

    Set SourceFd = fs.GetFolder(Target & "\")

    For Each f In SourceFd.Files

        ShowFileMsg Source & "\" & f.Name
        DoEvents

        If ExtName(f.Path) = "frm" Or ExtName(f.Path) = "dsr" Or ExtName(f.Path) = "dsn" Or ExtName(f.Path) = "ini" Or ExtName(f.Path) = "bas" Or ExtName(f.Path) = "cls" Or ExtName(f.Path) = "pag" Or ExtName(f.Path) = "ctl" Or ExtName(f.Path) = "pag" Or ExtName(f.Path) = "vbp" Then
            MyCopy Target & "\", Source & "\", f.Name
        Else

            On Error Resume Next

            fs.CopyFile f.Path, Source & "\" & f.Name, True

            On Error GoTo 0

        End If

    Next
    Set SourceFd = Nothing
    Set TargetFd = fs.GetFolder(Source & "\")

    For Each f In TargetFd.Files

        f.Attributes = f.Attributes And &HFFFFFFFE
    Next
    Set TargetFd = Nothing
    FileInCopy.Caption = ""
    DoEvents
    
    Command1.Enabled = True
    MsgBox "已安装完毕, 单击「确定」按钮结束!"
    Unload Me

End Sub

Public Sub DropFiles(ByVal hDrop&)

    Dim sFileName$, nCharsCopied&

    sFileName = String$(MAX_PATH, vbNullChar)
    nCharsCopied = DragQueryFile(hDrop, 0&, sFileName, MAX_PATH)
    DragFinish hDrop

    If nCharsCopied Then
        nCharsCopied = InStr(sFileName, Chr(0)) - 1
        sFileName = Left$(sFileName, nCharsCopied)
        Source = sFileName '获取文件路径

       ' If Right(Source, 1) <> "\" Then Source = Source & "\"
        Text1.Text = Source
        Target = Source & "-Back"
    
       ' If Right(Target, 1) <> "\" Then Target = Target & "\"
        Text2.Text = Target

    End If

End Sub

Private Sub Form_Load()

    Dim drv As Drive

    SetIcon Me.hWnd, "AAA"
    DragAcceptFiles Text1.hWnd, 1&
    procOld = SetWindowLong(Text1.hWnd, GWL_WNDPROC, AddressOf WindowProc)
    CDRomRoot = App.Path ' "D:\F8317"
    'For Each drv In fs.Drives
    '    If drv.DriveType = CDRom Then
    '        CDRomRoot = drv.DriveLetter & ":\F8317"
    '    End If
    'Next
    'If Len(CDRomRoot) = 0 Then
    '    MsgBox "找不到光驱, 无法安装!", vbCritical, "VB6 数据库程序设计, 安装程序"
    '    End
    'End If
    Source = CDRomRoot

   ' If Right(Source, 1) <> "\" Then Source = Source & "\"
    Text1.Text = Source
    Target = Source & "-Back" '
    Text2 = Target
    '    If Right(Target, 1) <> "\" Then Target = Target & "\"
    '       Text2.Text = Target
    ' Path = Array("ch01", "ch02", "ch03", "ch04", "ch05", "ch06", "ch07", "ch08", "ch09", "ch10", "ch11", "ch12", "ch13", "ch14", "ch16", "120,000", "mdb", "dbf", "dsn", "excel", "F202", "txt", "txtimp")

End Sub

Function ExtName(File As String)
    Dim pos As Integer
    pos = InStrRev(File, ".")
    If pos <> 0 Then
        ExtName = LCase(Mid(File, pos + 1))
    End If
End Function

Function MyMkDir(ByVal Path As String) As Boolean
    Dim pos As Integer
    If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
    If fs.FolderExists(Path) Then
        MyMkDir = True
        Exit Function
    Else
        pos = InStrRev(Path, "\")
        If pos > 3 Then
            MyMkDir = MyMkDir(Left(Path, pos - 1))
            fs.CreateFolder Path
        Else
            fs.CreateFolder Path
            Exit Function
        End If
    End If
End Function

Sub MyCopy(ByVal SourceDir As String, ByVal TargetDir As String, ByVal Name As String)

    Dim fin  As TextStream, fout As TextStream

    Dim temp As String

    Set fin = fs.OpenTextFile(SourceDir & Name)
    temp = fin.ReadAll
    Set fout = fs.CreateTextFile(TargetDir & Name, True)
    '    Open TargetDir & Name For Output As #1
    '    Print #1, fin.ReadAll
    '    Close
    fout.Write Replace(temp, Chr(10), Chr(13) & Chr(10))
    fin.Close

    fout.Close

End Sub

Sub ShowFileMsg(ByVal File As String)
    If Len(File) > 30 Then
        File = Left(File, 12) & " .... " & Right(File, 12)
    End If
    FileInCopy.Caption = File
End Sub

