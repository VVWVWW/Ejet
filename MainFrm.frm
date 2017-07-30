VERSION 5.00
Object = "{AAC8DFAF-8A34-11D3-B327-000021C5C8A9}#1.0#0"; "SYSTRAY.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Ejet"
   ClientHeight    =   7035
   ClientLeft      =   3195
   ClientTop       =   420
   ClientWidth     =   3210
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Tmr_Msg 
      Interval        =   1000
      Left            =   2400
      Top             =   4080
   End
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1920
      Top             =   4080
   End
   Begin SysTrayCtl.cSysTray cSysTray 
      Left            =   120
      Top             =   4080
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "MainFrm.frx":0ECA
      TrayTip         =   "Ejet"
   End
   Begin VB.Frame FrmTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
      Begin VB.ListBox Lit_Ply 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         IntegralHeight  =   0   'False
         ItemData        =   "MainFrm.frx":1DA4
         Left            =   120
         List            =   "MainFrm.frx":1DA6
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Lbl_LitCtrl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Lbl_LitCtrl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Lbl_LitCtrl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "一"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Lbl_LitCtrl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "保存"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame FrmTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2775
      Begin VB.DriveListBox Drive 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.FileListBox Flb 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   750
         Hidden          =   -1  'True
         Left            =   1440
         MultiSelect     =   2  'Extended
         Pattern         =   "*.m3u"
         System          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.DirListBox Dlb 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   720
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Lbl_ReplaceList 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "替换"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Lbl_Visible 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "列表"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Lbl_AddToList 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label LblCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Lbl_MscVol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Lbl_MscVol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin WMPLibCtl.WindowsMediaPlayer KERNEL 
      Height          =   480
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2011
      _cy             =   847
   End
   Begin VB.Image Ige_Ico 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LblTag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "播放列表"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   840
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Lbl_End 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LblCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "|>"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.Label LblCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Label LblCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "·)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.Label LblCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.Label LblCtrl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "三"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Lbl_Msg 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "MessageBox"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Lbl_Cpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Ejet"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim IsMoving As Boolean, IsMscLitChange As Boolean, ISSTOPPED As Boolean
Dim OldSize As Integer, MscVolume As Integer
Dim MyX As Long, MyY As Long, MyLft As Long, MyTop As Long
Dim MusicList() As String, MscPlyMod(2) As String
Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "该软件已运行！"
        End
    End If
    BoardSet 375
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    App.TaskVisible = True
    IsMoving = False
    IsMscLitChange = True
    ISSTOPPED = True
    MscVolume = 100
    MscPlyMod(0) = "R"
    MscPlyMod(1) = "L"
    MscPlyMod(2) = "O"
    ReDim MusicList(0)
    Ige_Ico.Picture = Me.Icon
    Initialization Command()
    LoadTheme
End Sub
Private Sub Form_Resize()
    If Me.WindowState = 1 Then Me.Hide
    MyTop = Me.Top
    MyLft = Me.Left
End Sub
Private Sub cSysTray_MouseUp(Button As Integer, Id As Long)
    Me.WindowState = 0
    Me.Show
    Me.Top = MyTop
    Me.Left = MyLft
    Me.SetFocus
End Sub
Private Sub Ige_Ico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        IsMoving = True
        MyX = X
        MyY = Y
    End If
End Sub
Private Sub Ige_Ico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsMoving Then Exit Sub
    Me.Left = Me.Left - MyX + X
    Me.Top = Me.Top - MyY + Y
    If Me.Top < 0 Then MyY = Y
    If Me.Left < 0 Then MyX = X
    If Me.Top + Me.Height > Screen.Height Then MyY = Y
    If Me.Width + Me.Left > Screen.Width Then MyX = X
    MyTop = Me.Top
    MyLft = Me.Left
End Sub
Private Sub Ige_Ico_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsMoving = False
    DrawLine
End Sub
Private Sub Lbl_Cpt_Change()
    If Len(Lbl_Cpt.Caption) > 25 Then Lbl_Cpt.Caption = Left(Lbl_Cpt.Caption, 25)
End Sub
Private Sub Lbl_Cpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then IsMoving = True
End Sub
Private Sub Lbl_Cpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or X > Lbl_Cpt.Width Or Not IsMoving Then Exit Sub
    If GetPlayingMsc() <> 0 And UBound(MusicList) <> 0 Then Music "TimeSet", X / Lbl_Cpt.Width * Music("TimeGet", "0")
End Sub
Private Sub Lbl_Cpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsMoving = False
End Sub
Private Sub Lbl_End_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Msg "单击此对话框以退出"
    Else
        Me.WindowState = 1
    End If
End Sub
Private Sub LblCtrl_Click(Index As Integer)
    Dim i As Integer
    Dim FilePath As String 'On Error Resume Next
    Select Case Index
    Case 0
        If UBound(MusicList) <> 0 Then
            ISSTOPPED = False
            Select Case GetMscPlyMod()
            Case 0
                FilePath = MusicList(Rand(1, UBound(MusicList)))
            Case 1
                If GetPlayingMsc() <> 1 Then
                    FilePath = MusicList(GetPlayingMsc() - 1)
                Else
                    FilePath = MusicList(UBound(MusicList))
                End If
            Case 2
                If GetPlayingMsc() <> 1 Then
                    FilePath = MusicList(GetPlayingMsc() - 1)
                Else
                    FilePath = MusicList(UBound(MusicList))
                End If
            End Select
            Music "play", FilePath
            If LblCtrl(1).Caption = "|>" Then LblCtrl_Click 1
        Else
            Music "stop"
            Msg "错误:当前播放列表内无歌曲"
        End If
    Case 1
        If LblCtrl(Index).Caption = "||" Then
            LblCtrl(Index).Caption = "|>"
            Music "pause"
        Else
            LblCtrl(Index).Caption = "||"
            Music "continue"
        End If
        If Music("URL") = "" And UBound(MusicList) <> 0 Then LblCtrl_Click 2
    Case 2
        If UBound(MusicList) <> 0 Then
            ISSTOPPED = False
            Select Case GetMscPlyMod()
            Case 0
                FilePath = MusicList(Rand(1, UBound(MusicList)))
            Case 1
                If GetPlayingMsc() <> UBound(MusicList) Then
                    FilePath = MusicList(GetPlayingMsc() + 1)
                Else
                    FilePath = MusicList(1)
                End If
            Case 2
                If GetPlayingMsc() <> UBound(MusicList) Then
                    FilePath = MusicList(GetPlayingMsc() + 1)
                Else
                    FilePath = MusicList(1)
                End If
            End Select
            Music "play", FilePath
            If LblCtrl(1).Caption = "|>" Then LblCtrl_Click 1
        Else
            Music "stop"
            Msg "错误:当前播放列表内无歌曲"
        End If
    Case 3
        Lbl_MscVol(0).Visible = Not Lbl_MscVol(0).Visible
        Lbl_MscVol(1).Visible = Lbl_MscVol(0).Visible
        LblCtrl(4).Visible = Not Lbl_MscVol(0).Visible
        LblCtrl(5).Visible = LblCtrl(4).Visible
    Case 4
        If GetMscPlyMod() <> UBound(MscPlyMod) Then
            LblCtrl(Index).Caption = MscPlyMod(GetMscPlyMod() + 1)
        Else
            LblCtrl(Index).Caption = MscPlyMod(0)
        End If
    Case 5
        LblTag.Visible = Not LblTag.Visible
        If LblTag.Visible Then
            Me.Height = Me.Height * 7
        Else
            Me.Height = Me.Height / 7
        End If
        If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
    End Select
    DrawLine
End Sub
Private Sub LblCtrl_DblClick(Index As Integer)
    Select Case Index
    Case 3
        If LblCtrl(Index).Caption = "·)" Then
            Music "volume", 0
            LblCtrl(Index).Caption = "·"
        Else
            Music "volume", Val(MscVolume)
            LblCtrl(Index).Caption = "·)"
        End If
    End Select
    LblCtrl_Click Index
End Sub
Private Sub Lbl_MscVol_Click(Index As Integer)
    Select Case Index
    Case 0
        MscVolume = MscVolume - 5
        If MscVolume < 0 Then MscVolume = 0
    Case 1
        MscVolume = MscVolume + 5
        If MscVolume > 100 Then MscVolume = 100
    End Select
    If LblCtrl(3).Caption = "·)" Then Music "volume", Val(MscVolume)
    Msg "音量：" & MscVolume & "%"
End Sub
Private Sub Lbl_MscVol_DblClick(Index As Integer)
    Lbl_MscVol_Click Index
End Sub
Private Sub Lbl_Msg_Click()
    Lbl_Msg.Visible = False
    Select Case Lbl_Msg.Caption
    Case "单击此对话框以退出"
        End
    End Select
End Sub
Private Sub LblTag_Click()
    Dim i As Integer
    If LblTag.Caption = "歌曲添加" Then
        If IsMscLitChange Then
            Dim FilePath As String
            Lit_Ply.Clear
            For i = 1 To UBound(MusicList)
                FilePath = MusicList(i)
                Lit_Ply.AddItem PathToName(FilePath)
            Next
            IsMscLitChange = False
        End If
        LblTag.Caption = "播放列表"
        FrmTag(0).ZOrder 0
    Else
        LblTag.Caption = "歌曲添加"
        FrmTag(1).ZOrder 0
    End If
End Sub
Private Sub Lit_Ply_DblClick()
    Dim i As Integer
    Dim FilePath As String
    For i = 0 To (UBound(MusicList) - 1)
        If Lit_Ply.Selected(i) Then Exit For
    Next
    FilePath = MusicList(i + 1)
    Music "play", FilePath
    LblCtrl(1).Caption = "||"
End Sub
Private Sub Lbl_LitCtrl_Click(Index As Integer)
    Dim i As Integer, FilePath As String, IsSelected As Boolean
    If Lit_Ply.ListCount = 0 Then Exit Sub
    IsSelected = False
    For i = 1 To UBound(MusicList)
        If Lit_Ply.Selected(i - 1) = True Then
            IsSelected = True
            Exit For
        End If
    Next
    Select Case Index
    Case 0
        Dim j As Integer
        For i = 1 To Lit_Ply.ListCount
            If Lit_Ply.Selected(i - 1) Then
                MusicListDel i
                Exit For
            End If
        Next
        IsMscLitChange = True
        LblTag.Caption = "歌曲添加"
        LblTag_Click
        If UBound(MusicList) <> 0 Then
            If i > 1 Then Lit_Ply.Selected(i - 2) = True
            If i = 1 Then Lit_Ply.Selected(0) = True
        End If
    Case 1
        If Not IsSelected Then Exit Sub
        For i = 0 To Lit_Ply.ListCount
            If Lit_Ply.Selected(i) Then Exit For
        Next
        If i = 0 Then Exit Sub
        Exchange MusicList(i), MusicList(i + 1)
        IsMscLitChange = True
        LblTag.Caption = "歌曲添加"
        LblTag_Click
        Lit_Ply.Selected(i - 1) = True
    Case 2
        If Not IsSelected Then Exit Sub
        For i = 0 To Lit_Ply.ListCount
            If Lit_Ply.Selected(i) Then Exit For
        Next
        If i = Lit_Ply.ListCount - 1 Then Exit Sub
        Exchange MusicList(i + 1), MusicList(i + 2)
        IsMscLitChange = True
        LblTag.Caption = "歌曲添加"
        LblTag_Click
        Lit_Ply.Selected(i + 1) = True
    Case 3
        On Error GoTo ErrorWrongPath
        SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
        FilePath = InputBox$("请输入目标文件全路径", , Dlb.Path & "\播放列表.m3u", Me.Left, Me.Top)
        SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
        If FilePath = "" Then Exit Sub
        PathStore FilePath
        Msg "列表 -> 目标路径"
    End Select
    Exit Sub
ErrorWrongPath:
    Msg "路径错误，创建列表失败"
End Sub
Private Sub Drive_Change()
    On Error GoTo ErrDisposing
    Dlb.Path = Drive.Drive
    Exit Sub
ErrDisposing:
End Sub
Private Sub Dlb_Change()
    Flb.Path = Dlb.Path
End Sub
Private Sub Flb_DblClick()
    Lbl_AddToList_Click
End Sub
Private Sub Lbl_Visible_Click()
    If Flb.Pattern = "*.m3u" Then
        Lbl_Visible.Caption = "歌曲"
        Flb.Pattern = "*.mp3"
    Else
        Lbl_Visible.Caption = "列表"
        Flb.Pattern = "*.m3u"
    End If
End Sub
Private Sub Lbl_Visible_DblClick()
    Lbl_Visible_Click
End Sub
Private Sub Lbl_ReplaceList_Click()
    If LblCtrl(1).Caption = "||" Then LblCtrl_Click 1
    MusicListDel -1
    Lbl_AddToList_Click
    LblCtrl_Click 2
End Sub
Private Sub Lbl_AddToList_Click()
    Dim i As Integer, OldMusicNum As Integer
    OldMusicNum = UBound(MusicList)
    If Flb.Pattern = "*.mp3" Then
        For i = 0 To (Flb.ListCount - 1)
            If Flb.Selected(i) Then MusicListAdd Dlb.Path & "\" & Flb.List(i)
        Next
    Else
        For i = 0 To (Flb.ListCount - 1)
            If Flb.Selected(i) Then
                AddM3UFile Dlb.Path & "\" & Flb.List(i)
            End If
        Next
    End If
    ListChk
    Msg (UBound(MusicList) - OldMusicNum) & "歌曲 -> 播放队列"
End Sub
Private Sub Tmr_Timer()
    If Music("TimeGet", "1") < Music("TimeGet") - 0.6 Then
        DrawLine
        Me.Line (0, Ige_Ico.Height * 0.9)-(Music("TimeGet", "1") / Music("TimeGet") * Me.Width, Ige_Ico.Height * 0.9), RGB(0, 128, 196)
    Else
        If GetMscPlyMod() <> 2 Then
            LblCtrl_Click 2
        Else
            Music "play", Music("URL")
        End If
    End If
End Sub
Private Sub Tmr_Msg_Timer()
    Lbl_Msg.Visible = False
    DrawLine
    Tmr_Msg.Enabled = False
End Sub
Function DrawLine()
    Me.Cls
    On Error Resume Next
    Me.PaintPicture Me.Picture, 0, 0, Me.Width, FrmTag(0).Top
    Me.Line (0, 0)-(Me.Width, 0), RGB(0, 128, 128)
    Me.Line (0, 0)-(0, Me.Height), RGB(0, 128, 128)
    Me.Line (0, Me.Height * 0.98)-(Me.Width, Me.Height * 0.98), RGB(0, 128, 128)
    Me.Line (Me.Width - 19, 0)-(Me.Width - 19, Me.Height), RGB(0, 128, 128)
End Function
Function BoardSet(SQR As Integer)
    Dim NewHei, NewWid, FontSize As Long
    Dim i As Integer
    With LblTag
        NewHei = SQR * 1
        NewWid = SQR * 2.5
        .AutoSize = True
        .FontSize = 2
        Do While .Width < NewWid And .Height < NewHei
            .FontSize = .FontSize + 1
        Loop
        .FontSize = .FontSize - 1
        FontSize = .FontSize
        .AutoSize = False
    End With
    Me.Height = SQR * 2
    Me.Width = SQR * 10
    With Ige_Ico
        .Height = SQR * 1
        .Width = SQR * 1
        .Top = 0
        .Left = 0
    End With
    With Lbl_Cpt
        .Height = SQR * 1
        .Width = SQR * 8
        .Top = 0
        .Left = SQR * 1
        .FontSize = FontSize
    End With
    With Lbl_Msg
        .Height = SQR * 1
        .Width = SQR * 8
        .Top = 0
        .Left = SQR * 1
        .FontSize = FontSize
    End With
    With Lbl_End
        .Height = SQR * 1
        .Width = SQR * 1
        .Top = 0
        .Left = SQR * 9
        .FontSize = FontSize
    End With
    For i = 0 To (LblCtrl.Count - 1)
        With LblCtrl(i)
            .Height = SQR * 1
            .Width = SQR * 1
            .Top = SQR * 1
            .Left = SQR * i
            If i > 2 Then .Left = .Left + SQR * 4
            .FontSize = FontSize
        End With
    Next
    For i = 0 To 1
        With Lbl_MscVol(i)
            .Height = SQR * 1
            .Width = SQR * 1
            .Top = SQR * 1
            .Left = SQR * (8 + i)
            .FontSize = FontSize
        End With
    Next
    For i = 0 To (FrmTag.Count - 1)
        With FrmTag(i)
            .Height = SQR * 12
            .Width = SQR * 10
            .Top = SQR * 2
            .Left = 0
        End With
    Next
    With LblTag
        .Height = SQR * 1
        .Width = SQR * 4
        .Top = SQR * 1
        .Left = SQR * 3
        .FontSize = FontSize
    End With
    With Lit_Ply
        .Height = SQR * 11
        .Width = SQR * 10
        .Top = 0
        .Left = 0
        .FontSize = FontSize
    End With
    For i = 0 To (Lbl_LitCtrl.Count - 1)
        With Lbl_LitCtrl(i)
            .Height = SQR * 1
            .Width = SQR * (10 / 4)
            .Top = SQR * 11
            .Left = SQR * (10 / 4) * i
            .FontSize = FontSize
        End With
    Next
    With Drive
        .Width = SQR * 5
        .Top = 0
        .Left = 0
        .FontSize = FontSize
    End With
    With Lbl_Visible
        .Height = SQR * 1
        .Width = SQR * 2
        .Top = 0
        .Left = SQR * 5
        .FontSize = FontSize
    End With
    With Lbl_ReplaceList
        .Height = SQR * 1
        .Width = SQR * 2
        .Top = 0
        .Left = SQR * 7
        .FontSize = FontSize
    End With
    With Lbl_AddToList
        .Height = SQR * 1
        .Width = SQR * 1
        .Top = 0
        .Left = SQR * 9
        .FontSize = FontSize
    End With
    With Dlb
        .Height = SQR * 11
        .Width = SQR * 5
        .Top = SQR * 1
        .Left = 0
        .FontSize = FontSize
    End With
    With Flb
        .Height = SQR * 11
        .Width = SQR * 5
        .Top = SQR * 1
        .Left = SQR * 5
        .FontSize = FontSize
    End With
End Function
Function MusicListAdd(MusicPath As String)
    Dim IsAdd As Boolean, i As Integer
    IsAdd = False
    For i = 0 To UBound(MusicList)
        If MusicList(i) = MusicPath Then IsAdd = True
    Next
    If Not IsAdd Then
        ReDim Preserve MusicList(UBound(MusicList) + 1)
        MusicList(UBound(MusicList)) = MusicPath
        IsMscLitChange = True
    End If
End Function
Function MusicListDel(Index As Integer)
    Dim i As Integer
    IsMscLitChange = True
    If Index = -1 Then
        For i = 1 To UBound(MusicList)
            MusicListDel UBound(MusicList)
        Next
        Exit Function
    End If
    For i = Index To (UBound(MusicList) - 1)
        MusicList(i) = MusicList(i + 1)
    Next
    ReDim Preserve MusicList(UBound(MusicList) - 1)
End Function
Function AddM3UFile(M3UFile As String)
    Dim PathLine As String, i As Integer
    Open M3UFile For Input As #1
    Do Until EOF(1)
        Line Input #1, PathLine
        If PathLine <> "" And Mid(PathLine, 1, 1) <> "#" Then MusicListAdd PathLine
    Loop
    Close #1
End Function
Function Rand(MinNum As Integer, MaxNum As Integer) As Integer
    Randomize
    MaxNum = MaxNum + 1
    Rand = MinNum + Int(Rnd * (MaxNum - MinNum))
End Function
Function PathToName(FullPath As String) As String
    PathToName = Left(Dir(FullPath), InStrRev(Dir(FullPath), ".") - 1)
End Function
Function GetMscPlyMod() As Integer
    Dim i As Byte
    For i = 0 To UBound(MscPlyMod)
        If LblCtrl(4).Caption = MscPlyMod(i) Then Exit For
    Next
    GetMscPlyMod = i
End Function
Function GetPlayingMsc() As Integer
    Dim i As Integer
    For i = 1 To UBound(MusicList)
        If MusicList(i) = Music("URL") Then Exit For
        If i = UBound(MusicList) Then
            GetPlayingMsc = 0
            Exit Function
        End If
    Next
    GetPlayingMsc = i
End Function
Function IsListDel() As Boolean
    PathStore App.Path & "\Log.Ejet.m3u"
    Dim OldNum As Integer
    OldNum = UBound(MusicList)
    ListChk
    If OldNum = UBound(MusicList) Then
        IsListDel = False
    Else
        IsListDel = True
    End If
End Function
Function PathStore(FilePath As String)
    Dim i As Integer
    Open FilePath For Output As #1
        For i = 1 To UBound(MusicList)
            Print #1, MusicList(i)
        Next
    Close #1
End Function
Function ListChk()
    Dim i As Integer
Restart:
    If UBound(MusicList) = 0 Then GoTo Tail
    For i = 1 To UBound(MusicList)
        If Dir(MusicList(i)) = "" Then
            MusicListDel i
            IsMscLitChange = True
            GoTo Restart
        End If
    Next
Tail:
    LblTag.Caption = "歌曲添加"
    LblTag_Click
End Function
Function Exchange(Arg1 As String, Arg2 As String)
    Dim ExchangeSpace As String
    ExchangeSpace = Arg1
    Arg1 = Arg2
    Arg2 = ExchangeSpace
End Function
Function Msg(Message As String)
    Lbl_Msg.Caption = Message
    Lbl_Msg.ToolTipText = Message
    Lbl_Msg.Visible = True
    Tmr_Msg.Enabled = False
    Tmr_Msg.Enabled = True
End Function
Function Initialization(Cmd As String)
    Dim Cmds() As String
    Cmds = Split(Cmd, ",")
    If UBound(Cmds) = -1 Then Exit Function
    If Left(Cmds(0), 1) = """" And Right(Cmds(0), 1) = """" Then Cmds(0) = Replace(Cmds(0), """", "")
    If Right(Cmds(0), Len(Cmds(0)) - InStrRev(Cmds(0), ".")) = "mp3" Then
        MusicListAdd Cmds(0)
    Else
        AddM3UFile Cmds(0)
    End If
    LblCtrl_Click 2
End Function
Function LoadTheme()
    On Error Resume Next
    Dim PathLine(5), Colors() As String, RGBs(5, 2), i As Byte
    If Dir(App.Path & "\Theme\Color.Ejet") <> "" Then
        Open App.Path & "\Theme\Color.Ejet" For Input As #1
        For i = 0 To 5
            Line Input #1, PathLine(i)
            Colors = Split(PathLine(i), " ")
            RGBs(i, 0) = Colors(0)
            RGBs(i, 1) = Colors(1)
            RGBs(i, 2) = Colors(2)
        Next
        Close #1
        Me.BackColor = RGB(RGBs(0, 0), RGBs(0, 1), RGBs(0, 2))
        Lbl_Msg.BackColor = RGB(RGBs(0, 0), RGBs(0, 1), RGBs(0, 2))
        Lbl_Cpt.ForeColor = RGB(RGBs(1, 0), RGBs(1, 1), RGBs(1, 2))
        Lbl_End.ForeColor = RGB(RGBs(1, 0), RGBs(1, 1), RGBs(1, 2))
        Lbl_Msg.ForeColor = RGB(RGBs(1, 0), RGBs(1, 1), RGBs(1, 2))
        For i = 0 To LblCtrl.Count
            LblCtrl(i).ForeColor = RGB(RGBs(1, 0), RGBs(1, 1), RGBs(1, 2))
        Next
        LblTag.BackColor = RGB(RGBs(2, 0), RGBs(2, 1), RGBs(2, 2))
        FrmTag(0).BackColor = RGB(RGBs(2, 0), RGBs(2, 1), RGBs(2, 2))
        Lit_Ply.BackColor = RGB(RGBs(2, 0), RGBs(2, 1), RGBs(2, 2))
        LblTag.ForeColor = RGB(RGBs(3, 0), RGBs(3, 1), RGBs(3, 2))
        Lit_Ply.ForeColor = RGB(RGBs(3, 0), RGBs(3, 1), RGBs(3, 2))
        For i = 0 To Lbl_LitCtrl.Count - 1
            Lbl_LitCtrl(i).ForeColor = RGB(RGBs(3, 0), RGBs(3, 1), RGBs(3, 2))
        Next
        FrmTag(1).BackColor = RGB(RGBs(4, 0), RGBs(4, 1), RGBs(4, 2))
        Drive.BackColor = RGB(RGBs(4, 0), RGBs(4, 1), RGBs(4, 2))
        Dlb.BackColor = RGB(RGBs(4, 0), RGBs(4, 1), RGBs(4, 2))
        Flb.BackColor = RGB(RGBs(4, 0), RGBs(4, 1), RGBs(4, 2))
        Drive.ForeColor = RGB(RGBs(5, 0), RGBs(5, 1), RGBs(5, 2))
        Lbl_Visible.ForeColor = RGB(RGBs(5, 0), RGBs(5, 1), RGBs(5, 2))
        Lbl_ReplaceList.ForeColor = RGB(RGBs(5, 0), RGBs(5, 1), RGBs(5, 2))
        Lbl_AddToList.ForeColor = RGB(RGBs(5, 0), RGBs(5, 1), RGBs(5, 2))
        Dlb.ForeColor = RGB(RGBs(5, 0), RGBs(5, 1), RGBs(5, 2))
        Flb.ForeColor = RGB(RGBs(5, 0), RGBs(5, 1), RGBs(5, 2))
    End If
    If Dir(App.Path & "\Theme\Skin.jpg") <> "" Then
        Me.Picture = LoadPicture(App.Path & "\Theme\Skin.jpg")
    End If
    If Dir(App.Path & "\Set\Path.Ejet") <> "" Then
        Open App.Path & "\Set\Path.Ejet" For Input As #1
            Line Input #1, PathLine(0)
        Close #1
        Dlb.Path = PathLine(0)
    End If
    If Dir(App.Path & "\Set\Size.Ejet") <> "" Then
        Open App.Path & "\Set\Size.Ejet" For Input As #1
            Line Input #1, PathLine(0)
        Close #1
        BoardSet Val(PathLine(0))
    End If
    DrawLine
End Function
Function Music(Cmd As String, Optional Arg As String = "0")
    If ISSTOPPED And Cmd <> "volume" Then
        LblCtrl(1).Caption = "|>"
        Exit Function
    End If
    Select Case Cmd
    Case "URL"
        Music = KERNEL.URL
    Case "play"
        If IsListDel() = True Then
            If UBound(MusicList) <> 0 Then
                Arg = MusicList(1)
            Else
                MusicListDel -1
                MsgBox "歌曲路径错误" & vbCrLf & "播放列表中的歌曲可能已被移除，播放列表已被保存至：" & App.Path & "\Log.Ejet.m3u"
                Music "stop"
                Exit Function
            End If
        End If
        KERNEL.URL = Arg
        KERNEL.Controls.play
        Lbl_Cpt.Caption = " " & PathToName(Arg)
        Sleep 250
        Tmr.Enabled = True
    Case "pause"
        KERNEL.Controls.pause
    Case "continue"
        KERNEL.Controls.play
    Case "stop"
        KERNEL.Controls.stop
        ISSTOPPED = True
        Tmr.Enabled = False
        LblCtrl(1).Caption = "|>"
        Lbl_Cpt.Caption = "Ejet"
        DrawLine
    Case "volume"
        KERNEL.settings.volume = Val(Arg)
    Case "TimeGet"
        Select Case Arg
        Case "0"
            Music = KERNEL.currentMedia.duration
        Case "1"
            Music = KERNEL.Controls.currentPosition
        End Select
    Case "TimeSet"
        KERNEL.Controls.currentPosition = Val(Arg)
    End Select
End Function
