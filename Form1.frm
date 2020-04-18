VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "搞飞机1.0.3bate"
   ClientHeight    =   9855
   ClientLeft      =   7650
   ClientTop       =   1710
   ClientWidth     =   7230
   ForeColor       =   &H0000FFFF&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":72FA
   ScaleHeight     =   9855
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   0
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Index           =   9
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   8
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   7
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   6
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   5
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   4
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   3
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   2
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   0
      Top             =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "基地血量"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   9600
      Width           =   7335
   End
   Begin VB.Label Lablv 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3120
      Picture         =   "Form1.frx":65BD1
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   4
      Left            =   1320
      Picture         =   "Form1.frx":6E19A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   3
      Left            =   5880
      Picture         =   "Form1.frx":6E7E3
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   2
      Left            =   3960
      Picture         =   "Form1.frx":6EE2C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   1
      Left            =   2880
      Picture         =   "Form1.frx":6F475
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   0
      Left            =   1080
      Picture         =   "Form1.frx":6FABE
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   9
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   8
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   7
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   6
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   5
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   4
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   3
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   2
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   1
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape shape1 
      Height          =   255
      Index           =   0
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer '子弹编码
Dim a(0 To 9) As Boolean '子弹可用性
Dim lv As Integer '等级
Dim kill As Integer '击杀量
Dim i As Integer
Dim k As Integer
Private Sub form_load()
n = 0                '初始化
kill = 0
lv = 1
Lablv.Caption = "Lv." & lv '等级标签属性
Lablv.FontSize = 20
Lablv.AutoSize = True
Lablv.BackStyle = 0
Lablv.ForeColor = vbYellow

Label3.BackColor = vbRed '血量标签属性
Label3.Left = 0
Label3.Width = Me.Width
Label3.Height = Me.Height / 15
Label3.Top = Me.Height - Label3.Height

shape1(n).BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255) '等级一的子弹颜色
For i = 0 To 9 '初始化子弹属性可用性敌军属性
    shape1(i).Height = 155
    shape1(i).Width = 275
    shape1(i).BackStyle = 1
    shape1(i).BorderStyle = 0
    shape1(i).Visible = False
    Timer1(i).Enabled = False
    Timer1(i).Interval = 10
    a(i) = True
    Image2(Int(i / 2)).Height = 500
    Image2(Int(i / 2)).Width = 500
    Image2(Int(i / 2)).Picture = LoadPicture("1.jpg")
    Image2(Int(i / 2)).Stretch = True
Next i
End Sub
Private Sub form_KeyPress(KeyAscii As Integer)
If a(n) = True Then
    Select Case KeyAscii          '空格键发射触发
    Case 32
        shape1(n).Left = Image1.Left + 1 / 2 * (Image1.Width) - 200
        shape1(n).Top = Image1.Top
        shape1(n).Visible = True
        a(n) = False
        Timer1(n).Enabled = True
        n = n + 1
        If n = lv Then n = 0
    End Select
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode                  '方向键移动触发
Case 65
    Image1.Left = Image1.Left - 200
Case 68
    Image1.Left = Image1.Left + 200
Case 87
    Image1.Top = Image1.Top - 200
Case 83
    Image1.Top = Image1.Top + 200
End Select
                                                                
If Image1.Left < 0 Then Image1.Left = 0 '限制自由移动范围
If Image1.Top < 0 Then Image1.Top = 0
If Image1.Left > Me.Width - Image1.Width Then Image1.Left = Me.Width - Image1.Width
If Image1.Top > Me.Height - Image1.Height - 600 Then Image1.Top = Me.Height - Image1.Height - 600
End Sub

Private Sub Timer1_Timer(n As Integer)
Timer1(n).Enabled = True
shape1(n).Top = shape1(n).Top - 50
For i = 0 To 4                                               '检测敌军状态
    If shape1(n).Top < 0 Then                                'if情况一：子弹打空判定
        Timer1(n).Enabled = False
        shape1(n).Visible = False
        a(n) = True
    End If
    If ((shape1(n).Top <= Image2(i).Top And shape1(n).Left >= Image2(i).Left _
        And shape1(n).Left <= Image2(i).Left + Image2(i).Width)) Then 'if情况二：子弹击中判定
            kill = kill + 1 '击杀数加一
                If kill Mod 5 = 0 Then Call lvup '击杀5个升一级
            shape1(n).BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
            Timer1(n).Enabled = False
            shape1(n).Visible = False
            a(n) = True
            Image2(i).Top = -Image2(i).Height '敌军返回屏幕顶端随机位置
            Image2(i).Left = Int(Rnd * (Me.Width - Image2(i).Width - 200))
    End If
Next i
End Sub

Private Sub Timer2_Timer()
For k = 0 To 4      '所有方块下降并循环
    Image2(k).Top = Image2(k).Top + 50
    If Image2(k).Top >= Me.Height Then
            Label3.Width = Label3.Width - 500
            Label3.Left = Label3.Left + 225
            Image2(k).Top = -Image2(k).Height '敌军返回屏幕顶端随机位置
            Image2(k).Left = Int(Rnd * (Me.Width - Image2(k).Width - 200))
    End If
Next k
If Label3.Width <= 400 Then
    Unload Me
    If kill >= 0 And kill < 20 Then
        MsgBox "击杀数：" & kill & vbCrLf & "真的垃圾啊！", , "评价"
    ElseIf kill >= 20 And kill < 50 Then
        MsgBox "击杀数：" & kill & vbCrLf & "垃圾啊！", , "评价"
    ElseIf kill >= 50 And kill <= 80 Then
        MsgBox "击杀数：" & kill & vbCrLf & "可以的！", , "评价"
    Else
        MsgBox "击杀数：" & kill & vbCrLf & "你就是我的游戏测试员了！", , "评价"
    End If
End If
End Sub

Public Sub lvup()
    lv = lv + 1
    Timer2.Interval = 60
    Timer2.Interval = Timer2.Interval - lv * 4
    If lv = 10 Then lv = 9
    Lablv.Caption = "Lv." & lv
    Label3.Left = Label3.Left - 225
    Label3.Width = Label3.Width + 500
    If Label3.Width >= 7335 Then Label3.Width = 7335: Label3.Left = 0
    Lablv.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

