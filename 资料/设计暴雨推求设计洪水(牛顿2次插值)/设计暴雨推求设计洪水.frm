VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "设计暴雨推求设计洪水"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   FillStyle       =   0  'Solid
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6480
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   4100
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2650
      Begin VB.TextBox Text7 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   14
         Top             =   3400
         Width           =   800
      End
      Begin VB.TextBox Text6 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   13
         Top             =   2900
         Width           =   800
      End
      Begin VB.TextBox Text5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   12
         Top             =   2400
         Width           =   800
      End
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   11
         Top             =   1900
         Width           =   800
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   10
         Top             =   1400
         Width           =   800
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   9
         Top             =   900
         Width           =   800
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1300
         TabIndex        =   8
         Top             =   400
         Width           =   800
      End
      Begin VB.Label Label14 
         Caption         =   "<1.0"
         Height          =   255
         Left            =   2160
         TabIndex        =   46
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "H24点="
         Height          =   255
         Left            =   400
         TabIndex        =   45
         Top             =   2980
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "mm"
         Height          =   255
         Left            =   2250
         TabIndex        =   23
         Top             =   2980
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "mm"
         Height          =   255
         Left            =   2250
         TabIndex        =   22
         Top             =   2480
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "1-8"
         Height          =   255
         Left            =   2205
         TabIndex        =   18
         Top             =   1980
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "‰"
         Height          =   255
         Left            =   2250
         TabIndex        =   17
         Top             =   1480
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "km"
         Height          =   375
         Left            =   2250
         TabIndex        =   16
         Top             =   980
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "km^2"
         Height          =   375
         Left            =   2205
         TabIndex        =   15
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label7 
         Caption         =   "地表径流系数="
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "初损I0="
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "所属暴雨区"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "流域坡降J="
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "河道长度L="
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   980
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "流域面积F="
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4050
      Left            =   2700
      TabIndex        =   0
      Top             =   80
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7144
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483643
      OLEDropMode     =   1
      TabCaption(0)   =   "推理公式法"
      TabPicture(0)   =   "设计暴雨推求设计洪水.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "经验单位线法"
      TabPicture(1)   =   "设计暴雨推求设计洪水.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text11"
      Tab(1).Control(1)=   "Text10"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(4)=   "Label22"
      Tab(1).Control(5)=   "Label21"
      Tab(1).Control(6)=   "Label20"
      Tab(1).Control(7)=   "Label19"
      Tab(1).ControlCount=   8
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   -73560
         TabIndex        =   42
         Top             =   2800
         Width           =   900
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   -73560
         TabIndex        =   41
         Top             =   2330
         Width           =   900
      End
      Begin VB.Frame Frame3 
         Caption         =   "请选择其中的一条经验单位线"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   32
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton Option8 
            BackColor       =   &H80000009&
            Caption         =   "500km^2＜F≤1000km^2   "
            Height          =   255
            Left            =   100
            TabIndex        =   38
            Top             =   1440
            Width           =   3255
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H80000009&
            Caption         =   "300km^2＜F≤500km^2   "
            Height          =   255
            Left            =   100
            TabIndex        =   37
            Top             =   1200
            Width           =   3255
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000009&
            Caption         =   "50km^2＜F≤300km^2的丘陵和山丘区"
            Height          =   255
            Left            =   100
            TabIndex        =   36
            Top             =   960
            Width           =   3255
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000009&
            Caption         =   "F≤50km^2的丘陵或山丘区，J≤25‰"
            Height          =   255
            Left            =   100
            TabIndex        =   35
            Top             =   720
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000009&
            Caption         =   "50km^2＜F≤300km^2的山区"
            Height          =   255
            Left            =   100
            TabIndex        =   34
            Top             =   480
            Width           =   3255
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000009&
            Caption         =   "F≤50km^2的山区，J≥25‰"
            Height          =   255
            Left            =   100
            TabIndex        =   33
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.TextBox Text9 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1560
         TabIndex        =   29
         Top             =   2685
         Width           =   900
      End
      Begin VB.TextBox Text8 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1560
         TabIndex        =   28
         Top             =   2205
         Width           =   900
      End
      Begin VB.CommandButton Command2 
         Caption         =   "综合无因次单位线推求设计洪水"
         Height          =   615
         Left            =   -74640
         TabIndex        =   25
         Top             =   3300
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "推理公式推求设计洪水"
         Height          =   615
         Left            =   480
         TabIndex        =   24
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Caption         =   "m~θ相关线"
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton Option2 
            Caption         =   "外包线(植被较差的丘陵、山区)"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "平均线(植被好的以森林为主的山区)"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   3255
         End
      End
      Begin VB.Label Label22 
         Caption         =   "万m^3"
         Height          =   375
         Left            =   -72480
         TabIndex        =   44
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "m^3"
         Height          =   255
         Left            =   -72380
         TabIndex        =   43
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "洪水总量="
         Height          =   255
         Left            =   -74520
         TabIndex        =   40
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "洪峰流量="
         Height          =   375
         Left            =   -74520
         TabIndex        =   39
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "万m^3"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "m^3/s"
         Height          =   375
         Left            =   2640
         TabIndex        =   30
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "洪水总量="
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "洪峰流量="
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim Rzi(24) As Single, Rsi(24) As Single, AveSumRRsi(24) As Single, SumRi(24) As Single
Dim F0!, L0!, J0!, b%, I0!, H24dian!, Y0!, i%, j%, k%

Public Sub 最大24h暴雨时程分配计算()
 Dim sa24(8, 8) As Single, sn2(40, 6) As Single, sn3(40, 6) As Single, shdp(40, 24) As Single
 Dim a24(8) As Single, n2(5, 6) As Single, n3(5, 6) As Single, hdp(5, 24) As Single, hdp1(5, 24) As Single
 Dim F1(5) As Integer, H(6) As Integer, F2(8) As Integer, hh(5) As Single, Ri(24) As Single, s As Single
 Open App.Path & "\暴雨查算手册原始数据\a24-F.txt" For Input As #1
 Call 赋初值
 For i = 1 To 8
  For j = 1 To 8
   Input #1, sa24(i, j)
  Next
 Next
 For i = 1 To 8
  a24(i) = sa24(b, i)
 Next
 Open App.Path & "\暴雨查算手册原始数据\n2-F-H24面.txt" For Input As #2
 For i = 1 To 40
  For j = 1 To 6
   Input #2, sn2(i, j)
  Next
 Next
 For i = 1 To 5
  For j = 1 To 6
   n2(i, j) = sn2(5 * (b - 1) + i, j)
  Next
 Next
 Open App.Path & "\暴雨查算手册原始数据\n3-F-H24面.txt" For Input As #3
 For i = 1 To 40
  For j = 1 To 6
   Input #3, sn3(i, j)
  Next
 Next
 For i = 1 To 5
  For j = 1 To 6
   n3(i, j) = sn3(5 * (b - 1) + i, j)
  Next
 Next
 Open App.Path & "\暴雨查算手册原始数据\最大24h概化雨型时程分配表.txt" For Input As #4
 For i = 1 To 40
  For j = 1 To 24
   Input #4, shdp(i, j)
  Next
 Next
 For i = 1 To 5
  For j = 1 To 24
   hdp(i, j) = shdp(5 * (b - 1) + i, j)
  Next
 Next
 F1(1) = 0
 F1(2) = 100
 F1(3) = 300
 F1(4) = 500
 F1(5) = 1000
 F2(1) = 0
 F2(2) = 50
 F2(3) = 100
 F2(4) = 150
 F2(5) = 200
 F2(6) = 300
 F2(7) = 500
 F2(8) = 1000
 H(1) = 50
 H(2) = 100
 H(3) = 150
 H(4) = 200
 H(5) = 300
 H(6) = 500
 i = 3
 Do Until (F2(i) >= F0)
  i = i + 1
 Loop
 a24dian = a24(i - 2) + (a24(i - 2) - a24(i - 1)) / (F2(i - 2) - F2(i - 1)) * (F0 - F2(i - 2)) + ((a24(i - 2) - a24(i - 1)) / (F2(i - 2) - F2(i - 1)) - (a24(i - 1) - a24(i)) / (F2(i - 1) - F2(i))) / (F2(i - 2) - F2(i)) * (F0 - F2(i - 2)) * (F0 - F2(i - 1)) '求点面折算系数
 h24mian = H24dian * a24dian    '求最大24h面雨量
 i = 3
 Do Until (F1(i) >= F0)
  i = i + 1
 Loop
 j = 3
 Do Until (H(j) >= h24mian)
  j = j + 1
 Loop
 n2a = n2(i - 2, j) + (n2(i - 2, j) - n2(i - 1, j)) / (F1(i - 2) - F1(i - 1)) * (F0 - F1(i - 2)) + ((n2(i - 2, j) - n2(i - 1, j)) / (F1(i - 2) - F1(i - 1)) - (n2(i - 1, j) - n2(i, j)) / (F1(i - 1) - F1(i))) / (F1(i - 2) - F1(i)) * (F0 - F1(i - 2)) * (F0 - F1(i - 1))
 n2b = n2(i - 2, j - 1) + (n2(i - 2, j - 1) - n2(i - 1, j - 1)) / (F1(i - 2) - F1(i - 1)) * (F0 - F1(i - 2)) + ((n2(i - 2, j - 1) - n2(i - 1, j - 1)) / (F1(i - 2) - F1(i - 1)) - (n2(i - 1, j - 1) - n2(i, j - 1)) / (F1(i - 1) - F1(i))) / (F1(i - 2) - F1(i)) * (F0 - F1(i - 2)) * (F0 - F1(i - 1))
 n2c = n2(i - 2, j - 2) + (n2(i - 2, j - 2) - n2(i - 1, j - 2)) / (F1(i - 2) - F1(i - 1)) * (F0 - F1(i - 2)) + ((n2(i - 2, j - 2) - n2(i - 1, j - 2)) / (F1(i - 2) - F1(i - 1)) - (n2(i - 1, j - 2) - n2(i, j - 2)) / (F1(i - 1) - F1(i))) / (F1(i - 2) - F1(i)) * (F0 - F1(i - 2)) * (F0 - F1(i - 1))
 n2mian = n2c + (n2c - n2b) / (H(j - 2) - H(j - 1)) * (h24mian - H(j - 2)) + ((n2c - n2b) / (H(j - 2) - H(j - 1)) - (n2b - n2a) / (H(j - 1) - H(j))) / (H(j - 2) - H(j)) * (h24mian - H(j - 2)) * (h24mian - H(j - 1))   '求系数n2mian
 n3a = n3(i - 2, j) + (n3(i - 2, j) - n3(i - 1, j)) / (F1(i - 2) - F1(i - 1)) * (F0 - F1(i - 2)) + ((n3(i - 2, j) - n3(i - 1, j)) / (F1(i - 2) - F1(i - 1)) - (n3(i - 1, j) - n3(i, j)) / (F1(i - 1) - F1(i))) / (F1(i - 2) - F1(i)) * (F0 - F1(i - 2)) * (F0 - F1(i - 1))
 n3b = n3(i - 2, j - 1) + (n3(i - 2, j - 1) - n3(i - 1, j - 1)) / (F1(i - 2) - F1(i - 1)) * (F0 - F1(i - 2)) + ((n3(i - 2, j - 1) - n3(i - 1, j - 1)) / (F1(i - 2) - F1(i - 1)) - (n3(i - 1, j - 1) - n3(i, j - 1)) / (F1(i - 1) - F1(i))) / (F1(i - 2) - F1(i)) * (F0 - F1(i - 2)) * (F0 - F1(i - 1))
 n3c = n3(i - 2, j - 2) + (n3(i - 2, j - 2) - n3(i - 1, j - 2)) / (F1(i - 2) - F1(i - 1)) * (F0 - F1(i - 2)) + ((n3(i - 2, j - 2) - n3(i - 1, j - 2)) / (F1(i - 2) - F1(i - 1)) - (n3(i - 1, j - 2) - n3(i, j - 2)) / (F1(i - 1) - F1(i))) / (F1(i - 2) - F1(i)) * (F0 - F1(i - 2)) * (F0 - F1(i - 1))
 n3mian = n3c + (n3c - n3b) / (H(j - 2) - H(j - 1)) * (h24mian - H(j - 2)) + ((n3c - n3b) / (H(j - 2) - H(j - 1)) - (n3b - n3a) / (H(j - 1) - H(j))) / (H(j - 2) - H(j)) * (h24mian - H(j - 2)) * (h24mian - H(j - 1))   ' 求系数n3mian
 H1 = h24mian * 24 ^ (n3mian - 1) * 6 ^ (n2mian - n3mian)    '求1h降雨
 H3 = h24mian * 24 ^ (n3mian - 1) * 6 ^ (n2mian - n3mian) * 3 ^ (1 - n2mian)    '求3h降雨
 H6 = h24mian * 24 ^ (n3mian - 1) * 6 ^ (1 - n3mian)    '求6h降雨
 H12 = h24mian * 24 ^ (n3mian - 1) * 12 ^ (1 - n3mian)    '求12h降雨
 hh(1) = H1
 hh(2) = H3 - H1
 hh(3) = H6 - H3
 hh(4) = H12 - H6
 hh(5) = h24mian - H12
 For i = 1 To 5
  For j = 1 To 24
   hdp1(i, j) = hdp(i, j) * (hh(i) / 100)
  Next
 Next
 For j = 1 To 24
  s = 0
  For i = 1 To 5
   s = s + hdp1(i, j)
  Next
  Ri(j) = s
 Next
 s = 0
 For i = 1 To 24
  s = s + Ri(i)
  SumRi(i) = s
 Next
 i = 1
 Do Until (SumRi(i) - I0 > 0)
  i = i + 1
 Loop
 k = i
 For i = 1 To k - 1           'Rzi(24)为扣除初损后的总径流的时程分配
  Rzi(i) = 0
 Next
 Rzi(k) = SumRi(k) - I0
 For i = k + 1 To 24
  Rzi(i) = Ri(i)
 Next
 For i = 1 To 24              'Rsi(24)为地表径流的时程分配
  Rsi(i) = Rzi(i) * Y0
 Next
 Close #1, #2, #3, #4
End Sub

Public Sub 赋初值()
 F0 = Val(Text1.Text)
 L0 = Val(Text2.Text)
 J0 = Val(Text3.Text)
 b = Int(Val(Text4.Text))
 I0 = Val(Text5.Text)
 H24dian = Val(Text6.Text)
 Y0 = Val(Text7.Text)
End Sub

Function Rt(t1 As Single) As Single
 i = 3
 Do Until (i >= t1)
  i = i + 1
 Loop
 Rt = AveSumRRsi(i - 2) + (AveSumRRsi(i - 2) - AveSumRRsi(i - 1)) / (-1) * (t1 - (i - 2)) + ((AveSumRRsi(i - 2) - AveSumRRsi(i - 1)) / (-1) - (AveSumRRsi(i - 1) - AveSumRRsi(i)) / (-1)) / (-2) * (t1 - (i - 2)) * (t1 - (i - 1))
End Function

Private Sub Command1_Click()
 Dim RRsi(24) As Single, SumRRsi(24) As Single, flag As Integer, SumRsi As Single, t As Single, SumRz As Single, FLB As Single, FF0(39, 24) As Single, FF01(39) As Single, FF1(80) As Single, FF2(80) As Single, FFZ(80) As Single, K0 As Integer, K1 As Integer, K2 As Integer, Qzzm As Single
 Call 赋初值
 If (F0 <= 0 Or L0 <= 0 Or J0 <= 0 Or b <= 0 Or I0 < 0 Or H24dian <= 0 Or Y0 <= 0 Or Y0 >= 1) Then MsgBox ("输入的数据有误!")
 If (F0 > 0 And L0 > 0 And J0 > 0 And b > 0 And I0 >= 0 And H24dian > 0 And Y0 > 0 And Y0 < 1) Then Call 最大24h暴雨时程分配计算
 For i = 1 To 24
  RRsi(i) = Rsi(i)
 Next
 For i = 1 To 23
  flag = 0
  For j = 1 To 23
   If (RRsi(j + 1) > RRsi(j)) Then
    flag = 1
    t = RRsi(j)
    RRsi(j) = RRsi(j + 1)
    RRsi(j + 1) = t
   End If
  Next
  If (flag = 0) Then Exit For
 Next
 u = 0
 For i = 1 To 25 - k
  u = u + RRsi(i)
  SumRRsi(i) = u
  AveSumRRsi(i) = SumRRsi(i) / i
 Next
 am = L0 / (F0 ^ 0.25 * (J0 / 1000) ^ 0.3333)
 If Option1.Value = True Then
  If am > 0 And am <= 25 Then bm = 0.145 * am ^ 0.489
  If am > 25 Then bm = 0.0228 * am ^ 1.067
 End If
 If Option2.Value = True Then
  If am > 0 And am <= 22 Then bm = 0.183 * am ^ 0.489
  If am > 22 Then bm = 0.0284 * am ^ 1.093
 End If
 t0 = 3
 Qmi = 0.278 * F0 * Rt(3)
 t = 0.278 * L0 / (bm * (J0 / 1000) ^ 0.3333 * Qmi ^ 0.25)
 Do Until (Abs(t - t0) < 0.001)
  t0 = t
  Qmi = 0.278 * F0 * Rt(t)
  t = 0.278 * L0 / (bm * (J0 / 1000) ^ 0.3333 * Qmi ^ 0.25)
 Loop
 Qmi = 0.278 * F0 * Rt(t)
 SumRzi = 0
 For i = 1 To 24
  SumRzi = SumRzi + Rzi(i)
 Next
 FLB = Qmi / (SumRzi * Y0 * F0 / 3.6 * 1)
 If (FLB >= 0.07 And FLB <= 0.3) Then
  Open App.Path & "\暴雨查算手册原始数据\径流分配系数表.txt" For Input As #7
  For i = 1 To 39
   For j = 1 To 24
    Input #7, FF0(i, j)
   Next
  Next
  If ((FLB - 0.07) * 100 - Int((FLB - 0.07) * 100) < 0.5) Then
   K0 = Int((FLB - 0.07) * 100) + 1
  End If
  If ((FLB - 0.07) * 100 - Int((FLB - 0.07) * 100) >= 0.5) Then
   K0 = Int((FLB - 0.07) * 100) + 2
  End If
  For i = 1 To 39
   FF01(i) = FF0(i, K0)
  Next
  i = 1
  Do Until (FF01(i) - FF01(i + 1) > 0)
   i = i + 1
  Loop
  K1 = i
  FF01(K1 + 1) = FF01(K1 + 1) - FLB + FF01(K1)     '调整第K1+1项
  FF01(K1) = FLB  '调整第K1项
  For i = 1 To 39
   FF1(i) = FF01(i) * (SumRzi * Y0 * F0 / 3.6)
  Next
  i = 2
  Do Until (FF01(i) = 0)
   i = i + 1
  Loop
  K2 = i    'K2为地面径流的底宽
  For i = 1 To K2
   FF2(i) = ((SumRzi * (1 - Y0) * F0 / 3.6) / 27) * (i / 27)
  Next
  For i = (K2 + 1) To (2 * K2 - 1)
   FF2(i) = ((SumRzi * (1 - Y0) * F0 / 3.6) / 27) * ((2 * K2 - i) / 27)
  Next
  For i = 1 To (2 * K2 - 1)
   FFZ(i) = FF1(i) + FF2(i)
  Next
  Qzzm = 0
  For i = 1 To 2 * K2 - 1
   If (FFZ(i) > Qzzm) Then
    Qzzm = FFZ(i)
   End If
  Next
  Text8.Text = Format(Qzzm, "###0.00")
  Text9.Text = Format(SumRzi / 10000, "###0.0000")
  Open App.Path & "\推理公式法成果.txt" For Output As #8
  Print #8, "洪峰流量=" & Format(Qzzm, "###0.00") & "m^3/s"
  Print #8, "洪水总量=" & Format(SumRzi / 10000, "####0.0000") & "万m^3"
  For i = 1 To 2 * K2 - 1
   Print #8, Format(i - 1, "000"), Format(FFZ(i), "00000.000")
  Next
  Close #7, #8
 End If
 If ((FLB < 0.07) Or (FLB > 0.3)) Then
  Text8.Text = Format(Qmi + 0.026 * F0, "###0.00")
  Text9.Text = Format(SumRzi / 10000, "###0.0000")
  Open App.Path & "\推理公式法成果.txt" For Output As #8
  Print #8, "FLB=" & FLB
  Print #8, "洪峰流量=" & Format(Qmi + 0.026 * F0, "###0.00") & "m^3/s"
  Print #8, "洪水总量=" & Format(SumRzi / 10000, "####0.0000") & "万m^3"
  Close #8
 End If
End Sub

Private Sub Command2_Click()
 Dim ki%, kj%
 Dim hnswycqi(45, 6) As Single, wycqi(45) As Single, sdqi(45) As Single, qi(200, 24), Rs(24) As Single, SumRs(200) As Single, Qd(200) As Single, Qz(200) As Single, Ts As Integer
 Call 赋初值
 If (F0 <= 0 Or L0 <= 0 Or J0 <= 0 Or b <= 0 Or I0 < 0 Or H24dian <= 0 Or Y0 <= 0 Or Y0 >= 1) Then MsgBox ("输入的数据有误!")
 If (F0 > 0 And L0 > 0 And J0 > 0 And b > 0 And I0 >= 0 And H24dian > 0 And Y0 > 0 And Y0 < 1) Then Call 最大24h暴雨时程分配计算
 If (Option3.Value = True) Then
  ki = 1
  kj = 22
 End If
 If (Option4.Value = True) Then
  ki = 2
  kj = 33
 End If
 If (Option5.Value = True) Then
  ki = 3
  kj = 33
 End If
 If (Option6.Value = True) Then
  ki = 4
  kj = 35
 End If
 If (Option7.Value = True) Then
  ki = 5
  kj = 39
 End If
 If (Option8.Value = True) Then
  ki = 6
  kj = 45
 End If
 Open App.Path & "\暴雨查算手册原始数据\综合无因次单位线表.txt" For Input As #5
 For i = 1 To 45
  For j = 1 To 6
   Input #5, hnswycqi(i, j)
  Next
 Next
 Open App.Path & "\经验单位线法成果.txt" For Output As #6
 For i = 1 To 45
  wycqi(i) = hnswycqi(i, ki)
  sdqi(i) = wycqi(i) * 2.778 * F0
 Next
 For i = 1 To 25 - k
  Rs(i) = Rsi(k + i - 1)
 Next
 For i = 1 To 25 - k
  For j = 1 To 45
   qi(i + j - 1, i) = sdqi(j) * (Rs(i) / 10)
  Next
 Next
 For i = 1 To 69 - k
  s = 0
  For j = 1 To 25 - k
   s = s + qi(i, j)
  Next
  SumRs(i) = s
 Next
 SumRz = 0
 For i = 1 To 24
  SumRz = SumRz + Rzi(i)
 Next
 Ts = kj + 24 - k
 Qdm = F0 * (SumRz * (1 - Y0)) / (3.6 * Ts)
 For i = 1 To Ts
  Qd(i) = Qdm / Ts * i
 Next
 For i = Ts + 1 To 2 * Ts
  Qd(i) = Qdm * ((2 * Ts - i) / Ts)
 Next
 For i = 1 To 2 * Ts
  Qz(i) = SumRs(i) + Qd(i)
 Next
 Qzm = 0
 For i = 1 To 2 * Ts
  If (Qz(i) > Qzm) Then
   Qzm = Qz(i)
  End If
 Next
 Wz = SumRz * F0 * 1000
 Text10.Text = Format(Qzm, "###0.00")
 Text11.Text = Format(SumRz / 10000, "###0.0000")
 Print #6, "洪峰流量=" & Format(Qzm, "###0.00") & "m^3/s"
 Print #6, "洪水总量=" & Format(SumRz / 10000, "###0.0000") & "万m^3"
 For i = 1 To 2 * Ts
  Print #6, Format(i - 1, "000"), Format(Qz(i), "00000.000")
 Next
 Close #5, #6
End Sub
