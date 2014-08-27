VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form21 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "DirectX"
   ClientHeight    =   7395
   ClientLeft      =   225
   ClientTop       =   615
   ClientWidth     =   13740
   Icon            =   "Form21.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13740
   StartUpPosition =   2  '螢幕中央
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog openfileDialog 
      Left            =   960
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   " "
      Height          =   3015
      Index           =   1
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton Option1 
         Caption         =   "800 X 600"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1024 X 768"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1280 X 1024"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1600 X 1200"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2048 X 1536"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "依原圖的解析度"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "640 X 480"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   75
         Picture         =   "Form21.frx":EA72
         Stretch         =   -1  'True
         Top             =   30
         Width           =   300
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   2280
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2280
         X2              =   2280
         Y1              =   240
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   240
         Y2              =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "調整圖片解析度"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   75
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   285
         Index           =   1
         Left            =   2010
         Picture         =   "Form21.frx":F07E
         Top             =   0
         Width           =   285
      End
      Begin VB.Image Image2 
         Height          =   315
         Index           =   1
         Left            =   0
         Picture         =   "Form21.frx":F534
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2325
      End
   End
   Begin VB.Image Image3 
      Height          =   285
      Index           =   1
      Left            =   480
      Picture         =   "Form21.frx":1054E
      Top             =   6000
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   285
      Index           =   0
      Left            =   120
      Picture         =   "Form21.frx":109F5
      Top             =   6000
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "選擇圖片資料夾"
      Index           =   0
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "調整圖片解析度"
      Index           =   1
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "上一頁(↑)"
      Index           =   2
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "下一頁(↓)"
      Index           =   3
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "說明(H)"
      Index           =   4
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'作者:小賢 lbt95@yahoo.com.tw
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim Texture() As Direct3DTexture8

Dim Pic_Back As String
Const PR As Byte = 2 'ma+1)背景 ma+2) 遮罩
'Const Pic_Drive As String = ""

Dim Cosine(360) As Single
Dim Sine(360) As Single
Dim Xdis() As Single, Ydis() As Single
Dim O_Xdis() As Single, O_Ydis() As Single
Dim D_Xdis() As Single, D_Ydis() As Single
Dim fileArray() As String '圖片路徑含名稱陣列
Dim ma As Integer '圖片數量
Dim S(11) As Integer '0)g圖片移動精細度 1)每排圖數 2)圖片迴圈起始 3)圖片迴圈終點 4)上下頁 5)頁數累加 6)程式即將結束 7)滑鼠沒按到任何的圖 8)亂數背景 9)顯示DX文字 10)選擇圖片資料夾 11)開頭動畫
Dim Pic_W As Single '圖寬
Dim Pic_H As Single '圖高
Dim PiX As Single 'X解析度
Dim PiY As Single 'Y解析度
Dim old_X As Single '移動模擬視窗之暫存X
Dim old_Y As Single

Dim old_Vertex_x As Single
Dim old_Vertex_y As Single
Dim P_Move(3) As Single '0)按住滑鼠左鍵 1)滑鼠點下去的X位置 2)滑鼠點下去的Y位置 3)滑鼠選中的圖片
Dim G_X As Single '中間值
Dim G_Y As Single '中間值
Dim AF() As Integer '優先順序號碼(唯一一個)
Dim File_ele As D3DXIMAGE_INFO '檔案資訊
Dim Bermuda '百慕達三角洲.........................只要在此位置宣告必定錯誤
Dim DXWord As String 'DX文字
Dim Form_Name As String '表單原始名稱
Dim XP_Active(4) As String '標題列動畫

Dim P_Size() As Pic_DX
Private Type Pic_DX
    A_Move As Byte '圖片已移動次數（歸位)
    A_Turn As Integer '圖片的旋轉角度
    OpenTurn As Byte '旋轉開關
    Xcenter As Single
    Ycenter As Single
    O_XCen As Single '起始中心
    O_YCen As Single '起始中心
    D_XCen As Single '目的中心
    D_YCen As Single '目的中心
    Large As Boolean '縮放
    Alpha As Byte
    Dis_Large As Byte '不允許縮放
    Perfect As Byte '細緻程度
End Type
Private Sub Form_Load() '表單載入
Call Option1_Click(1) '預設為800 X 600
Call Three 'Δ函數
Call Load_START '只讀一次
Call Form_KeyDown(72, 0) '顯示說明
Call Dx_START(True, Me) 'DX初始
Call DXFont 'DX文字
End Sub
Private Sub Form_Activate()
Do
    Call START(openfileDialog.fileName) '初始化
    Call BackPicture(Pic_Back) '背景
    Call Dx_Texture   'DX材質
    Call Core '核心
    If S(10) = 1 Then Me.Caption = Form_Name: Call Ma_Clear
    If S(6) = 1 Then S(6) = 2: Unload Me '結束程式(ESC)
Loop Until S(6) <> 0
Call Ma_Clear '清除數值資料
End Sub
Private Sub START(fileName As String) '初始化 a)預設為畢業光碟程式路徑
Dim f As Byte, j As Byte, strArr() As String
Dim length As Long

strArr = Split(fileName, vbNullChar)
length = UBound(strArr)

If length > 0 Then
    ma = length - 1
    ReDim fileArray(ma) As String
    
    For f = 1 To length
        fileArray(f - 1) = strArr(f)
    Next
Else
    ma = length
    fileArray = strArr
End If

If ma = -1 Then ma = 0

ReDim Texture(ma + PR) As Direct3DTexture8
ReDim Vertex(3, ma + PR) As TLVERTEX
ReDim P_Size(ma) As Pic_DX
ReDim AF(ma) As Integer
ReDim Xdis(3, ma) As Single, Ydis(3, ma) As Single
ReDim O_Xdis(3, ma) As Single, O_Ydis(3, ma) As Single
ReDim D_Xdis(3, ma) As Single, D_Ydis(3, ma) As Single

S(0) = 10 '(載入材質部分、右鍵)'圖片移動精細度
S(1) = 10 '每列共10張 10 X 10 = 100張(全)
S(2) = 0 '-----------圖片迴圈初始值
S(3) = IIf(ma < S(1) ^ 2 - 1, ma, S(1) ^ 2 - 1) '圖片迴圈終止值s

G_X = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
Pic_W = D3dParam.BackBufferWidth / S(1)
Pic_H = D3dParam.BackBufferHeight / S(1)
End Sub
Private Sub Dx_Texture() '載入材質
Dim a(1) As Single, b(1) As Single, f As Integer, j As Integer, i As Byte, t As Long

If UBound(fileArray) = -1 Then Exit Sub

t = GetTickCount '計時開始
Call Central(a(0), a(1), b(0), b(1)) '發牌起始計算
For f = S(2) To S(3)
    Set Texture(f) = LoadTexture(fileArray(f), Pic_W, Pic_H) '讀取圖片
    Me.Caption = Form_Name & "　" & f + 1 & " / " & S(3) + 1 & "　圖片載入中　" & XP_Active(f \ 3 Mod 5) & "　共" & ma + 1 & "張　" & "耗時" & (GetTickCount - t) \ 1000 & "秒"
    
    Vertex(0, f) = Ver(a(0), b(0), 0, 0) '起始位置
    Vertex(1, f) = Ver(a(1), b(0), 1, 0)
    Vertex(2, f) = Ver(a(0), b(1), 0, 1)
    Vertex(3, f) = Ver(a(1), b(1), 1, 1)

    j = Deal(f) '目地位置計算
    With P_Size(f)
        .Xcenter = (Vertex(0, f).X + Vertex(1, f).X) / 2 '中心
        .Ycenter = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
        For i = 0 To 3
            Call Auto_Dis(f, Pic_W, Pic_H)
        Next
        .D_XCen = D_Xdis(3, f) + Pic_W * (j Mod S(1)) '目的中心
        .D_YCen = D_Ydis(3, f) + Pic_H * (f \ S(1) Mod S(1))
        .Perfect = S(0) '移動細緻度
        .Alpha = 255
        .A_Turn = 270
        .OpenTurn = 2
    End With
    AF(f) = f '優先順序
    Call Swap(f)
    DoEvents
    
    If S(6) = 1 Or S(4) <> 0 Or S(10) = 1 Then Exit Sub '如果程式將要結束則離開迴圈
    Call Render(f) '秀圖
    Call Pic_Move(f) '只移動目前頁面的圖
Next
Me.Caption = Form_Name & "　" & f & " / " & S(3) + 1 & "　共" & ma + 1 & "張　" & "耗時" & (GetTickCount - t) \ 1000 & "秒"
Do Until P_Size(S(3)).A_Move = P_Size(S(3)).Perfect '後續動畫延續
    DoEvents
    If S(6) = 1 Or S(4) <> 0 Or S(10) = 1 Then Exit Sub '如果程式將要結束則離開迴圈
    Call Render(S(3)) '秀圖
    Call Pic_Move(S(3)) '只移動目前頁面的圖
    Sleep (20)
Loop
S(11) = 1 '開頭動畫已結束
End Sub
Private Sub Core() '核心程式 ☆
Do
    DoEvents
    If S(4) <> 0 Then Call Page '如果按了上下頁則
    If S(10) = 1 Then Exit Sub '如果按了圖片資料夾
    Call Pic_Move(S(3)) '移動
    Call Render(S(3)) '秀圖
    Sleep (20)
Loop While S(6) = 0 '結束程式
End Sub
Private Function Pic_coll(Ax As Single, Ay As Single, f As Integer, Bx As Single, By As Single) As Boolean
If Ax > Bx And Ax < Bx + Xdis(3, f) * 2 And Ay > By And Ay < By + Ydis(3, f) * 2 Then Pic_coll = True
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '滑鼠按下
Dim f As Integer, j As Integer, w As Single, h As Single, tx As Single, ty As Single
S(9) = 0 '清除DX文字
P_Move(1) = X / G_X
P_Move(2) = Y / G_Y
For f = S(2) To S(3)
    If Pic_coll(P_Move(1), P_Move(2), AF(f), Vertex(0, AF(f)).X, Vertex(0, AF(f)).Y) Then
        P_Move(3) = AF(f)
        old_Vertex_x = P_Size(AF(f)).Xcenter - P_Move(1)
        old_Vertex_y = P_Size(AF(f)).Ycenter - P_Move(2)
        Call Swap(AF(f)) '優先順序交換
        Exit For
    End If
Next
S(7) = IIf(f <> S(3) + 1, 0, 1) '是否沒有點到圖片

Select Case Button
    Case 1 '左鍵
        P_Move(0) = 1 '按住左鍵
        If f = S(3) + 1 Then
            P_Move(0) = 0 '取消左鍵
            For f = S(2) To S(3)
                P_Move(3) = f
                Call Form_MouseDown(2, 1, X, Y) '如果是放大的狀態則縮小
            Next
        End If
    Case 2 '右鍵
        If f <> S(3) + 1 Or Shift = 1 Then
            With P_Size(P_Move(3))
                If .Large Then Call Form_MouseDown(3, 1, X, Y) '如果是放大的狀態則縮小
                
                tx = .D_XCen
                ty = .D_YCen
                j = Deal(P_Move(3)) '目地位置計算
                .D_XCen = D_Xdis(3, P_Move(3)) + Pic_W * (j Mod S(1)) '目地位置
                .D_YCen = D_Ydis(3, P_Move(3)) + Pic_H * ((P_Move(3) \ S(1)) Mod S(1))
                
                If tx = .D_XCen And ty = .D_YCen Then Exit Sub
                
                .Perfect = 10 '細緻度為20
                .A_Move = 0
            End With
        End If
    Case 3 '放大
        If f <> S(3) + 1 Or Shift = 1 Then
            With P_Size(P_Move(3))
                .Perfect = 10 '細緻度為20
                .A_Move = 0
                .Large = Not .Large '縮放
                .Alpha = 100 '透明值
                Debug.Print .A_Turn
                If .Large Then
                    If PiX = -1 Then '如果是原生解析度或是同學一句話則
                        Set Texture(P_Move(3)) = LoadTexture(fileArray(P_Move(3)))
                        w = File_ele.Width
                        h = File_ele.Height
                    Else
                        Set Texture(P_Move(3)) = LoadTexture(fileArray(P_Move(3)), PiX, PiY)
                        w = IIf(File_ele.Width < File_ele.Height, PiY, PiX) '直式橫式判斷
                        h = IIf(File_ele.Width < File_ele.Height, PiX, PiY)
                    End If
                Else
                    '.OpenTurn = 1 '開啟旋轉
                    Set Texture(P_Move(3)) = LoadTexture(fileArray(P_Move(3)), Pic_W, Pic_H)
                    w = Pic_W '寬
                    h = Pic_H '高
                End If
                Call Auto_Dis(P_Move(3), w, h)
            End With
        End If
End Select
End Sub
Private Sub Swap(ByVal fx As Integer) '優先順序交換演算法ＱＱ (想了很久終於想到這個解決方案)
Dim a As Integer, t As Integer, f As Integer
a = fx
For f = S(2) To S(3)
    t = AF(f)
    AF(f) = a
    If t = fx Then Exit For
    a = t
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If P_Move(0) = 1 Then
    X = X / G_X
    Y = Y / G_Y
    With P_Size(P_Move(3))
        .A_Move = 0
        .Perfect = 1
        .D_XCen = X + old_Vertex_x
        .D_YCen = Y + old_Vertex_y
        .Dis_Large = 1 '因為使用者要移動所以不充許縮放
    End With
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If S(7) = 1 Then Exit Sub

If Button = 1 Then
    P_Move(0) = 0 '放開左鍵
    If P_Size(P_Move(3)).Dis_Large = 0 Then
        ' 防止一開始按開啟檔案時變成無檔案放大
        If UBound(fileArray) >= 0 Then
            Call Form_MouseDown(3, 1, X, Y) '如果使用者目的是要縮放則縮放
        End If
    End If
    P_Size(P_Move(3)).Dis_Large = 0 '清除紀錄
End If
End Sub
Private Sub Pic_Move(ByVal j As Integer) '圖片自動移動
Dim f As Integer, c As Long, i As Byte
Dim X(1) As Single, w As Single
Dim Y(1) As Single, h As Single
Dim a(3) As Single, b(3) As Single, old_Color As Long, d

j = IIf(j > S(3), S(3), j) '限制範圍為本頁
For f = S(2) To j
    If P_Size(f).A_Move < P_Size(f).Perfect Then
        If P_Size(f).A_Move = 0 Then '原始位置
            With P_Size(f) '初中心
                .Xcenter = (Vertex(0, f).X + Vertex(1, f).X) / 2
                .Ycenter = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
                .O_XCen = .Xcenter
                .O_YCen = .Ycenter
            End With
            For i = 0 To 3 '距離
                Xdis(i, f) = Vertex(i, f).X - P_Size(f).O_XCen
                Ydis(i, f) = Vertex(i, f).Y - P_Size(f).O_YCen
                O_Xdis(i, f) = Xdis(i, f)
                O_Ydis(i, f) = Ydis(i, f)
            Next
        End If
        With P_Size(f) '移動核心
            .A_Move = IIf(.OpenTurn = 0, .A_Move + 1, .A_Move + 1)
            .Xcenter = .Xcenter + (.D_XCen - .O_XCen) / .Perfect '中心
            .Ycenter = .Ycenter + (.D_YCen - .O_YCen) / .Perfect
        End With
        For i = 0 To 3 '縮放核心
            Xdis(i, f) = Xdis(i, f) + (D_Xdis(i, f) - O_Xdis(i, f)) / P_Size(f).Perfect '距離
            Ydis(i, f) = Ydis(i, f) + (D_Ydis(i, f) - O_Ydis(i, f)) / P_Size(f).Perfect
        Next
        If S(11) = 0 And P_Size(f).A_Move = P_Size(f).Perfect And P_Size(f).OpenTurn = 2 Then P_Size(f).OpenTurn = 1 '開頭動畫
    End If
    
    With P_Size(f)
        If .Alpha < 251 Then '透明化
            .Alpha = .Alpha + 5
            c = D3DColorARGB(.Alpha, 255, 255, 255)
        End If
        
        If .OpenTurn = 1 Then '旋轉核心
            .A_Turn = (.A_Turn + 17) Mod 361
            For i = 0 To 3 '還原頂點位置
                a(i) = .Xcenter + Xdis(i, f) / 2 + Cosine(.A_Turn) - Sine(.A_Turn) * Xdis(i, f) / 2 '頂點
                b(i) = .Ycenter + Ydis(i, f) + Sine(.A_Turn) + Cosine(.A_Turn) * Xdis(i, f)
            Next
            If .A_Turn <= 270 + 16 And .A_Turn >= 270 - 16 Then .OpenTurn = 0: .A_Turn = 270 '如果旋到正面則關閉旋轉
        Else '不旋轉
            For i = 0 To 3 '還原頂點位置
                a(i) = .Xcenter + Xdis(i, f)
                b(i) = .Ycenter + Ydis(i, f)
            Next
        End If
    End With

    old_Color = Vertex(0, f).color '存入舊的顏色
    Vertex(0, f) = Ver(a(0), b(0), 0, 0)
    Vertex(1, f) = Ver(a(1), b(1), 1, 0)
    Vertex(2, f) = Ver(a(2), b(2), 0, 1)
    Vertex(3, f) = Ver(a(3), b(3), 1, 1)
    d = IIf(c <> 0, c, old_Color) '還原原來的顏色
    For i = 0 To 3
        Vertex(i, f).color = d
    Next
    c = 0
    DoEvents
Next
End Sub
Sub Render(j As Integer) '秀圖
Dim f As Integer
On Error GoTo Err
With D3ddevice
    .BeginScene
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
            .SetTexture 0, Texture(ma + 1) '背景
            .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, ma + 1), Len(Vertex(0, ma + 1))
            For f = j To S(2) Step -1 '繪出'圖片
                .SetTexture 0, Texture(AF(f))
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, AF(f)), Len(Vertex(0, AF(f)))
            Next
            
            If S(9) <> 0 Then
                If S(9) = 2 Then .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, ma + 2), Len(Vertex(0, ma + 2))
                D3DX.DrawText MainFont, &HFFFFFFFF, DXWord, TextRect, DT_TOP
            End If
        .SetRenderState D3DRS_ALPHABLENDENABLE, False
    .EndScene
    .Present ByVal 0, ByVal 0, 0, ByVal 0 'Flip切頁
End With
Exit Sub
Err:
Unload Me
End Sub
Private Sub Load_START() '只讀一次
Dim f As Byte, j As Byte

'支援多檔開啟
openfileDialog.flags = cdlOFNAllowMultiselect + cdlOFNExplorer

'亂數背景
Randomize
S(8) = Int(Rnd * 3) + 1
S(8) = 63 + S(8) * 2 '=65 67 69
Pic_Back = App.Path & "\BackPicture\" & S(8) & ".jpg"

'■□□□□
'□■□□□
'□□■□□
'□□□■□
'□□□□■

For f = 0 To 4
    For j = 0 To 4
        XP_Active(f) = XP_Active(f) & IIf(f = j, "■", "□")
    Next
Next
Form_Name = Me.Caption
End Sub
Function LoadTexture(ByVal fileName, Optional w As Single = D3DX_DEFAULT, Optional h As Single = D3DX_DEFAULT, Optional color As Long = 0) As Direct3DTexture8
On Error Resume Next

Set LoadTexture = D3DX.CreateTextureFromFileEx(D3ddevice, fileName, w, h, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, color, File_ele, ByVal 0)
End Function
Private Sub Auto_Dis(f, w As Single, h As Single) '距離自動化
Dim a As Single, b As Single
a = w / 2
b = h / 2
D_Xdis(0, f) = -a
D_Ydis(0, f) = -b
D_Xdis(1, f) = a
D_Ydis(1, f) = -b
D_Xdis(2, f) = -a
D_Ydis(2, f) = b
D_Xdis(3, f) = a
D_Ydis(3, f) = b
End Sub
Private Sub Page() '換頁
Dim a As Integer, t1 As Integer, t2 As Integer '暫存S2,S3

S(9) = 1 '顯示第幾頁
S(5) = S(5) + S(4) '累加頁數
t1 = S(1) ^ 2 * S(5) '迴圈起始
t2 = (S(1) ^ 2) * (S(5) + 1) - 1 '迴圈終止
t2 = IIf(t2 > ma, ma, t2) '預防超出最大圖片張數
S(4) = 0 '清除
S(11) = 0 '清除已載入開頭動畫
a = (ma + 1) \ S(1) ^ 2
If (ma + 1) Mod S(1) ^ 2 <> 0 Then a = a + 1
DXWord = "第" & S(5) + 1 & "頁／共" & a & "頁　" & t1 + 1 & "～" & t2 + 1

Call Glide(255) '淡出

S(2) = t1
S(3) = t2

Vertex(1, 0).X = -Pic_W '隱藏第0張圖
Vertex(3, 0).X = -Pic_W

D3ddevice.SetTexture 0, Nothing '清除與重建

Erase Texture(), P_Size(), AF()
ReDim Texture(ma + PR), P_Size(ma), AF(ma)

'換背景
S(8) = (S(8) - 65) / 2 '0 2 4 => 0 1 2
S(8) = (S(8) + 1) Mod 3 '換下一張背景
S(8) = S(8) * 2 + 65
Call BackPicture(App.Path & "\BackPicture\" & S(8) & ".jpg") '載入背景

Call Glide(0, 1) '淡入
S(9) = 0

Call Dx_Texture
End Sub
Private Sub Glide(i As Byte, Optional b As Integer = -1) '切換畫面
Dim f As Byte, j As Integer
Do
    i = i + b
    For f = 0 To 3
        Vertex(f, ma + 1).color = D3DColorARGB(i, i, i, i)
        For j = S(2) To S(3)
            Vertex(f, j).color = D3DColorARGB(i, i, i, i)
        Next
    Next
    D3ddevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0, 18, 0
    Call Render(S(3))
Loop While IIf(b = -1, i > 1, i < 254)
End Sub
Private Sub BackPicture(fileName As String) '背景
Dim j As Byte

Set Texture(ma + 1) = LoadTexture(fileName)
Vertex(0, ma + 1) = Ver(0, 0, 0, 0)
Vertex(1, ma + 1) = Ver((D3dParam.BackBufferWidth), 0, 1, 0)
Vertex(2, ma + 1) = Ver(0, (D3dParam.BackBufferHeight), 0, 1)
Vertex(3, ma + 1) = Ver((D3dParam.BackBufferWidth), (D3dParam.BackBufferHeight), 1, 1)

'遮罩
Vertex(0, ma + 2) = Ver(0, 0, 0, 0)
Vertex(1, ma + 2) = Ver((D3dParam.BackBufferWidth), 0, 1, 0)
Vertex(2, ma + 2) = Ver(0, (D3dParam.BackBufferHeight), 0, 1)
Vertex(3, ma + 2) = Ver((D3dParam.BackBufferWidth), (D3dParam.BackBufferHeight), 1, 1)
For j = 0 To 3
    Vertex(j, ma + 2).color = &HAA000000
Next
End Sub
Sub Central(a0 As Single, a1 As Single, b0 As Single, b1 As Single) '發牌起始計算
With D3dParam
    a0 = .BackBufferWidth - Pic_W * 2 '計算位置
    b0 = .BackBufferHeight - Pic_H * 3
    a1 = .BackBufferWidth
    b1 = .BackBufferHeight - Pic_H
End With
End Sub
Function Deal(f) As Integer '發牌目地位置計算
Deal = IIf((f \ S(1)) Mod 2 = 0, f, S(1) - (f + 1) Mod S(1))
End Function
Private Sub Three() 'Δ函數
Dim f As Integer
Const PI As Single = 3.14159265358979
For f = 0 To 360
    Cosine(f) = Cos(f / 180 * PI)
    Sine(f) = Sin(f / 180 * PI)
Next
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38 'UP
        If S(5) > 0 Then S(4) = -1
    Case 40 'Down
        If S(3) >= ma Then S(5) = -1
        S(4) = 1
    Case 72 'Help
        DXWord = "滑鼠：" & vbCrLf & "　　左鍵：單擊放大" & vbCrLf & "　　右鍵：圖片歸位" & vbCrLf & "鍵盤：" & vbCrLf & "　　↑：上一頁" & vbCrLf & "　　↓：下一頁" & vbCrLf & "　　Ｈ：說明"
        S(9) = (S(9) + 2) Mod 4
    Case 27 'ESC
        Unload Me
    Case 13
        If Shift = 4 Then
            Call Dx_START(False, Me)
            Call START(openfileDialog.fileName) '初始化
            Call BackPicture(Pic_Back) '背景
            Call Dx_Texture 'DX材質
        End If
Case Else
    S(9) = 0 '不顯示DX文字
End Select
End Sub
Private Sub Form_Resize()
G_X = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
End Sub
Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Label1_MouseDown(Index, Button, Shift, X, Y)
End Sub
Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Label1_MouseMove(Index, Button, Shift, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    old_X = X
    old_Y = Y
    Frame1(Index).ZOrder 0
End If
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Frame1(Index).Left = Frame1(Index).Left + X - old_X
    Frame1(Index).Top = Frame1(Index).Top + Y - old_Y
End If
End Sub
Private Sub Pic_Drives_Click(Index As Integer) '檔案功能表
Select Case Index
    Case 0 '選擇圖片資料夾
        openfileDialog.CancelError = True
        openfileDialog.MaxFileSize = 10240
        
        On Error GoTo Err
        openfileDialog.fileName = ""
        openfileDialog.ShowOpen
        S(10) = 1 ' 重新選擇圖片
    Case 1 '放大解析度
        Frame1(1).Visible = True
    Case 2 '上一頁
        Call Form_KeyDown(38, 0)
    Case 3 '下一頁
        Call Form_KeyDown(40, 0)
    Case 4 '說明
        Call Form_KeyDown(72, 0)
End Select

Err:
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(Index).Picture = Image3(1).Picture
End Sub
Private Sub Image1_Click(Index As Integer) '模擬X關閉
Frame1(Index).Visible = False
End Sub
Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(Index).Picture = Image3(0).Picture
End Sub
Private Sub Option1_Click(Index As Integer)
Dim a() As String
If Index = 6 Then PiX = -1: Exit Sub
a = Split(Option1(Index).Caption, " X ")
PiX = a(0) 'X解析度
PiY = a(1) 'Y解析度
End Sub
Private Sub Ma_Clear() '清除數值資料
Dim t As Byte '背景數值暫存
t = S(8)
Erase Texture(), Vertex(), P_Size(), AF(), P_Move(), Xdis(), Ydis(), O_Xdis(), O_Ydis(), D_Xdis(), D_Ydis(), S(), fileArray '清除資料(換圖片目錄)
S(8) = t
End Sub
Private Sub Form_Unload(Cancel As Integer) '表單移除
If S(6) < 2 Then '即將結束
    S(6) = 1
    Cancel = 1
Else '結束
    Call UnDX
End If
End Sub
