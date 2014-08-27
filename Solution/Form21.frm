VERSION 5.00
Begin VB.Form Form21 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
   Caption         =   "DirectX 0.7"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  '像素
   ScaleWidth      =   792
   StartUpPosition =   2  '螢幕中央
   Begin VB.FileListBox File1 
      Height          =   1350
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "選擇圖片目錄"
      Height          =   4215
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "取消"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "確定"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Height          =   3030
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "圖片目錄"
      Index           =   0
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "上一頁(↑)"
      Index           =   1
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "下一頁(↓)"
      Index           =   2
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim D3DX As D3DX8
Dim Texture() As Direct3DTexture8

Const Pic_Drive As String = "C:\Windows\Web\Wallpaper\Nature" '預設圖片目錄
Dim Cosine(360) As Single
Dim Sine(360) As Single
Dim Xcenter() As Single, Ycenter() As Single
Dim Xdis() As Single, Ydis() As Single
Dim ma As Integer '圖片數量
Dim Form_Name As String '表單名稱
Dim s(7) As Integer '0)g圖片移動精細度 1)每排圖數 2)圖片迴圈起始 3)圖片迴圈終點 4)上下頁 5)頁數累加 6)程式即將結束 7)換圖片目錄
Dim Pic_W As Single '圖寬
Dim Pic_H As Single '圖高

Dim old_Vertex_x As Single
Dim old_Vertex_y As Single
Dim oSx As Single '開始的寬
Dim oSy As Single '開始的高
Dim P_Move(4) As Single '0)按住滑鼠左鍵 1)滑鼠點下去的X位置 2)滑鼠點下去的Y位置 3)滑鼠選中的圖片 4)Frame1 的滑鼠左鍵
Dim G_x As Single '中間值
Dim G_Y As Single '中間值
Dim AF() As Integer '優先順序號碼(唯一一個)
Dim File_ele As D3DXIMAGE_INFO '檔案資訊
Dim old_X As Single '滑鼠點下去的X(Frame1)
Dim old_Y As Single '滑鼠點下去的Y
Dim XP_Active(4) As String '標題動畫

Dim P_Size() As Pic_DX
Private Type Pic_DX
    A_Move As Integer '圖片已移動次數（歸位)
    A_Turn As Integer '圖片已旋轉次數
    Left As Single '來源
    Top As Single
    Width As Single
    Height As Single
    XCen As Single
    YCen As Single
    D_Left As Single '目地
    D_Top As Single
    D_Width As Single
    D_Height As Single
    D_XCen As Single
    D_YCen As Single
    Large As Boolean '縮放
End Type

Private Sub Form_Load() '表單載入
Call Dx_START 'DX初始
Call Three 'Δ函數

Dir1.Path = Pic_Drive '圖片目錄
Do
    Call START(Dir1.Path) '初始化
    Call Dx_Texture 'DX材質
    Call Core '核心
    If s(6) = 1 Then s(6) = 2: Unload Me '結束程式(ESC)
    If s(7) = 1 Then Me.Caption = Form_Name: Erase Texture(), Vertex(), P_Size(), AF(), P_Move(), Xcenter(), Ycenter(), Xdis(), Ydis(), s(), XP_Active() '清除資料(換圖片目錄)
Loop Until s(6) <> 0
End Sub
Private Sub START(a As String) '初始化 a)預設為畢業光碟程式路徑
Dim f As Byte, j As Byte, b As String

Set D3DX = New D3DX8

File1.Path = a '圖片路徑
ma = File1.ListCount - 1
If ma = -1 Then ma = 0

ReDim Texture(ma) As Direct3DTexture8
ReDim Vertex(3, ma) As TLVERTEX
ReDim P_Size(ma) As Pic_DX
ReDim AF(ma) As Integer
ReDim Xcenter(ma) As Single, Ycenter(ma) As Single
ReDim Xdis(3, ma) As Single, Ydis(3, ma) As Single

s(0) = 10 '(載入材質部分、右鍵)'圖片移動精細度
s(1) = 10 '每列共10張 10 X 10 = 100張(全)
s(2) = 0 '-----------圖片迴圈初始值
s(3) = IIf(ma < s(1) ^ 2 - 1, ma, s(1) ^ 2 - 1) '圖片迴圈終止值

G_x = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
Pic_W = D3dParam.BackBufferWidth / s(1)
Pic_H = D3dParam.BackBufferHeight / s(1)

Form_Name = Me.Caption

For f = 0 To 4
    For j = 0 To 4
        XP_Active(f) = XP_Active(f) & IIf(f = j, "■", "□")
    Next
Next

Me.Width = Screen.Width
Me.Height = Screen.Height
Me.Show

End Sub
Private Sub Dx_Texture() '載入材質
Dim a(1) As Single, b(1) As Single, f As Integer, t As Long, Time_S As Long, j As Integer

Call Central(a(0), a(1), b(0), b(1)) '發牌起始計算
t = GetTickCount
For f = s(2) To s(3)
    Time_S = (GetTickCount - t) \ 1000
    Call Pic_Load(f, Pic_W, Pic_H) '讀取圖片
    Me.Caption = Form_Name & "　" & f + 1 & "/" & s(3) + 1 & "　" & XP_Active(f / 3 Mod 5) & "　" & "耗時 " & Time_S \ 60 & "分" & Time_S Mod 60 & "秒" & "　" & "共" & ma + 1 & "張"
    
    Vertex(0, f) = Ver(a(0), b(0), 0, 0) '起始位置
    Vertex(1, f) = Ver(a(1), b(0), 1, 0)
    Vertex(2, f) = Ver(a(0), b(1), 0, 1)
    Vertex(3, f) = Ver(a(1), b(1), 1, 1)

    j = Deal(f) '目地位置計算
    P_Size(f).D_Left = Pic_W * (j Mod s(1)) '目地位置
    P_Size(f).D_Top = Pic_H * ((f \ s(1)) Mod s(1))
    P_Size(f).D_Width = Pic_W
    P_Size(f).D_Height = Pic_H
    P_Size(f).D_XCen = (P_Size(f).D_Left + P_Size(f).D_Width) / 2
    P_Size(f).D_YCen = (P_Size(f).D_Top + P_Size(f).D_Height) / 2
    
    AF(f) = f '優先順序
    Call Swap(f)
    
    DoEvents
    If s(6) = 1 Or s(4) <> 0 Or s(7) = 1 Then Exit Sub '如果程式將要結束則離開迴圈
    Call Render '秀圖
    Call Pic_Move(f) '只移動目前頁面的圖
Next

Me.Caption = Form_Name & "　" & f + 1 - 1 & "/" & s(3) + 1 & "　" & "耗時 " & Time_S \ 60 & "分" & Time_S Mod 60 & "秒" & "　" & "共" & ma + 1 & "張"

End Sub
Private Sub Pic_Load(f, w As Single, h As Single)
On Error GoTo Err:
Dim FileName As String

FileName = File1.Path & "\" & File1.List(f)
Set Texture(f) = D3DX.CreateTextureFromFileEx(D3dDevice, FileName, w, h, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, File_ele, ByVal 0)
Exit Sub

Err:
    MsgBox "錯誤的圖片!", 64, "訊息"
End Sub
Private Sub Core() '核心程式 ☆
Do
    DoEvents
    If s(6) = 1 Or s(7) = 1 Then Exit Sub
    If s(4) <> 0 Then Call Page '如果按了上下頁則
    Call Pic_Move(s(3)) '移動
    Call Render '秀圖
    Sleep (20)
Loop
End Sub
Private Sub Page() '換頁
D3dDevice.SetTexture 0, Nothing
Erase Texture(), P_Size(), AF()
ReDim Texture(ma), P_Size(ma), AF(ma)
Vertex(0, 0).X = -Pic_W
Call Vertex_P(0, -Pic_W, -Pic_H, Pic_W, Pic_H)

s(5) = s(5) + s(4) '累加頁數
s(2) = s(1) ^ 2 * s(5) '迴圈起始
s(3) = (s(1) ^ 2) * (s(5) + 1) - 1 '迴圈終止
s(3) = IIf(s(3) > ma, ma, s(3)) '預防超出最大圖片張數
s(4) = 0 '清除
Call Dx_Texture
End Sub
Sub Render() '秀圖
Dim f As Integer, color As Long
With D3dDevice
    .Clear 0, ByVal 0, D3DCLEAR_TARGET, color, 18, 0
    .BeginScene
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
            For f = s(3) To s(2) Step -1 '繪出
                .SetTexture 0, Texture(AF(f))
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, AF(f)), Len(Vertex(0, AF(f)))
            Next
        .SetRenderState D3DRS_ALPHABLENDENABLE, False
    .EndScene
    
    .Present ByVal 0, ByVal 0, 0, ByVal 0
End With
End Sub
Sub Central(a0 As Single, a1 As Single, b0 As Single, b1 As Single) '發牌起始計算
a0 = D3dParam.BackBufferWidth / s(1) * (s(1) - 1) - Pic_W  '計算位置
b0 = D3dParam.BackBufferHeight / s(1) * (s(1) - 2) - Pic_H
a1 = D3dParam.BackBufferWidth / s(1) * (s(1) - 1) + Pic_W
b1 = D3dParam.BackBufferHeight / s(1) * (s(1) - 2) + Pic_H
End Sub
Function Deal(f) As Integer '發牌目地位置計算
Deal = IIf((f \ s(1)) Mod 2 = 0, f, s(1) - (f + 1) Mod s(1))
End Function
Private Function Pic_coll(Ax As Single, Ay As Single, f As Integer, Bx As Single, By As Single) As Boolean
If Ax > Bx And Ax < Bx + P_Size(f).D_Width And Ay > By And Ay < By + P_Size(f).D_Height Then Pic_coll = True
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '滑鼠按下
Dim f As Integer, a As Single, b As Single, j As Integer
P_Move(1) = X / G_x
P_Move(2) = Y / G_Y
For f = s(2) To s(3)
    If Pic_coll(P_Move(1), P_Move(2), AF(f), Vertex(0, AF(f)).X, Vertex(0, AF(f)).Y) Then
        P_Move(3) = AF(f)
        old_Vertex_x = P_Move(1) - Vertex(0, AF(f)).X
        old_Vertex_y = P_Move(2) - Vertex(0, AF(f)).Y
        Call Swap(AF(f)) '優先順序交換
        Exit For
    End If
Next
If P_Size(P_Move(3)).A_Turn < 270 Then Exit Sub '如果這張圖還沒轉完則離開
Select Case Button
    Case 1 '左鍵
        P_Move(0) = 1 '按住左鍵
        If f = s(3) + 1 Then
            P_Move(0) = 0 '取消左鍵
            For f = s(2) To s(3)
                If P_Size(f).A_Turn = 270 Then '如果已經轉完則執行
                    P_Size(f).A_Move = 0
                    P_Move(3) = f: Call Form_MouseDown(2, 1, X, Y) '如果是放大的狀態則縮小
                End If
            Next
        End If
    Case 2 '右鍵
        If f <> s(3) + 1 Or Shift = 1 Then
            P_Size(P_Move(3)).A_Move = 0 '清除已移動次數(歸位)
            
            If P_Size(P_Move(3)).Large Then Call Form_MouseDown(4, 1, X, Y) '如果是放大的狀態則縮小
            
            j = Deal(P_Move(3)) '目地位置計算
            P_Size(P_Move(3)).D_Left = Pic_W * (j Mod s(1)) '目地位置
            P_Size(P_Move(3)).D_Top = Pic_H * ((P_Move(3) \ s(1)) Mod s(1))
        End If
    Case 4 '中鍵
        If f <> s(3) + 1 Or Shift = 1 Then
            P_Size(P_Move(3)).A_Move = 0 '清除已移動次數(歸位)
            
            P_Size(P_Move(3)).Large = Not P_Size(P_Move(3)).Large '縮放

            If P_Size(P_Move(3)).Large Then
                Call Pic_Load(P_Move(3), 640, 480) 'D3DX_DEFAULT, D3DX_DEFAULT)
                
                a = IIf(File_ele.Width < File_ele.Height, 480, 640) '直式橫式判斷
                b = IIf(File_ele.Width < File_ele.Height, 640, 480)
                
                P_Size(P_Move(3)).D_Width = a 'File_ele.Width
                P_Size(P_Move(3)).D_Height = b 'File_ele.Height
            Else
                Call Pic_Load(P_Move(3), Pic_W, Pic_H)
                P_Size(P_Move(3)).D_Width = Pic_W
                P_Size(P_Move(3)).D_Height = Pic_H
            End If
            P_Size(P_Move(3)).D_Left = Vertex(0, P_Move(3)).X + (Vertex(1, P_Move(3)).X - Vertex(0, P_Move(3)).X) / 2 - P_Size(P_Move(3)).D_Width / 2
            P_Size(P_Move(3)).D_Top = Vertex(0, P_Move(3)).Y + (Vertex(2, P_Move(3)).Y - Vertex(0, P_Move(3)).Y) / 2 - P_Size(P_Move(3)).D_Height / 2
        End If
End Select
End Sub
Private Sub Swap(ByVal fx As Integer) '優先順序交換演算法ＱＱ (想了很久終於想到這個解決方案)
Dim a As Integer, t As Integer, f As Integer
a = fx
For f = s(2) To s(3)
    t = AF(f)
    AF(f) = a
    If t = fx Then Exit For
    a = t
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Single, b As Single
If P_Move(0) = 1 Then
    X = X / G_x
    Y = Y / G_Y
    Vertex(0, P_Move(3)).X = P_Move(1) + X - P_Move(1) - old_Vertex_x
    Vertex(0, P_Move(3)).Y = P_Move(2) + Y - P_Move(2) - old_Vertex_y
    Call Vertex_P(P_Move(3), Vertex(0, P_Move(3)).X, Vertex(0, P_Move(3)).Y, P_Size(P_Move(3)).D_Width, P_Size(P_Move(3)).D_Height)
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then P_Move(0) = 0 '放開左鍵
End Sub
Private Sub Vertex_P(ByVal pps As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    Vertex(1, pps).X = X + Width
    Vertex(1, pps).Y = Y
    Vertex(2, pps).X = X
    Vertex(2, pps).Y = Y + Height
    Vertex(3, pps).X = X + Width
    Vertex(3, pps).Y = Y + Height
End Sub
Private Sub Pic_Move(ByVal j As Integer) '圖片自動移動
Dim f As Integer
Dim X(1) As Single, w As Single
Dim Y(1) As Single, h As Single

j = IIf(j > s(3), s(3), j) '限制範圍為本頁
For f = s(2) To j
    If P_Size(f).A_Move < s(0) Then
        If P_Size(f).A_Move = 0 Then '如果此元件還沒有移動過則
            P_Size(f).Left = Vertex(0, f).X '來源
            P_Size(f).Top = Vertex(0, f).Y
            P_Size(f).Width = Vertex(1, f).X - Vertex(0, f).X
            P_Size(f).Height = Vertex(2, f).Y - Vertex(0, f).Y
        End If
        
        w = Vertex(1, f).X - Vertex(0, f).X
        h = Vertex(2, f).Y - Vertex(0, f).Y

        P_Size(f).A_Move = P_Size(f).A_Move + 1
        
        Vertex(0, f).X = Vertex(0, f).X + (P_Size(f).D_Left - P_Size(f).Left) / s(0)
        Vertex(0, f).Y = Vertex(0, f).Y + (P_Size(f).D_Top - P_Size(f).Top) / s(0)
        w = w + (P_Size(f).D_Width - P_Size(f).Width) / s(0)
        h = h + (P_Size(f).D_Height - P_Size(f).Height) / s(0)
        
        Xcenter(f) = Xcenter(f) + (P_Size(f).D_XCen - P_Size(f).XCen) / s(0)
        Ycenter(f) = Ycenter(f) + (P_Size(f).D_YCen - P_Size(f).YCen) / s(0)
        
        Call Vertex_P(f, Vertex(0, f).X, Vertex(0, f).Y, w, h)
        DoEvents
    Else
        Call Turn(f)
    End If
Next

End Sub
Sub Turn(L As Integer) '旋轉核心
Dim a(3) As Single, b(3) As Single, f As Integer, j As Integer, i As Integer

For f = s(2) To L
    If P_Size(f).A_Turn < 270 Then
        If P_Size(f).A_Turn = 0 Then '頂點基準點
            P_Size(f).XCen = (Vertex(0, f).X + Vertex(1, f).X) / 2 '原基準點
            P_Size(f).YCen = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
            Xcenter(f) = P_Size(f).XCen
            Ycenter(f) = P_Size(f).YCen
            
            Xdis(0, f) = Vertex(0, f).X - Xcenter(f) '距離
            Ydis(0, f) = Vertex(0, f).Y - Ycenter(f) '
            Xdis(1, f) = Vertex(1, f).X - Xcenter(f) '距離
            Ydis(1, f) = Vertex(1, f).Y - Ycenter(f) '
            Xdis(2, f) = Vertex(2, f).X - Xcenter(f) '距離
            Ydis(2, f) = Vertex(2, f).Y - Ycenter(f) '
            Xdis(3, f) = Vertex(3, f).X - Xcenter(f) '距離
            Ydis(3, f) = Vertex(3, f).Y - Ycenter(f)
        End If
    
        P_Size(f).A_Turn = (P_Size(f).A_Turn + 1) Mod 361
        i = P_Size(f).A_Turn
        
        For j = 0 To 3
            a(j) = Xcenter(f) + Xdis(j, f) / 2 + Cosine(i) - Sine(i) * Xdis(j, f) / 2
            b(j) = Ycenter(f) + Ydis(j, f) + Sine(i) + Cosine(i) * Xdis(j, f)
        Next
        Vertex(0, f) = Ver(a(0), b(0), 0, 0)
        Vertex(1, f) = Ver(a(1), b(1), 1, 0)
        Vertex(2, f) = Ver(a(2), b(2), 0, 1)
        Vertex(3, f) = Ver(a(3), b(3), 1, 1)
        
        DoEvents
    End If
Next
End Sub
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
    Case 38 '上
        If s(5) > 0 Then s(4) = -1
    Case 40 '下
        If s(3) < ma Then s(4) = 1
    Case 27 'ESC
        Unload Me
End Select
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    P_Move(4) = 1
    old_X = X \ 15
    old_Y = Y \ 15
End If
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If P_Move(4) = 1 Then
    X = X \ 15
    Y = Y \ 15
    Frame1.Left = Frame1.Left + X - old_X
    Frame1.Top = Frame1.Top + Y - old_Y
End If
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then P_Move(4) = 0
End Sub
Private Sub Pic_Drives_Click(Index As Integer) '功能表列
Select Case Index
    Case 0
        Frame1.Visible = True
    Case 1 '上一頁
        Call Form_KeyDown(38, 0)
    Case 2 '下一頁
        Call Form_KeyDown(40, 0)
End Select
End Sub
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then s(7) = 1
Frame1.Visible = False
End Sub
Private Sub Drive1_Change()
On Error GoTo Err:
Dir1.Path = Drive1.Drive
Exit Sub

Err:
    MsgBox "沒有磁片!", 48, "警告!"
End Sub
Private Sub Form_Resize()
G_x = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
End Sub
Private Sub Form_Unload(Cancel As Integer) '表單移除
If s(6) < 2 Then '即將結束
    s(6) = 1
    Cancel = 1
Else '結束
    Call UnDX
End If
End Sub
