VERSION 5.00
Begin VB.Form Form21 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "DirectX 0.7"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  '����
   ScaleWidth      =   792
   StartUpPosition =   2  '�ù�����
   Begin VB.FileListBox File1 
      Height          =   1350
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ܹϤ��ؿ�"
      Height          =   4215
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "����"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�T�w"
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
      Caption         =   "�Ϥ��ؿ�"
      Index           =   0
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "�W�@��(��)"
      Index           =   1
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "�U�@��(��)"
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

Const Pic_Drive As String = "C:\Windows\Web\Wallpaper\Nature" '�w�]�Ϥ��ؿ�
Dim Cosine(360) As Single
Dim Sine(360) As Single
Dim Xcenter() As Single, Ycenter() As Single
Dim Xdis() As Single, Ydis() As Single
Dim ma As Integer '�Ϥ��ƶq
Dim Form_Name As String '���W��
Dim s(7) As Integer '0)g�Ϥ����ʺ�ӫ� 1)�C�ƹϼ� 2)�Ϥ��j��_�l 3)�Ϥ��j����I 4)�W�U�� 5)���Ʋ֥[ 6)�{���Y�N���� 7)���Ϥ��ؿ�
Dim Pic_W As Single '�ϼe
Dim Pic_H As Single '�ϰ�

Dim old_Vertex_x As Single
Dim old_Vertex_y As Single
Dim oSx As Single '�}�l���e
Dim oSy As Single '�}�l����
Dim P_Move(4) As Single '0)����ƹ����� 1)�ƹ��I�U�h��X��m 2)�ƹ��I�U�h��Y��m 3)�ƹ��襤���Ϥ� 4)Frame1 ���ƹ�����
Dim G_x As Single '������
Dim G_Y As Single '������
Dim AF() As Integer '�u�����Ǹ��X(�ߤ@�@��)
Dim File_ele As D3DXIMAGE_INFO '�ɮ׸�T
Dim old_X As Single '�ƹ��I�U�h��X(Frame1)
Dim old_Y As Single '�ƹ��I�U�h��Y
Dim XP_Active(4) As String '���D�ʵe

Dim P_Size() As Pic_DX
Private Type Pic_DX
    A_Move As Integer '�Ϥ��w���ʦ��ơ]�k��)
    A_Turn As Integer '�Ϥ��w���স��
    Left As Single '�ӷ�
    Top As Single
    Width As Single
    Height As Single
    XCen As Single
    YCen As Single
    D_Left As Single '�ئa
    D_Top As Single
    D_Width As Single
    D_Height As Single
    D_XCen As Single
    D_YCen As Single
    Large As Boolean '�Y��
End Type

Private Sub Form_Load() '�����J
Call Dx_START 'DX��l
Call Three '�G���

Dir1.Path = Pic_Drive '�Ϥ��ؿ�
Do
    Call START(Dir1.Path) '��l��
    Call Dx_Texture 'DX����
    Call Core '�֤�
    If s(6) = 1 Then s(6) = 2: Unload Me '�����{��(ESC)
    If s(7) = 1 Then Me.Caption = Form_Name: Erase Texture(), Vertex(), P_Size(), AF(), P_Move(), Xcenter(), Ycenter(), Xdis(), Ydis(), s(), XP_Active() '�M�����(���Ϥ��ؿ�)
Loop Until s(6) <> 0
End Sub
Private Sub START(a As String) '��l�� a)�w�]�����~���е{�����|
Dim f As Byte, j As Byte, b As String

Set D3DX = New D3DX8

File1.Path = a '�Ϥ����|
ma = File1.ListCount - 1
If ma = -1 Then ma = 0

ReDim Texture(ma) As Direct3DTexture8
ReDim Vertex(3, ma) As TLVERTEX
ReDim P_Size(ma) As Pic_DX
ReDim AF(ma) As Integer
ReDim Xcenter(ma) As Single, Ycenter(ma) As Single
ReDim Xdis(3, ma) As Single, Ydis(3, ma) As Single

s(0) = 10 '(���J���賡���B�k��)'�Ϥ����ʺ�ӫ�
s(1) = 10 '�C�C�@10�i 10 X 10 = 100�i(��)
s(2) = 0 '-----------�Ϥ��j���l��
s(3) = IIf(ma < s(1) ^ 2 - 1, ma, s(1) ^ 2 - 1) '�Ϥ��j��פ��

G_x = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
Pic_W = D3dParam.BackBufferWidth / s(1)
Pic_H = D3dParam.BackBufferHeight / s(1)

Form_Name = Me.Caption

For f = 0 To 4
    For j = 0 To 4
        XP_Active(f) = XP_Active(f) & IIf(f = j, "��", "��")
    Next
Next

Me.Width = Screen.Width
Me.Height = Screen.Height
Me.Show

End Sub
Private Sub Dx_Texture() '���J����
Dim a(1) As Single, b(1) As Single, f As Integer, t As Long, Time_S As Long, j As Integer

Call Central(a(0), a(1), b(0), b(1)) '�o�P�_�l�p��
t = GetTickCount
For f = s(2) To s(3)
    Time_S = (GetTickCount - t) \ 1000
    Call Pic_Load(f, Pic_W, Pic_H) 'Ū���Ϥ�
    Me.Caption = Form_Name & "�@" & f + 1 & "/" & s(3) + 1 & "�@" & XP_Active(f / 3 Mod 5) & "�@" & "�Ӯ� " & Time_S \ 60 & "��" & Time_S Mod 60 & "��" & "�@" & "�@" & ma + 1 & "�i"
    
    Vertex(0, f) = Ver(a(0), b(0), 0, 0) '�_�l��m
    Vertex(1, f) = Ver(a(1), b(0), 1, 0)
    Vertex(2, f) = Ver(a(0), b(1), 0, 1)
    Vertex(3, f) = Ver(a(1), b(1), 1, 1)

    j = Deal(f) '�ئa��m�p��
    P_Size(f).D_Left = Pic_W * (j Mod s(1)) '�ئa��m
    P_Size(f).D_Top = Pic_H * ((f \ s(1)) Mod s(1))
    P_Size(f).D_Width = Pic_W
    P_Size(f).D_Height = Pic_H
    P_Size(f).D_XCen = (P_Size(f).D_Left + P_Size(f).D_Width) / 2
    P_Size(f).D_YCen = (P_Size(f).D_Top + P_Size(f).D_Height) / 2
    
    AF(f) = f '�u������
    Call Swap(f)
    
    DoEvents
    If s(6) = 1 Or s(4) <> 0 Or s(7) = 1 Then Exit Sub '�p�G�{���N�n�����h���}�j��
    Call Render '�q��
    Call Pic_Move(f) '�u���ʥثe��������
Next

Me.Caption = Form_Name & "�@" & f + 1 - 1 & "/" & s(3) + 1 & "�@" & "�Ӯ� " & Time_S \ 60 & "��" & Time_S Mod 60 & "��" & "�@" & "�@" & ma + 1 & "�i"

End Sub
Private Sub Pic_Load(f, w As Single, h As Single)
On Error GoTo Err:
Dim FileName As String

FileName = File1.Path & "\" & File1.List(f)
Set Texture(f) = D3DX.CreateTextureFromFileEx(D3dDevice, FileName, w, h, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, File_ele, ByVal 0)
Exit Sub

Err:
    MsgBox "���~���Ϥ�!", 64, "�T��"
End Sub
Private Sub Core() '�֤ߵ{�� ��
Do
    DoEvents
    If s(6) = 1 Or s(7) = 1 Then Exit Sub
    If s(4) <> 0 Then Call Page '�p�G���F�W�U���h
    Call Pic_Move(s(3)) '����
    Call Render '�q��
    Sleep (20)
Loop
End Sub
Private Sub Page() '����
D3dDevice.SetTexture 0, Nothing
Erase Texture(), P_Size(), AF()
ReDim Texture(ma), P_Size(ma), AF(ma)
Vertex(0, 0).X = -Pic_W
Call Vertex_P(0, -Pic_W, -Pic_H, Pic_W, Pic_H)

s(5) = s(5) + s(4) '�֥[����
s(2) = s(1) ^ 2 * s(5) '�j��_�l
s(3) = (s(1) ^ 2) * (s(5) + 1) - 1 '�j��פ�
s(3) = IIf(s(3) > ma, ma, s(3)) '�w���W�X�̤j�Ϥ��i��
s(4) = 0 '�M��
Call Dx_Texture
End Sub
Sub Render() '�q��
Dim f As Integer, color As Long
With D3dDevice
    .Clear 0, ByVal 0, D3DCLEAR_TARGET, color, 18, 0
    .BeginScene
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
            For f = s(3) To s(2) Step -1 'ø�X
                .SetTexture 0, Texture(AF(f))
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, AF(f)), Len(Vertex(0, AF(f)))
            Next
        .SetRenderState D3DRS_ALPHABLENDENABLE, False
    .EndScene
    
    .Present ByVal 0, ByVal 0, 0, ByVal 0
End With
End Sub
Sub Central(a0 As Single, a1 As Single, b0 As Single, b1 As Single) '�o�P�_�l�p��
a0 = D3dParam.BackBufferWidth / s(1) * (s(1) - 1) - Pic_W  '�p���m
b0 = D3dParam.BackBufferHeight / s(1) * (s(1) - 2) - Pic_H
a1 = D3dParam.BackBufferWidth / s(1) * (s(1) - 1) + Pic_W
b1 = D3dParam.BackBufferHeight / s(1) * (s(1) - 2) + Pic_H
End Sub
Function Deal(f) As Integer '�o�P�ئa��m�p��
Deal = IIf((f \ s(1)) Mod 2 = 0, f, s(1) - (f + 1) Mod s(1))
End Function
Private Function Pic_coll(Ax As Single, Ay As Single, f As Integer, Bx As Single, By As Single) As Boolean
If Ax > Bx And Ax < Bx + P_Size(f).D_Width And Ay > By And Ay < By + P_Size(f).D_Height Then Pic_coll = True
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '�ƹ����U
Dim f As Integer, a As Single, b As Single, j As Integer
P_Move(1) = X / G_x
P_Move(2) = Y / G_Y
For f = s(2) To s(3)
    If Pic_coll(P_Move(1), P_Move(2), AF(f), Vertex(0, AF(f)).X, Vertex(0, AF(f)).Y) Then
        P_Move(3) = AF(f)
        old_Vertex_x = P_Move(1) - Vertex(0, AF(f)).X
        old_Vertex_y = P_Move(2) - Vertex(0, AF(f)).Y
        Call Swap(AF(f)) '�u�����ǥ洫
        Exit For
    End If
Next
If P_Size(P_Move(3)).A_Turn < 270 Then Exit Sub '�p�G�o�i���٨S�৹�h���}
Select Case Button
    Case 1 '����
        P_Move(0) = 1 '������
        If f = s(3) + 1 Then
            P_Move(0) = 0 '��������
            For f = s(2) To s(3)
                If P_Size(f).A_Turn = 270 Then '�p�G�w�g�৹�h����
                    P_Size(f).A_Move = 0
                    P_Move(3) = f: Call Form_MouseDown(2, 1, X, Y) '�p�G�O��j�����A�h�Y�p
                End If
            Next
        End If
    Case 2 '�k��
        If f <> s(3) + 1 Or Shift = 1 Then
            P_Size(P_Move(3)).A_Move = 0 '�M���w���ʦ���(�k��)
            
            If P_Size(P_Move(3)).Large Then Call Form_MouseDown(4, 1, X, Y) '�p�G�O��j�����A�h�Y�p
            
            j = Deal(P_Move(3)) '�ئa��m�p��
            P_Size(P_Move(3)).D_Left = Pic_W * (j Mod s(1)) '�ئa��m
            P_Size(P_Move(3)).D_Top = Pic_H * ((P_Move(3) \ s(1)) Mod s(1))
        End If
    Case 4 '����
        If f <> s(3) + 1 Or Shift = 1 Then
            P_Size(P_Move(3)).A_Move = 0 '�M���w���ʦ���(�k��)
            
            P_Size(P_Move(3)).Large = Not P_Size(P_Move(3)).Large '�Y��

            If P_Size(P_Move(3)).Large Then
                Call Pic_Load(P_Move(3), 640, 480) 'D3DX_DEFAULT, D3DX_DEFAULT)
                
                a = IIf(File_ele.Width < File_ele.Height, 480, 640) '������P�_
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
Private Sub Swap(ByVal fx As Integer) '�u�����ǥ洫�t��k�ߢ� (�Q�F�ܤ[�ש�Q��o�ӸѨM���)
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
If Button = 1 Then P_Move(0) = 0 '��}����
End Sub
Private Sub Vertex_P(ByVal pps As Integer, X As Single, Y As Single, Width As Single, Height As Single)
    Vertex(1, pps).X = X + Width
    Vertex(1, pps).Y = Y
    Vertex(2, pps).X = X
    Vertex(2, pps).Y = Y + Height
    Vertex(3, pps).X = X + Width
    Vertex(3, pps).Y = Y + Height
End Sub
Private Sub Pic_Move(ByVal j As Integer) '�Ϥ��۰ʲ���
Dim f As Integer
Dim X(1) As Single, w As Single
Dim Y(1) As Single, h As Single

j = IIf(j > s(3), s(3), j) '����d�򬰥���
For f = s(2) To j
    If P_Size(f).A_Move < s(0) Then
        If P_Size(f).A_Move = 0 Then '�p�G�������٨S�����ʹL�h
            P_Size(f).Left = Vertex(0, f).X '�ӷ�
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
Sub Turn(L As Integer) '����֤�
Dim a(3) As Single, b(3) As Single, f As Integer, j As Integer, i As Integer

For f = s(2) To L
    If P_Size(f).A_Turn < 270 Then
        If P_Size(f).A_Turn = 0 Then '���I����I
            P_Size(f).XCen = (Vertex(0, f).X + Vertex(1, f).X) / 2 '�����I
            P_Size(f).YCen = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
            Xcenter(f) = P_Size(f).XCen
            Ycenter(f) = P_Size(f).YCen
            
            Xdis(0, f) = Vertex(0, f).X - Xcenter(f) '�Z��
            Ydis(0, f) = Vertex(0, f).Y - Ycenter(f) '
            Xdis(1, f) = Vertex(1, f).X - Xcenter(f) '�Z��
            Ydis(1, f) = Vertex(1, f).Y - Ycenter(f) '
            Xdis(2, f) = Vertex(2, f).X - Xcenter(f) '�Z��
            Ydis(2, f) = Vertex(2, f).Y - Ycenter(f) '
            Xdis(3, f) = Vertex(3, f).X - Xcenter(f) '�Z��
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
Private Sub Three() '�G���
Dim f As Integer
Const PI As Single = 3.14159265358979
For f = 0 To 360
    Cosine(f) = Cos(f / 180 * PI)
    Sine(f) = Sin(f / 180 * PI)
Next
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 38 '�W
        If s(5) > 0 Then s(4) = -1
    Case 40 '�U
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
Private Sub Pic_Drives_Click(Index As Integer) '�\���C
Select Case Index
    Case 0
        Frame1.Visible = True
    Case 1 '�W�@��
        Call Form_KeyDown(38, 0)
    Case 2 '�U�@��
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
    MsgBox "�S���Ϥ�!", 48, "ĵ�i!"
End Sub
Private Sub Form_Resize()
G_x = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
End Sub
Private Sub Form_Unload(Cancel As Integer) '��沾��
If s(6) < 2 Then '�Y�N����
    s(6) = 1
    Cancel = 1
Else '����
    Call UnDX
End If
End Sub
