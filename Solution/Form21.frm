VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form21 
   Appearance      =   0  '����
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
   StartUpPosition =   2  '�ù�����
   WindowState     =   2  '�̤j��
   Begin MSComDlg.CommonDialog openfileDialog 
      Left            =   960
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '�S���ؽu
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
         Caption         =   "�̭�Ϫ��ѪR��"
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
         BackStyle       =   0  '�z��
         Caption         =   "�վ�Ϥ��ѪR��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
      Caption         =   "��ܹϤ���Ƨ�"
      Index           =   0
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "�վ�Ϥ��ѪR��"
      Index           =   1
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "�W�@��(��)"
      Index           =   2
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "�U�@��(��)"
      Index           =   3
   End
   Begin VB.Menu Pic_Drives 
      Caption         =   "����(H)"
      Index           =   4
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�@��:�p�� lbt95@yahoo.com.tw
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim Texture() As Direct3DTexture8

Dim Pic_Back As String
Const PR As Byte = 2 'ma+1)�I�� ma+2) �B�n
'Const Pic_Drive As String = ""

Dim Cosine(360) As Single
Dim Sine(360) As Single
Dim Xdis() As Single, Ydis() As Single
Dim O_Xdis() As Single, O_Ydis() As Single
Dim D_Xdis() As Single, D_Ydis() As Single
Dim fileArray() As String '�Ϥ����|�t�W�ٰ}�C
Dim ma As Integer '�Ϥ��ƶq
Dim S(11) As Integer '0)g�Ϥ����ʺ�ӫ� 1)�C�ƹϼ� 2)�Ϥ��j��_�l 3)�Ϥ��j����I 4)�W�U�� 5)���Ʋ֥[ 6)�{���Y�N���� 7)�ƹ��S������󪺹� 8)�üƭI�� 9)���DX��r 10)��ܹϤ���Ƨ� 11)�}�Y�ʵe
Dim Pic_W As Single '�ϼe
Dim Pic_H As Single '�ϰ�
Dim PiX As Single 'X�ѪR��
Dim PiY As Single 'Y�ѪR��
Dim old_X As Single '���ʼ����������ȦsX
Dim old_Y As Single

Dim old_Vertex_x As Single
Dim old_Vertex_y As Single
Dim P_Move(3) As Single '0)����ƹ����� 1)�ƹ��I�U�h��X��m 2)�ƹ��I�U�h��Y��m 3)�ƹ��襤���Ϥ�
Dim G_X As Single '������
Dim G_Y As Single '������
Dim AF() As Integer '�u�����Ǹ��X(�ߤ@�@��)
Dim File_ele As D3DXIMAGE_INFO '�ɮ׸�T
Dim Bermuda '�ʼ}�F�T���w.........................�u�n�b����m�ŧi���w���~
Dim DXWord As String 'DX��r
Dim Form_Name As String '����l�W��
Dim XP_Active(4) As String '���D�C�ʵe

Dim P_Size() As Pic_DX
Private Type Pic_DX
    A_Move As Byte '�Ϥ��w���ʦ��ơ]�k��)
    A_Turn As Integer '�Ϥ������ਤ��
    OpenTurn As Byte '����}��
    Xcenter As Single
    Ycenter As Single
    O_XCen As Single '�_�l����
    O_YCen As Single '�_�l����
    D_XCen As Single '�ت�����
    D_YCen As Single '�ت�����
    Large As Boolean '�Y��
    Alpha As Byte
    Dis_Large As Byte '�����\�Y��
    Perfect As Byte '�ӽo�{��
End Type
Private Sub Form_Load() '�����J
Call Option1_Click(1) '�w�]��800 X 600
Call Three '�G���
Call Load_START '�uŪ�@��
Call Form_KeyDown(72, 0) '��ܻ���
Call Dx_START(True, Me) 'DX��l
Call DXFont 'DX��r
End Sub
Private Sub Form_Activate()
Do
    Call START(openfileDialog.fileName) '��l��
    Call BackPicture(Pic_Back) '�I��
    Call Dx_Texture   'DX����
    Call Core '�֤�
    If S(10) = 1 Then Me.Caption = Form_Name: Call Ma_Clear
    If S(6) = 1 Then S(6) = 2: Unload Me '�����{��(ESC)
Loop Until S(6) <> 0
Call Ma_Clear '�M���ƭȸ��
End Sub
Private Sub START(fileName As String) '��l�� a)�w�]�����~���е{�����|
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

S(0) = 10 '(���J���賡���B�k��)'�Ϥ����ʺ�ӫ�
S(1) = 10 '�C�C�@10�i 10 X 10 = 100�i(��)
S(2) = 0 '-----------�Ϥ��j���l��
S(3) = IIf(ma < S(1) ^ 2 - 1, ma, S(1) ^ 2 - 1) '�Ϥ��j��פ��s

G_X = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
Pic_W = D3dParam.BackBufferWidth / S(1)
Pic_H = D3dParam.BackBufferHeight / S(1)
End Sub
Private Sub Dx_Texture() '���J����
Dim a(1) As Single, b(1) As Single, f As Integer, j As Integer, i As Byte, t As Long

If UBound(fileArray) = -1 Then Exit Sub

t = GetTickCount '�p�ɶ}�l
Call Central(a(0), a(1), b(0), b(1)) '�o�P�_�l�p��
For f = S(2) To S(3)
    Set Texture(f) = LoadTexture(fileArray(f), Pic_W, Pic_H) 'Ū���Ϥ�
    Me.Caption = Form_Name & "�@" & f + 1 & " / " & S(3) + 1 & "�@�Ϥ����J���@" & XP_Active(f \ 3 Mod 5) & "�@�@" & ma + 1 & "�i�@" & "�Ӯ�" & (GetTickCount - t) \ 1000 & "��"
    
    Vertex(0, f) = Ver(a(0), b(0), 0, 0) '�_�l��m
    Vertex(1, f) = Ver(a(1), b(0), 1, 0)
    Vertex(2, f) = Ver(a(0), b(1), 0, 1)
    Vertex(3, f) = Ver(a(1), b(1), 1, 1)

    j = Deal(f) '�ئa��m�p��
    With P_Size(f)
        .Xcenter = (Vertex(0, f).X + Vertex(1, f).X) / 2 '����
        .Ycenter = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
        For i = 0 To 3
            Call Auto_Dis(f, Pic_W, Pic_H)
        Next
        .D_XCen = D_Xdis(3, f) + Pic_W * (j Mod S(1)) '�ت�����
        .D_YCen = D_Ydis(3, f) + Pic_H * (f \ S(1) Mod S(1))
        .Perfect = S(0) '���ʲӽo��
        .Alpha = 255
        .A_Turn = 270
        .OpenTurn = 2
    End With
    AF(f) = f '�u������
    Call Swap(f)
    DoEvents
    
    If S(6) = 1 Or S(4) <> 0 Or S(10) = 1 Then Exit Sub '�p�G�{���N�n�����h���}�j��
    Call Render(f) '�q��
    Call Pic_Move(f) '�u���ʥثe��������
Next
Me.Caption = Form_Name & "�@" & f & " / " & S(3) + 1 & "�@�@" & ma + 1 & "�i�@" & "�Ӯ�" & (GetTickCount - t) \ 1000 & "��"
Do Until P_Size(S(3)).A_Move = P_Size(S(3)).Perfect '����ʵe����
    DoEvents
    If S(6) = 1 Or S(4) <> 0 Or S(10) = 1 Then Exit Sub '�p�G�{���N�n�����h���}�j��
    Call Render(S(3)) '�q��
    Call Pic_Move(S(3)) '�u���ʥثe��������
    Sleep (20)
Loop
S(11) = 1 '�}�Y�ʵe�w����
End Sub
Private Sub Core() '�֤ߵ{�� ��
Do
    DoEvents
    If S(4) <> 0 Then Call Page '�p�G���F�W�U���h
    If S(10) = 1 Then Exit Sub '�p�G���F�Ϥ���Ƨ�
    Call Pic_Move(S(3)) '����
    Call Render(S(3)) '�q��
    Sleep (20)
Loop While S(6) = 0 '�����{��
End Sub
Private Function Pic_coll(Ax As Single, Ay As Single, f As Integer, Bx As Single, By As Single) As Boolean
If Ax > Bx And Ax < Bx + Xdis(3, f) * 2 And Ay > By And Ay < By + Ydis(3, f) * 2 Then Pic_coll = True
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '�ƹ����U
Dim f As Integer, j As Integer, w As Single, h As Single, tx As Single, ty As Single
S(9) = 0 '�M��DX��r
P_Move(1) = X / G_X
P_Move(2) = Y / G_Y
For f = S(2) To S(3)
    If Pic_coll(P_Move(1), P_Move(2), AF(f), Vertex(0, AF(f)).X, Vertex(0, AF(f)).Y) Then
        P_Move(3) = AF(f)
        old_Vertex_x = P_Size(AF(f)).Xcenter - P_Move(1)
        old_Vertex_y = P_Size(AF(f)).Ycenter - P_Move(2)
        Call Swap(AF(f)) '�u�����ǥ洫
        Exit For
    End If
Next
S(7) = IIf(f <> S(3) + 1, 0, 1) '�O�_�S���I��Ϥ�

Select Case Button
    Case 1 '����
        P_Move(0) = 1 '������
        If f = S(3) + 1 Then
            P_Move(0) = 0 '��������
            For f = S(2) To S(3)
                P_Move(3) = f
                Call Form_MouseDown(2, 1, X, Y) '�p�G�O��j�����A�h�Y�p
            Next
        End If
    Case 2 '�k��
        If f <> S(3) + 1 Or Shift = 1 Then
            With P_Size(P_Move(3))
                If .Large Then Call Form_MouseDown(3, 1, X, Y) '�p�G�O��j�����A�h�Y�p
                
                tx = .D_XCen
                ty = .D_YCen
                j = Deal(P_Move(3)) '�ئa��m�p��
                .D_XCen = D_Xdis(3, P_Move(3)) + Pic_W * (j Mod S(1)) '�ئa��m
                .D_YCen = D_Ydis(3, P_Move(3)) + Pic_H * ((P_Move(3) \ S(1)) Mod S(1))
                
                If tx = .D_XCen And ty = .D_YCen Then Exit Sub
                
                .Perfect = 10 '�ӽo�׬�20
                .A_Move = 0
            End With
        End If
    Case 3 '��j
        If f <> S(3) + 1 Or Shift = 1 Then
            With P_Size(P_Move(3))
                .Perfect = 10 '�ӽo�׬�20
                .A_Move = 0
                .Large = Not .Large '�Y��
                .Alpha = 100 '�z����
                Debug.Print .A_Turn
                If .Large Then
                    If PiX = -1 Then '�p�G�O��͸ѪR�שάO�P�Ǥ@�y�ܫh
                        Set Texture(P_Move(3)) = LoadTexture(fileArray(P_Move(3)))
                        w = File_ele.Width
                        h = File_ele.Height
                    Else
                        Set Texture(P_Move(3)) = LoadTexture(fileArray(P_Move(3)), PiX, PiY)
                        w = IIf(File_ele.Width < File_ele.Height, PiY, PiX) '������P�_
                        h = IIf(File_ele.Width < File_ele.Height, PiX, PiY)
                    End If
                Else
                    '.OpenTurn = 1 '�}�ұ���
                    Set Texture(P_Move(3)) = LoadTexture(fileArray(P_Move(3)), Pic_W, Pic_H)
                    w = Pic_W '�e
                    h = Pic_H '��
                End If
                Call Auto_Dis(P_Move(3), w, h)
            End With
        End If
End Select
End Sub
Private Sub Swap(ByVal fx As Integer) '�u�����ǥ洫�t��k�ߢ� (�Q�F�ܤ[�ש�Q��o�ӸѨM���)
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
        .Dis_Large = 1 '�]���ϥΪ̭n���ʩҥH���R�\�Y��
    End With
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If S(7) = 1 Then Exit Sub

If Button = 1 Then
    P_Move(0) = 0 '��}����
    If P_Size(P_Move(3)).Dis_Large = 0 Then
        ' ����@�}�l���}���ɮ׮��ܦ��L�ɮש�j
        If UBound(fileArray) >= 0 Then
            Call Form_MouseDown(3, 1, X, Y) '�p�G�ϥΪ̥ت��O�n�Y��h�Y��
        End If
    End If
    P_Size(P_Move(3)).Dis_Large = 0 '�M������
End If
End Sub
Private Sub Pic_Move(ByVal j As Integer) '�Ϥ��۰ʲ���
Dim f As Integer, c As Long, i As Byte
Dim X(1) As Single, w As Single
Dim Y(1) As Single, h As Single
Dim a(3) As Single, b(3) As Single, old_Color As Long, d

j = IIf(j > S(3), S(3), j) '����d�򬰥���
For f = S(2) To j
    If P_Size(f).A_Move < P_Size(f).Perfect Then
        If P_Size(f).A_Move = 0 Then '��l��m
            With P_Size(f) '�줤��
                .Xcenter = (Vertex(0, f).X + Vertex(1, f).X) / 2
                .Ycenter = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
                .O_XCen = .Xcenter
                .O_YCen = .Ycenter
            End With
            For i = 0 To 3 '�Z��
                Xdis(i, f) = Vertex(i, f).X - P_Size(f).O_XCen
                Ydis(i, f) = Vertex(i, f).Y - P_Size(f).O_YCen
                O_Xdis(i, f) = Xdis(i, f)
                O_Ydis(i, f) = Ydis(i, f)
            Next
        End If
        With P_Size(f) '���ʮ֤�
            .A_Move = IIf(.OpenTurn = 0, .A_Move + 1, .A_Move + 1)
            .Xcenter = .Xcenter + (.D_XCen - .O_XCen) / .Perfect '����
            .Ycenter = .Ycenter + (.D_YCen - .O_YCen) / .Perfect
        End With
        For i = 0 To 3 '�Y��֤�
            Xdis(i, f) = Xdis(i, f) + (D_Xdis(i, f) - O_Xdis(i, f)) / P_Size(f).Perfect '�Z��
            Ydis(i, f) = Ydis(i, f) + (D_Ydis(i, f) - O_Ydis(i, f)) / P_Size(f).Perfect
        Next
        If S(11) = 0 And P_Size(f).A_Move = P_Size(f).Perfect And P_Size(f).OpenTurn = 2 Then P_Size(f).OpenTurn = 1 '�}�Y�ʵe
    End If
    
    With P_Size(f)
        If .Alpha < 251 Then '�z����
            .Alpha = .Alpha + 5
            c = D3DColorARGB(.Alpha, 255, 255, 255)
        End If
        
        If .OpenTurn = 1 Then '����֤�
            .A_Turn = (.A_Turn + 17) Mod 361
            For i = 0 To 3 '�٭쳻�I��m
                a(i) = .Xcenter + Xdis(i, f) / 2 + Cosine(.A_Turn) - Sine(.A_Turn) * Xdis(i, f) / 2 '���I
                b(i) = .Ycenter + Ydis(i, f) + Sine(.A_Turn) + Cosine(.A_Turn) * Xdis(i, f)
            Next
            If .A_Turn <= 270 + 16 And .A_Turn >= 270 - 16 Then .OpenTurn = 0: .A_Turn = 270 '�p�G�ۨ쥿���h��������
        Else '������
            For i = 0 To 3 '�٭쳻�I��m
                a(i) = .Xcenter + Xdis(i, f)
                b(i) = .Ycenter + Ydis(i, f)
            Next
        End If
    End With

    old_Color = Vertex(0, f).color '�s�J�ª��C��
    Vertex(0, f) = Ver(a(0), b(0), 0, 0)
    Vertex(1, f) = Ver(a(1), b(1), 1, 0)
    Vertex(2, f) = Ver(a(2), b(2), 0, 1)
    Vertex(3, f) = Ver(a(3), b(3), 1, 1)
    d = IIf(c <> 0, c, old_Color) '�٭��Ӫ��C��
    For i = 0 To 3
        Vertex(i, f).color = d
    Next
    c = 0
    DoEvents
Next
End Sub
Sub Render(j As Integer) '�q��
Dim f As Integer
On Error GoTo Err
With D3ddevice
    .BeginScene
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
            .SetTexture 0, Texture(ma + 1) '�I��
            .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, ma + 1), Len(Vertex(0, ma + 1))
            For f = j To S(2) Step -1 'ø�X'�Ϥ�
                .SetTexture 0, Texture(AF(f))
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, AF(f)), Len(Vertex(0, AF(f)))
            Next
            
            If S(9) <> 0 Then
                If S(9) = 2 Then .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, ma + 2), Len(Vertex(0, ma + 2))
                D3DX.DrawText MainFont, &HFFFFFFFF, DXWord, TextRect, DT_TOP
            End If
        .SetRenderState D3DRS_ALPHABLENDENABLE, False
    .EndScene
    .Present ByVal 0, ByVal 0, 0, ByVal 0 'Flip����
End With
Exit Sub
Err:
Unload Me
End Sub
Private Sub Load_START() '�uŪ�@��
Dim f As Byte, j As Byte

'�䴩�h�ɶ}��
openfileDialog.flags = cdlOFNAllowMultiselect + cdlOFNExplorer

'�üƭI��
Randomize
S(8) = Int(Rnd * 3) + 1
S(8) = 63 + S(8) * 2 '=65 67 69
Pic_Back = App.Path & "\BackPicture\" & S(8) & ".jpg"

'����������
'����������
'����������
'����������
'����������

For f = 0 To 4
    For j = 0 To 4
        XP_Active(f) = XP_Active(f) & IIf(f = j, "��", "��")
    Next
Next
Form_Name = Me.Caption
End Sub
Function LoadTexture(ByVal fileName, Optional w As Single = D3DX_DEFAULT, Optional h As Single = D3DX_DEFAULT, Optional color As Long = 0) As Direct3DTexture8
On Error Resume Next

Set LoadTexture = D3DX.CreateTextureFromFileEx(D3ddevice, fileName, w, h, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, color, File_ele, ByVal 0)
End Function
Private Sub Auto_Dis(f, w As Single, h As Single) '�Z���۰ʤ�
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
Private Sub Page() '����
Dim a As Integer, t1 As Integer, t2 As Integer '�ȦsS2,S3

S(9) = 1 '��ܲĴX��
S(5) = S(5) + S(4) '�֥[����
t1 = S(1) ^ 2 * S(5) '�j��_�l
t2 = (S(1) ^ 2) * (S(5) + 1) - 1 '�j��פ�
t2 = IIf(t2 > ma, ma, t2) '�w���W�X�̤j�Ϥ��i��
S(4) = 0 '�M��
S(11) = 0 '�M���w���J�}�Y�ʵe
a = (ma + 1) \ S(1) ^ 2
If (ma + 1) Mod S(1) ^ 2 <> 0 Then a = a + 1
DXWord = "��" & S(5) + 1 & "�����@" & a & "���@" & t1 + 1 & "��" & t2 + 1

Call Glide(255) '�H�X

S(2) = t1
S(3) = t2

Vertex(1, 0).X = -Pic_W '���ò�0�i��
Vertex(3, 0).X = -Pic_W

D3ddevice.SetTexture 0, Nothing '�M���P����

Erase Texture(), P_Size(), AF()
ReDim Texture(ma + PR), P_Size(ma), AF(ma)

'���I��
S(8) = (S(8) - 65) / 2 '0 2 4 => 0 1 2
S(8) = (S(8) + 1) Mod 3 '���U�@�i�I��
S(8) = S(8) * 2 + 65
Call BackPicture(App.Path & "\BackPicture\" & S(8) & ".jpg") '���J�I��

Call Glide(0, 1) '�H�J
S(9) = 0

Call Dx_Texture
End Sub
Private Sub Glide(i As Byte, Optional b As Integer = -1) '�����e��
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
Private Sub BackPicture(fileName As String) '�I��
Dim j As Byte

Set Texture(ma + 1) = LoadTexture(fileName)
Vertex(0, ma + 1) = Ver(0, 0, 0, 0)
Vertex(1, ma + 1) = Ver((D3dParam.BackBufferWidth), 0, 1, 0)
Vertex(2, ma + 1) = Ver(0, (D3dParam.BackBufferHeight), 0, 1)
Vertex(3, ma + 1) = Ver((D3dParam.BackBufferWidth), (D3dParam.BackBufferHeight), 1, 1)

'�B�n
Vertex(0, ma + 2) = Ver(0, 0, 0, 0)
Vertex(1, ma + 2) = Ver((D3dParam.BackBufferWidth), 0, 1, 0)
Vertex(2, ma + 2) = Ver(0, (D3dParam.BackBufferHeight), 0, 1)
Vertex(3, ma + 2) = Ver((D3dParam.BackBufferWidth), (D3dParam.BackBufferHeight), 1, 1)
For j = 0 To 3
    Vertex(j, ma + 2).color = &HAA000000
Next
End Sub
Sub Central(a0 As Single, a1 As Single, b0 As Single, b1 As Single) '�o�P�_�l�p��
With D3dParam
    a0 = .BackBufferWidth - Pic_W * 2 '�p���m
    b0 = .BackBufferHeight - Pic_H * 3
    a1 = .BackBufferWidth
    b1 = .BackBufferHeight - Pic_H
End With
End Sub
Function Deal(f) As Integer '�o�P�ئa��m�p��
Deal = IIf((f \ S(1)) Mod 2 = 0, f, S(1) - (f + 1) Mod S(1))
End Function
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
    Case 38 'UP
        If S(5) > 0 Then S(4) = -1
    Case 40 'Down
        If S(3) >= ma Then S(5) = -1
        S(4) = 1
    Case 72 'Help
        DXWord = "�ƹ��G" & vbCrLf & "�@�@����G������j" & vbCrLf & "�@�@�k��G�Ϥ��k��" & vbCrLf & "��L�G" & vbCrLf & "�@�@���G�W�@��" & vbCrLf & "�@�@���G�U�@��" & vbCrLf & "�@�@�֡G����"
        S(9) = (S(9) + 2) Mod 4
    Case 27 'ESC
        Unload Me
    Case 13
        If Shift = 4 Then
            Call Dx_START(False, Me)
            Call START(openfileDialog.fileName) '��l��
            Call BackPicture(Pic_Back) '�I��
            Call Dx_Texture 'DX����
        End If
Case Else
    S(9) = 0 '�����DX��r
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
Private Sub Pic_Drives_Click(Index As Integer) '�ɮץ\���
Select Case Index
    Case 0 '��ܹϤ���Ƨ�
        openfileDialog.CancelError = True
        openfileDialog.MaxFileSize = 10240
        
        On Error GoTo Err
        openfileDialog.fileName = ""
        openfileDialog.ShowOpen
        S(10) = 1 ' ���s��ܹϤ�
    Case 1 '��j�ѪR��
        Frame1(1).Visible = True
    Case 2 '�W�@��
        Call Form_KeyDown(38, 0)
    Case 3 '�U�@��
        Call Form_KeyDown(40, 0)
    Case 4 '����
        Call Form_KeyDown(72, 0)
End Select

Err:
End Sub
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(Index).Picture = Image3(1).Picture
End Sub
Private Sub Image1_Click(Index As Integer) '����X����
Frame1(Index).Visible = False
End Sub
Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(Index).Picture = Image3(0).Picture
End Sub
Private Sub Option1_Click(Index As Integer)
Dim a() As String
If Index = 6 Then PiX = -1: Exit Sub
a = Split(Option1(Index).Caption, " X ")
PiX = a(0) 'X�ѪR��
PiY = a(1) 'Y�ѪR��
End Sub
Private Sub Ma_Clear() '�M���ƭȸ��
Dim t As Byte '�I���ƭȼȦs
t = S(8)
Erase Texture(), Vertex(), P_Size(), AF(), P_Move(), Xdis(), Ydis(), O_Xdis(), O_Ydis(), D_Xdis(), D_Ydis(), S(), fileArray '�M�����(���Ϥ��ؿ�)
S(8) = t
End Sub
Private Sub Form_Unload(Cancel As Integer) '��沾��
If S(6) < 2 Then '�Y�N����
    S(6) = 1
    Cancel = 1
Else '����
    Call UnDX
End If
End Sub
