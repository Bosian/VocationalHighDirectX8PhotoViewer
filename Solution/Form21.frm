VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form21 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   0  '�S���ؽu
   Caption         =   "DirectX"
   ClientHeight    =   8220
   ClientLeft      =   105
   ClientTop       =   -195
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   548
   ScaleMode       =   3  '����
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   WindowState     =   2  '�̤j��
   Begin VB.Frame Frame1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�Ϥ��ѪR�� "
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   1
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command2 
         Appearance      =   0  '����
         Caption         =   "�T�w"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         Caption         =   "��j�Ϥ����ѪR��"
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "�̭�Ϫ��ѪR��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   16
            Top             =   2400
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "2048 X 1536"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   15
            Top             =   2040
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "1600 X 1200"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "1280 X 1024"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "1024 X 768"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "800 X 600"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "640 X 480"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '�S���ؽu
      Caption         =   "Frame2"
      Height          =   2295
      Left            =   9120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   2295
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   2775
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4895
         _cy             =   4048
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.*|*.*"
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  '����
      Height          =   1290
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "��ܹϤ��ؿ�"
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton Command1 
         Appearance      =   0  '����
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
         Appearance      =   0  '����
         Caption         =   "�T�w"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  '����
         Height          =   3030
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  '����
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
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

Dim Texture() As Direct3DTexture8

Const Pic_Drive As String = "C:\Windows\Web\Wallpaper\Nature" '�w�]�Ϥ��ؿ�
Const Pic_Drive2 As String = "C:\Documents and Settings\Makubex\�ୱ\�s��Ƨ� (2)" '�w�]�Ҩ��ؿ�
Const VideoURL As String = "music\memories.mp3" '�v����m
Const Pic_Back2 As String = "SYS\index2\01.jpg"
Dim Pic_Back As String

Dim Cosine(360) As Single
Dim Sine(360) As Single
Dim Xcenter() As Single, Ycenter() As Single
Dim Xdis() As Single, Ydis() As Single
Dim ma As Integer '�Ϥ��ƶq
Dim Form_Name As String '���W��
Dim S(10) As Integer '0)g�Ϥ����ʺ�ӫ� 1)�C�ƹϼ� 2)�Ϥ��j��_�l 3)�Ϥ��j����I 4)�W�U�� 5)���Ʋ֥[ 6)�{���Y�N���� 7)���Ϥ��ؿ� 8)�H����� 9)�ƹ��S������󪺹� 10)�üƭI��
Dim Pic_W As Single '�ϼe
Dim Pic_H As Single '�ϰ�
Dim PiX As Single 'X�ѪR��
Dim PiY As Single 'Y�ѪR��

Dim old_Vertex_x As Single
Dim old_Vertex_y As Single
Dim oSx As Single '�}�l���e
Dim oSy As Single '�}�l����
Dim P_Move(5) As Single '0)����ƹ����� 1)�ƹ��I�U�h��X��m 2)�ƹ��I�U�h��Y��m 3)�ƹ��襤���Ϥ� 4)Frame1(0) ���ƹ����� 5)Frame1(1)������
Dim G_x As Single '������
Dim G_Y As Single '������
Dim AF() As Integer '�u�����Ǹ��X(�ߤ@�@��)
Dim File_ele As D3DXIMAGE_INFO '�ɮ׸�T
Dim old_X(1) As Single '�ƹ��I�U�h��X(frame1(0))
Dim old_Y(1) As Single '�ƹ��I�U�h��Y
Dim XP_Active(4) As String '���D�ʵe
Dim RandView_T As Long '�H����ܮɶ�
Dim Word As String 'Dx��r
Dim Vis_Right As Boolean '�k�����é����

Dim P_Size() As Pic_DX
Private Type Pic_DX
    A_Move As Integer '�Ϥ��w���ʦ��ơ]�k��)
    A_Turn As Integer '�Ϥ��w���স��
    Left As Single '�ӷ�
    top As Single
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
    Alpha As Byte
    Mov_CF As Byte '�O�_�Q�ƹ��I�L
    Dis_Large As Byte '�����\�Y��
    Perfect As Byte '�ӽo�{��
End Type
Private Sub Form_Load() '�����J
Call Dx_START 'DX��l
Call Three '�G���
Call Option1_Click(0) '�w�]��640 X 480
Call Load_START '�uŪ�@��
End Sub
Private Sub Form_Activate()
Do
    Call START(Dir1.Path) '��l��
    Call BackPicture(IIf(data222 = 1, Pic_Back2, Pic_Back)) '�I��
    Call Dx_Texture 'DX����
    Call Core '�֤�
    If S(6) = 1 Then S(6) = 2: Unload Me '�����{��(ESC)
    If S(7) = 1 Then Me.Caption = Form_Name: Call Ma_Clear '�M���ƭȸ��
Loop Until S(6) <> 0
Call Ma_Clear '�M���ƭȸ��
End Sub
Private Sub START(a As String) '��l�� a)�w�]�����~���е{�����|
Dim f As Byte, j As Byte, b As String

File1.Path = a '�Ϥ����|
ma = File1.ListCount - 1
If ma = -1 Then ma = 0

ReDim Texture(ma + 2) As Direct3DTexture8
ReDim Vertex(3, ma + 2) As TLVERTEX
ReDim P_Size(ma) As Pic_DX
ReDim AF(ma) As Integer
ReDim Xcenter(ma) As Single, Ycenter(ma) As Single
ReDim Xdis(3, ma) As Single, Ydis(3, ma) As Single

S(0) = 10 '(���J���賡���B�k��)'�Ϥ����ʺ�ӫ�
S(1) = 10 '�C�C�@10�i 10 X 10 = 100�i(��)
S(2) = 0 '-----------�Ϥ��j���l��
S(3) = IIf(data222 = 0, IIf(ma < S(1) ^ 2 - 1, ma, S(1) ^ 2 - 1), IIf(ma < 39, ma, 39)) '�Ϥ��j��פ��

G_x = Me.ScaleWidth / D3dParam.BackBufferWidth
G_Y = Me.ScaleHeight / D3dParam.BackBufferHeight
Pic_W = IIf(data222 = 0, D3dParam.BackBufferWidth / S(1), D3dParam.BackBufferWidth / S(1))
Pic_H = IIf(data222 = 0, D3dParam.BackBufferHeight / S(1), D3dParam.BackBufferHeight / 4)

Form_Name = Me.Caption

For f = 0 To 4
    For j = 0 To 4
        XP_Active(f) = XP_Active(f) & IIf(f = j, "��", "��")
    Next
Next

Vis_Right = False
With Frame2
    .Width = Pic_W * 2
    .Height = Pic_H * 2.7
    .Left = Me.ScaleWidth - .Width
    .top = 0
    .Visible = False
End With
With WindowsMediaPlayer1
    .Left = 0
    .top = 0
    .Width = Frame2.Width * 15
    .Height = Frame2.Height * 15
End With
End Sub
Private Sub Dx_Texture() '���J����
Dim a(1) As Single, b(1) As Single, f As Integer, t As Long, Time_S As Long, j As Integer

Call Central(a(0), a(1), b(0), b(1)) '�o�P�_�l�p��
t = GetTickCount
For f = S(2) To S(3)
    Time_S = (GetTickCount - t) \ 1000
    Call Pic_Load(f, Pic_W, Pic_H)  'Ū���Ϥ�
    Me.Caption = Form_Name & "�@" & f + 1 & "/" & S(3) + 1 & "�@" & XP_Active(f / 3 Mod 5) & "�@" & "�Ӯ� " & Time_S \ 60 & "��" & Time_S Mod 60 & "��" & "�@" & "�@" & ma + 1 & "�i"
    
    Vertex(0, f) = Ver(a(0), b(0), 0, 0) '�_�l��m
    Vertex(1, f) = Ver(a(1), b(0), 1, 0)
    Vertex(2, f) = Ver(a(0), b(1), 0, 1)
    Vertex(3, f) = Ver(a(1), b(1), 1, 1)

    j = Deal(f) '�ئa��m�p��
    With P_Size(f)
        .D_Left = Pic_W * (j Mod S(1)) '�ئa��m
        .D_Top = Pic_H * ((f \ S(1)) Mod S(1))
        .D_Width = Pic_W
        .D_Height = Pic_H
        .D_XCen = (.D_Left + .D_Width) / 2
        .D_YCen = (.D_Top + .D_Height) / 2
        .Perfect = S(0) '���ʲӽo��
        .Alpha = 15 '�z����
    End With
    
    AF(f) = f '�u������
    Call Swap(f)
    
    DoEvents
    If S(6) = 1 Or S(4) <> 0 Or S(7) = 1 Then Exit Sub '�p�G�{���N�n�����h���}�j��
    Call Render '�q��
    Call Pic_Move(f) '�u���ʥثe��������
Next

Call Rand_Seat(Pic_W * 10, Pic_H * 3, (D3dParam.BackBufferWidth), Pic_H * 6) '�H�����񪺦�m

Me.Caption = Form_Name & "�@" & f + 1 - 1 & "/" & S(3) + 1 & "�@" & "�Ӯ� " & Time_S \ 60 & "��" & Time_S Mod 60 & "��" & "�@" & "�@" & ma + 1 & "�i"

End Sub
Private Sub Core() '�֤ߵ{�� ��
Dim t As Long
Do
    DoEvents
    If S(6) = 1 Or S(7) = 1 Then Exit Sub
    If S(4) <> 0 Then Call Page '�p�G���F�W�U���h
    Call RandView '�H�����
    Call Pic_Move(S(3)) '����
    Call Render '�q��
    Sleep (20)
Loop
End Sub
Private Function Pic_coll(Ax As Single, Ay As Single, f As Integer, Bx As Single, By As Single) As Boolean
If Ax > Bx And Ax < Bx + P_Size(f).D_Width And Ay > By And Ay < By + P_Size(f).D_Height Then Pic_coll = True
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '�ƹ����U
Dim f As Integer, a As Single, b As Single, j As Integer
P_Move(1) = X / G_x
P_Move(2) = Y / G_Y
For f = S(2) To S(3)
    If Pic_coll(P_Move(1), P_Move(2), AF(f), Vertex(0, AF(f)).X, Vertex(0, AF(f)).Y) Then
        P_Move(3) = AF(f)
        old_Vertex_x = P_Move(1) - Vertex(0, AF(f)).X
        old_Vertex_y = P_Move(2) - Vertex(0, AF(f)).Y
        Call Swap(AF(f)) '�u�����ǥ洫
        Exit For
    End If
Next
If P_Size(P_Move(3)).A_Turn < 270 Then Exit Sub '�p�G�o�i���٨S�৹�h���}

If f <> S(3) + 1 Then '�O�_�S�]���j��
    S(9) = 0
    P_Size(P_Move(3)).Mov_CF = 1 '�]���w�ƹ�����
Else
    S(9) = 1 '�S�I���
End If

Select Case Button
    Case 1 '����
        P_Move(0) = 1 '������
        If f = S(3) + 1 Then
            P_Move(0) = 0 '��������
            For f = S(2) To S(3)
                If P_Size(f).A_Turn = 270 Then '�p�G�w�g�৹�h����
                    If P_Size(f).Mov_CF = 1 Then '�w�Q�ƹ����ʹL�h
                        P_Size(f).A_Move = 0
                        P_Move(3) = f: Call Form_MouseDown(2, 1, X, Y) '�p�G�O��j�����A�h�Y�p
                    End If
                End If
            Next
        Else
            Call Blend(P_Move(3)) '�z��
        End If
    Case 2 '�k��
        If f <> S(3) + 1 Or Shift = 1 Then
            With P_Size(P_Move(3))
                .Perfect = 20 '�ӽo�׬�20
                .Mov_CF = 0 '�M���w�ƹ�����
                .A_Move = 0 '�M���w���ʦ���(�k��)
                If .Large Then Call Form_MouseDown(3, 1, X, Y) '�p�G�O��j�����A�h�Y�p
                
                j = Deal(P_Move(3)) '�ئa��m�p��
                .D_Left = Pic_W * (j Mod S(1)) '�ئa��m
                .D_Top = Pic_H * ((P_Move(3) \ S(1)) Mod S(1))
                
                P_Size(P_Move(3)).Alpha = 7
                Call Blend(P_Move(3)) '�z��
            End With
        End If
    Case 3 '����
        If f <> S(3) + 1 Or Shift = 1 Then
            With P_Size(P_Move(3))
                .Perfect = 20 '�ӽo�׬�20
                .A_Move = 0 '�M���w���ʦ���(�k��)
                .Alpha = 7
                .Large = Not .Large '�Y��
                If .Large Then
                    If data222 = 0 Then '��{��
                        If PiX = -1 Then '�p�G�O��͸ѪR�׫h
                            Call Pic_Load(P_Move(3), D3DX_DEFAULT, D3DX_DEFAULT)
                            a = File_ele.Width
                            b = File_ele.Height
                        Else
                            Call Pic_Load(P_Move(3), PiX, PiY)  'D3DX_DEFAULT, D3DX_DEFAULT)
                            a = IIf(File_ele.Width < File_ele.Height, PiY, PiX) '������P�_
                            b = IIf(File_ele.Width < File_ele.Height, PiX, PiY)
                        End If
                        .D_Width = a 'File_ele.Width
                        .D_Height = b 'File_ele.Height
                    Else
                        Call Pic_Load(P_Move(3), D3DX_DEFAULT, D3DX_DEFAULT)
                        .D_Width = File_ele.Width
                        .D_Height = File_ele.Height
                    End If
                Else
                    Call Pic_Load(P_Move(3), Pic_W, Pic_H)
                    .D_Width = Pic_W
                    .D_Height = Pic_H
                End If
                .D_Left = Vertex(0, P_Move(3)).X + (Vertex(1, P_Move(3)).X - Vertex(0, P_Move(3)).X) / 2 - .D_Width / 2
                .D_Top = Vertex(0, P_Move(3)).Y + (Vertex(2, P_Move(3)).Y - Vertex(0, P_Move(3)).Y) / 2 - .D_Height / 2
                
                Call Blend(P_Move(3)) '�z��
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
Dim a As Single, b As Single
If P_Move(0) = 1 Then
    X = X / G_x
    Y = Y / G_Y
    Vertex(0, P_Move(3)).X = P_Move(1) + X - P_Move(1) - old_Vertex_x
    Vertex(0, P_Move(3)).Y = P_Move(2) + Y - P_Move(2) - old_Vertex_y
    Call Vertex_P(P_Move(3), Vertex(0, P_Move(3)).X, Vertex(0, P_Move(3)).Y, P_Size(P_Move(3)).D_Width, P_Size(P_Move(3)).D_Height)
    P_Size(P_Move(3)).Dis_Large = 1 '�]���ϥΪ̭n���ʩҥH���R�\�Y��
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If S(9) = 1 Then Exit Sub

If Button = 1 Then
    P_Move(0) = 0 '��}����
    P_Size(P_Move(3)).Alpha = 7
    If P_Size(P_Move(3)).Dis_Large = 0 Then Call Form_MouseDown(3, 1, X, Y) '�p�G�ϥΪ̥ت��O�n�Y��h�Y��
    P_Size(P_Move(3)).Dis_Large = 0 '�M������
End If
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
Dim f As Integer, a As String
Dim X(1) As Single, w As Single
Dim Y(1) As Single, h As Single

j = IIf(j > S(3), S(3), j) '����d�򬰥���
For f = S(2) To j
    If P_Size(f).A_Move < P_Size(f).Perfect Then
        If P_Size(f).A_Move = 0 Then '�p�G�������٨S�����ʹL�h
            With P_Size(f)
                .Left = Vertex(0, f).X '�ӷ�
                .top = Vertex(0, f).Y
                .Width = Vertex(1, f).X - Vertex(0, f).X
                .Height = Vertex(2, f).Y - Vertex(0, f).Y
            End With
        End If
        
        w = Vertex(1, f).X - Vertex(0, f).X
        h = Vertex(2, f).Y - Vertex(0, f).Y

        P_Size(f).A_Move = P_Size(f).A_Move + 1
        
        Vertex(0, f).X = Vertex(0, f).X + (P_Size(f).D_Left - P_Size(f).Left) / P_Size(f).Perfect
        Vertex(0, f).Y = Vertex(0, f).Y + (P_Size(f).D_Top - P_Size(f).top) / P_Size(f).Perfect
        With P_Size(f)
            w = w + (.D_Width - .Width) / P_Size(f).Perfect
            h = h + (.D_Height - .Height) / P_Size(f).Perfect
            Xcenter(f) = Xcenter(f) + (.D_XCen - .XCen) / P_Size(f).Perfect
            Ycenter(f) = Ycenter(f) + (.D_YCen - .YCen) / P_Size(f).Perfect
        End With
        
        
        Call Vertex_P(f, Vertex(0, f).X, Vertex(0, f).Y, w, h)
        DoEvents
    Else
        Call Turn(f)
        
        If P_Size(f).Alpha < 15 Then
            a = Creep(P_Size(f).Alpha, 1)
            Vertex(0, f).color = a
            Vertex(1, f).color = a
            Vertex(2, f).color = a
            Vertex(3, f).color = a
        End If
    End If
Next

End Sub
Sub Turn(L As Integer) '����֤�
Dim a(3) As Single, b(3) As Single, f As Integer, j As Integer, I As Integer

For f = S(2) To L
    If P_Size(f).A_Turn < 270 Then
        If P_Size(f).A_Turn = 0 Then '���I����I
            With P_Size(f)
                .XCen = (Vertex(0, f).X + Vertex(1, f).X) / 2 '�����I
                .YCen = (Vertex(0, f).Y + Vertex(2, f).Y) / 2
                Xcenter(f) = .XCen
                Ycenter(f) = .YCen
            End With
            
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
        I = P_Size(f).A_Turn
        
        For j = 0 To 3
            a(j) = Xcenter(f) + Xdis(j, f) / 2 + Cosine(I) - Sine(I) * Xdis(j, f) / 2
            b(j) = Ycenter(f) + Ydis(j, f) + Sine(I) + Cosine(I) * Xdis(j, f)
        Next
        Vertex(0, f) = Ver(a(0), b(0), 0, 0)
        Vertex(1, f) = Ver(a(1), b(1), 1, 0)
        Vertex(2, f) = Ver(a(2), b(2), 0, 1)
        Vertex(3, f) = Ver(a(3), b(3), 1, 1)
        
        DoEvents
    End If
Next
End Sub
Sub Render() '�q��
Dim f As Integer, color As Long
With D3ddevice
    .Clear 0, ByVal 0, D3DCLEAR_TARGET, color, 18, 0
    .BeginScene
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
            .SetTexture 0, Texture(ma + 2)
            .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, ma + 2), Len(Vertex(0, ma + 2))
             For f = S(3) To S(2) Step -1 'ø�X
                .SetTexture 0, Texture(AF(f))
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, AF(f)), Len(Vertex(0, AF(f)))
             Next
            .SetTexture 0, Texture(ma + 1)
            .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0, ma + 1), Len(Vertex(0, ma + 1))
        .SetRenderState D3DRS_ALPHABLENDENABLE, False
    .EndScene
    
    .Present ByVal 0, ByVal 0, 0, ByVal 0
End With
End Sub
Private Sub Load_START() '�uŪ�@��
Dir1.Path = IIf(data222 = 0, Pic_Drive, Pic_Drive2)
If data222 = 0 Then WindowsMediaPlayer1.URL = VideoURL

'�üƭI��
Randomize
S(10) = Int(Rnd * 3) + 1
S(10) = 63 + S(10) * 2 '=65 67 69
Pic_Back = S(10) & ".jpg" 'Pic_Back = "SYS\index2\" & S(10) & ".jpg"

Word = "�ƹ��G" & vbCrLf & _
       "�@�@����G������i�Y��Ϥ��C" & vbCrLf & _
       "�@�@�@�@�@�����i���ʹϤ��C" & vbCrLf & _
       "�@�@�k��G�I�ϫ�i���Ϥ��k��C" & vbCrLf & _
       "��L�G" & vbCrLf & _
       "�@�@Page UP�G�W�@��" & vbCrLf & _
       "�@�@Page Down�G�U�@��"
End Sub
Private Sub Page() '����
D3ddevice.SetTexture 0, Nothing
Erase Texture(), P_Size(), AF()
ReDim Texture(ma + 2), P_Size(ma), AF(ma)
Vertex(0, 0).X = -Pic_W
Call Vertex_P(0, -Pic_W, -Pic_H, Pic_W, Pic_H)
Vertex(0, ma + 1).X = -Pic_W
Call Vertex_P(ma + 1, -Pic_W, -Pic_H, Pic_W, Pic_H)

S(5) = S(5) + S(4) '�֥[����
S(2) = S(1) ^ 2 * S(5) '�j��_�l
S(3) = (S(1) ^ 2) * (S(5) + 1) - 1 '�j��פ�
S(3) = IIf(S(3) > ma, ma, S(3)) '�w���W�X�̤j�Ϥ��i��
S(4) = 0 '�M��
Call BackPicture(Pic_Back)
Call Dx_Texture
End Sub
Private Sub BackPicture(FileName As String) '�I��
Set Texture(ma + 2) = D3DX.CreateTextureFromFileEx(D3ddevice, FileName, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, D3DX_DEFAULT, File_ele, ByVal 0)
Vertex(0, ma + 2) = Ver(0, 0, 0, 0)
Vertex(1, ma + 2) = Ver((D3dParam.BackBufferWidth), 0, 1, 0)
Vertex(2, ma + 2) = Ver(0, (D3dParam.BackBufferHeight), 0, 1)
Vertex(3, ma + 2) = Ver((D3dParam.BackBufferWidth), (D3dParam.BackBufferHeight), 1, 1)
End Sub
Private Sub Rand_Seat(X As Single, Y As Single, w As Single, h As Single) '�H�����񪺦�m
'�H�����񪺦�m
Vertex(0, ma + 1) = Ver(X, Y, 0, 0)
Vertex(1, ma + 1) = Ver(w, Y, 1, 0)
Vertex(2, ma + 1) = Ver(X, h, 0, 1)
Vertex(3, ma + 1) = Ver(w, h, 1, 1)
End Sub
Private Sub Pic_Load(f, w As Single, h As Single) 'Ū��
On Error GoTo Err:
Dim FileName As String

FileName = File1.Path & "\" & File1.List(f)
Set Texture(f) = D3DX.CreateTextureFromFileEx(D3ddevice, FileName, w, h, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, File_ele, ByVal 0)
Exit Sub

Err:
    MsgBox "���~���Ϥ�!", 64, "�T��"
End Sub
Private Sub RandView() '�H�����
Dim a As Integer, c As String
Randomize
If S(8) = 0 Then
    RandView_T = GetTickCount
    a = Int(Rnd * (S(3) - S(2) + 1)) + S(2)
    Vertex(0, ma + 1).color = &HFF000000
    Vertex(1, ma + 1).color = &HFF000000
    Vertex(2, ma + 1).color = &HFF000000
    Vertex(3, ma + 1).color = &HFF000000
    Set Texture(ma + 1) = Texture(a)
End If
c = Creep(S(8), 1)
Vertex(0, ma + 1).color = c
Vertex(1, ma + 1).color = c
Vertex(2, ma + 1).color = c
Vertex(3, ma + 1).color = c
If GetTickCount >= RandView_T + 2000 Then S(8) = 0
End Sub
Private Function Creep(a, b) As String '����
Dim c As String, d As Byte
d = IIf(b > 0, 15, 0) '���ƭ˼�

a = IIf(a = d, d, a + 1 * b)
If a > 9 Then c = UCase(Chr(65 + a Mod 10)) Else c = a
Creep = "&H" & c & c & "FFFFFF"
End Function
Private Sub Blend(f) '�z����
Vertex(0, f).color = &H77FFFFFF
Vertex(1, f).color = &H77FFFFFF
Vertex(2, f).color = &H77FFFFFF
Vertex(3, f).color = &H77FFFFFF
End Sub
Sub Central(a0 As Single, a1 As Single, b0 As Single, b1 As Single) '�o�P�_�l�p��
With D3dParam
    a0 = .BackBufferWidth / S(1) * 9 - Pic_W * 1.2 '�p���m
    b0 = .BackBufferHeight / S(1) * 8 - Pic_H
    a1 = .BackBufferWidth / S(1) * 9 + Pic_W * 1.2
    b1 = .BackBufferHeight - Pic_H / 4
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
If data222 = 1 And KeyCode <> 27 Then MsgBox "�Шϥάݥ����Ϥ�������", 64, "�T��": Exit Sub
Select Case KeyCode
    Case 33 'Page Down
        If S(5) > 0 Then S(4) = -1
    Case 34 '40 'Page UP
        If S(3) >= ma Then S(5) = -1
        S(4) = 1
    Case 27 'ESC
        Unload Me
End Select
End Sub
Private Sub Frame1_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    P_Move(4 + index) = 1
    old_X(index) = X \ 15
    old_Y(index) = Y \ 15
End If
End Sub
Private Sub Frame1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If P_Move(4 + index) = 1 Then
    X = X \ 15
    Y = Y \ 15
    With Frame1(index)
        .Left = .Left + X - old_X(index)
        .top = .top + Y - old_Y(index)
    End With
End If
End Sub
Private Sub Frame1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then P_Move(4 + index) = 0
End Sub
Private Sub Pic_Drives_Click(index As Integer) '�\���C
If data222 = 1 Then MsgBox "�Шϥάݥ����Ϥ�������", 64, "�T��": Exit Sub
Select Case index
    Case 0 '�Ϥ��ؿ�
        Frame1(0).Visible = True
    Case 1 '�ﶵ
        Frame1(1).Visible = True
    Case 2 '�W�@��
        Call Form_KeyDown(33, 0)
    Case 3 '�U�@��
        Call Form_KeyDown(34, 0)
End Select
End Sub
Private Sub Music_Click(index As Integer)
Dim f As Integer
Dim a As String
Select Case index
    Case 0 '�v�����
        If data222 = 1 Then MsgBox "�Шϥάݥ����Ϥ�������", 64, "�T��": Exit Sub
        CommonDialog1.ShowOpen
        a = CommonDialog1.FileName
        If a = "" Then Exit Sub
        WindowsMediaPlayer1.URL = a
    Case 1 '�k�����é����
        If P_Size(S(3)).A_Turn < 270 Then MsgBox "�е��ݩҦ����Ϥ����৹��!", 64, "�T��": Exit Sub
        Vis_Right = Not Vis_Right
        
        Frame2.Visible = Vis_Right
        Pic_W = IIf(Vis_Right, (D3dParam.BackBufferWidth / S(1) * 8) / S(1), D3dParam.BackBufferWidth / S(1))
        For f = S(2) To S(3)
            P_Size(f).D_Width = Pic_W
            P_Size(f).Mov_CF = 1
        Next
        Call Form_MouseDown(1, 0, -1, -1)
        If Vis_Right Then
            Call Rand_Seat(Pic_W * 10, Pic_H * 3, (D3dParam.BackBufferWidth), Pic_H * 6) '�H�����񪺦�m
        Else
            Call Rand_Seat(Pic_W * 10 + Pic_W, Pic_H * 3, (D3dParam.BackBufferWidth + Pic_W), Pic_H * 6) '�H�����񪺦�m
        End If
End Select
End Sub
Private Sub Change_Ver_Click() '��������
Unload Me
'If data222 = 0 Then form1.Show
End Sub
Private Sub Command1_Click(index As Integer)
If index = 0 Then S(7) = 1
Frame1(0).Visible = False
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
Private Sub Option1_Click(index As Integer)
Dim a() As String
If index = 6 Then PiX = -1: Exit Sub
a = Split(Option1(index).Caption, " X ")
PiX = a(0) 'X�ѪR��
PiY = a(1) 'Y�ѪR��
End Sub
Private Sub Command2_Click()
Frame1(1).Visible = False
End Sub
Private Sub Ma_Clear() '�M���ƭȸ��
data222 = 0
Erase Texture(), Vertex(), P_Size(), AF(), P_Move(), Xcenter(), Ycenter(), Xdis(), Ydis(), S(), XP_Active()  '�M�����(���Ϥ��ؿ�)
End Sub
Private Sub Form_Unload(Cancel As Integer) '��沾��
If S(6) < 2 Then '�Y�N����
    S(6) = 1
    Cancel = 1
Else '����
    Call UnDX
    'Form2.Show
End If
End Sub
