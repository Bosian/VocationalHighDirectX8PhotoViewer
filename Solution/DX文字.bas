Attribute VB_Name = "Module23"
Public MainFont As D3DXFont
Public Fnt As New StdFont
Public MainFontDesc As IFont

Public TextRect As RECT '�@�Ӯئ��r��X�d��RECT
Sub DXFont()

Fnt.Name = "�з���" '�r��
Fnt.Size = 20      '�j�p

Set MainFontDesc = Fnt
Set MainFont = D3DX.CreateFont(D3ddevice, MainFontDesc.hFont) '�إ�MainFont(D3DXFont���O)����

'�]�w��X�d��
TextRect.Right = 130 'Width
TextRect.bottom = 100 'Height
TextRect.Left = D3dParam.BackBufferWidth / 2 - TextRect.Right
TextRect.Top = D3dParam.BackBufferHeight / 2 - TextRect.bottom
End Sub

'�ѦҸ�� KYO VBDX http://kyovbdx.myweb.hinet.net/
