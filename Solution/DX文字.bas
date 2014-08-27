Attribute VB_Name = "Module23"
Public MainFont As D3DXFont
Public Fnt As New StdFont
Public MainFontDesc As IFont

Public TextRect As RECT '一個框住文字輸出範圍的RECT
Sub DXFont()

Fnt.Name = "標楷體" '字型
Fnt.Size = 20      '大小

Set MainFontDesc = Fnt
Set MainFont = D3DX.CreateFont(D3ddevice, MainFontDesc.hFont) '建立MainFont(D3DXFont類別)物件

'設定輸出範圍
TextRect.Right = 130 'Width
TextRect.bottom = 100 'Height
TextRect.Left = D3dParam.BackBufferWidth / 2 - TextRect.Right
TextRect.Top = D3dParam.BackBufferHeight / 2 - TextRect.bottom
End Sub

'參考資料 KYO VBDX http://kyovbdx.myweb.hinet.net/
