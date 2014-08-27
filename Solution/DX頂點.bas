Attribute VB_Name = "Module22"
Public Vertex() As TLVERTEX '頂點
Public D3DX As D3DX8

Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    uv As D3DVECTOR2
End Type

Type D3DXIMAGE_INFO
    Width As Long
    Height As Long
    MipLevels As Long
    Depth As Long
    Format As Long
End Type

Public Const TLFVF = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1

Function Ver(X As Single, Y As Single, tu As Single, tv As Single) As TLVERTEX
    Ver.X = X
    Ver.Y = Y
    Ver.z = 0
    Ver.rhw = 1#
    Ver.color = &HFFFFFFFF
    Ver.uv.X = tu
    Ver.uv.Y = tv
End Function

'參考資料 KYO VBDX http://kyovbdx.myweb.hinet.net/


