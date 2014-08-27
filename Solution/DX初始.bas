Attribute VB_Name = "Module21"
Option Explicit

Public Dx As DirectX8
Public D3d As Direct3D8
Public D3ddevice As Direct3DDevice8
Public D3dPlayMode As D3DDISPLAYMODE
Public D3dParam As D3DPRESENT_PARAMETERS

Sub Dx_START(Bwindow As Boolean, Dxwindow As Form)

Dim BehaviorFlag As Long
Dim Caps As D3DCAPS8

Set Dx = New DirectX8
Set D3d = Dx.Direct3DCreate()
Set Dxwindow = Form21

If Bwindow Then
    D3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, D3dPlayMode
    D3dParam.BackBufferFormat = D3dPlayMode.Format
    D3dParam.Windowed = True
Else
    D3dParam.BackBufferFormat = D3DFMT_R5G6B5
    D3dParam.Windowed = False
End If
    
With D3dParam
    .SwapEffect = D3DSWAPEFFECT_FLIP
    .BackBufferCount = 1
    .BackBufferWidth = Screen.Width / 15
    .BackBufferHeight = Screen.Height / 15
    .hDeviceWindow = Dxwindow.hWnd
End With
D3d.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
BehaviorFlag = IIf(Caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DCREATE_SOFTWARE_VERTEXPROCESSING)

Set D3ddevice = D3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Dxwindow.hWnd, BehaviorFlag, D3dParam)

With D3ddevice
    .SetVertexShader TLFVF
    .SetRenderState D3DRS_LIGHTING, False
    .SetRenderState D3DRS_ZWRITEENABLE, False
    
    .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    .SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
End With
Set D3DX = New D3DX8
End Sub
Sub UnDX()
Set D3ddevice = Nothing
Set D3d = Nothing
Set Dx = Nothing
End Sub

'參考資料 KYO VBDX http://kyovbdx.myweb.hinet.net/
