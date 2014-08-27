Attribute VB_Name = "Module21"
Option Explicit

Public Dx As DirectX8
Public D3d As Direct3D8
Public D3ddevice As Direct3DDevice8
Public D3dPlayMode As D3DDISPLAYMODE
Public D3dParam As D3DPRESENT_PARAMETERS

Sub Dx_START()
Dim Bwindow As Boolean, Dxwindow As Form

Set Dx = New DirectX8
Set D3d = Dx.Direct3DCreate()
Set Dxwindow = Form21

If Bwindow Then
    D3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, D3dPlayMode
    D3dParam.BackBufferFormat = D3dPlayMode.Format
    D3dParam.Windowed = True
Else
    D3dParam.BackBufferFormat = D3DFMT_R5G6B5
End If
D3dParam.SwapEffect = D3DSWAPEFFECT_FLIP
D3dParam.BackBufferCount = 1
D3dParam.BackBufferWidth = 1280
D3dParam.BackBufferHeight = 1024
D3dParam.hDeviceWindow = Dxwindow.hWnd

Dim BehaviorFlag As Long
Dim Caps As D3DCAPS8

D3d.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
BehaviorFlag = IIf(Caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DCREATE_SOFTWARE_VERTEXPROCESSING)

Set D3ddevice = D3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Dxwindow.hWnd, BehaviorFlag, D3dParam)

D3ddevice.SetVertexShader TLFVF
D3ddevice.SetRenderState D3DRS_LIGHTING, False
D3ddevice.SetRenderState D3DRS_ZWRITEENABLE, False

D3ddevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
D3ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3ddevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD

Set D3DX = New D3DX8
End Sub

Sub UnDX()
Set D3ddevice = Nothing
Set D3d = Nothing
Set Dx = Nothing
End Sub


