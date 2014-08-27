Attribute VB_Name = "Module21"
Option Explicit

Public Dx As DirectX8
Public D3d As Direct3D8
Public D3dDevice As Direct3DDevice8
Public D3dPlayMode As D3DDISPLAYMODE
Public D3dParam As D3DPRESENT_PARAMETERS

Sub Dx_START()
Dim Bwindow As Boolean, Dxwindow As Form

Set Dx = New DirectX8
Set D3d = Dx.Direct3DCreate()
Set Dxwindow = Form21
Bwindow = True
If Bwindow Then
    D3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, D3dPlayMode
    D3dParam.BackBufferFormat = D3dPlayMode.Format
    D3dParam.Windowed = True
Else
    D3dParam.BackBufferFormat = D3DFMT_R5G6B5
End If
D3dParam.SwapEffect = D3DSWAPEFFECT_FLIP
D3dParam.BackBufferCount = 1
D3dParam.BackBufferWidth = Screen.Width \ 15
D3dParam.BackBufferHeight = Screen.Height \ 15
D3dParam.hDeviceWindow = Dxwindow.hWnd

Dim BehaviorFlag As Long
Dim Caps As D3DCAPS8

D3d.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
BehaviorFlag = IIf(Caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DCREATE_SOFTWARE_VERTEXPROCESSING)

Set D3dDevice = D3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Dxwindow.hWnd, BehaviorFlag, D3dParam)

D3dDevice.SetVertexShader TLFVF
D3dDevice.SetRenderState D3DRS_LIGHTING, False
D3dDevice.SetRenderState D3DRS_ZWRITEENABLE, False
End Sub

Sub UnDX()
Set D3dDevice = Nothing
Set D3d = Nothing
Set Dx = Nothing
End Sub
