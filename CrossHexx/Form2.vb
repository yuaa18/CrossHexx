
Imports System.Runtime.InteropServices

Public Class Form2
    '----------オーバレイ宣言----------------
    Public Const HWND_TOPMOST = (-1)
    Public Const SWP_NOSIZE = &H1&
    Public Const SWP_NOMOVE = &H2&

    Public Const GWL_EXSTYLE As Long = (-20)

    Public Const WS_EX_LAYERED = &H80000
    Public Const WS_EX_TRANSPARENT = &H20
    Public Const LWA_ALPHA = &H2

    Const TOPMOST_FLAGS As UInteger = (SWP_NOSIZE Or SWP_NOMOVE)

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function SetWindowPos(ByVal hWnd As IntPtr, _
        ByVal hWndInsertAfter As IntPtr, _
        ByVal X As Integer, _
        ByVal Y As Integer, _
        ByVal cx As Integer, _
        ByVal cy As Integer, _
        ByVal uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function SetWindowLong(hWnd As IntPtr, _
        <MarshalAs(UnmanagedType.I4)> nIndex As Integer, _
        dwNewLong As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowLong(hWnd As IntPtr, _
        <MarshalAs(UnmanagedType.I4)> nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function SetLayeredWindowAttributes(hwnd As IntPtr, crKey As UInteger, bAlpha As Byte, dwFlags As UInteger) As Boolean
    End Function
    '-------------ここまで-------------------
  

    Private Sub Form2_MouseHover(sender As Object, e As EventArgs) Handles MyBase.MouseHover
        SetWindowLong(Me.Handle, GWL_EXSTYLE, GetWindowLong(Me.Handle, GWL_EXSTYLE) Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Me.Handle, 0, 255, LWA_ALPHA)
        SetWindowPos(Me.Handle, HWND_TOPMOST, 0, 0, 0, 0, _
            TOPMOST_FLAGS)
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs)
        SetWindowLong(Me.Handle, GWL_EXSTYLE, GetWindowLong(Me.Handle, GWL_EXSTYLE) Xor WS_EX_LAYERED Xor WS_EX_TRANSPARENT)

        SetLayeredWindowAttributes(Me.Handle, 0, 255, LWA_ALPHA)
        SetWindowPos(Me.Handle, HWND_TOPMOST, 0, 0, 0, 0, _
            TOPMOST_FLAGS)
    End Sub
End Class