Public Class Form1
    ''' <summary>
    ''' 建立HOOK钩子，钩取窗体消息
    ''' 但是只允许HOOK同进程窗体，如果要钩取其他进程，需要做成DLL，并注入
    ''' </summary>
    ''' <returns></returns>
    Private Declare Function GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As IntPtr
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewinteger As Integer) As Integer
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Integer, ByVal hwnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Delegate Function SubWndProc(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
    Private Const GWL_WNDPROC = (-4)
    Private OldWinProc As Integer
    Private Const WM_PAINT = &HF

    Dim DesktopHandle As IntPtr = Me.Handle

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim leftWndProc As SubWndProc = AddressOf MyWndproc
        OldWinProc = GetWindowLong(DesktopHandle, GWL_WNDPROC) '记录原始的委托地址，方便退出程序时恢复
        SetWindowLong(DesktopHandle, GWL_WNDPROC, Runtime.InteropServices.Marshal.GetFunctionPointerForDelegate(leftWndProc)) '开始截取消息

        MsgBox(OldWinProc,, "初始委托地址")
        If OldWinProc = 0 Then End
        Me.Text = "Hook 成功！"
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        SetWindowLong(DesktopHandle, GWL_WNDPROC, OldWinProc) '取消消息截取
    End Sub

    Private Function MyWndproc(ByVal hwnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        If Msg = WM_PAINT Then
            Debug.Print(Msg & " - " & wParam & " - " & lParam)
        End If

        Return CallWindowProc(OldWinProc, hwnd, Msg, wParam, lParam)
    End Function

End Class
