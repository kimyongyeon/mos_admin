Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Drawing
Imports System.Threading

Module Keyboard
    Public Declare Function UnhookWindowsHookEx Lib "user32" _
      (ByVal hHook As Integer) As Integer

    Public Declare Function SetWindowsHookEx Lib "user32" _
      Alias "SetWindowsHookExA" (ByVal idHook As Integer, _
      ByVal lpfn As KeyboardHookDelegate, ByVal hmod As Integer, _
      ByVal dwThreadId As Integer) As Integer

    Private Declare Function GetAsyncKeyState Lib "user32" _
      (ByVal vKey As Integer) As Integer

    Private Declare Function CallNextHookEx Lib "user32" _
      (ByVal hHook As Integer, _
      ByVal nCode As Integer, _
      ByVal wParam As Integer, _
      ByVal lParam As KBDLLHOOKSTRUCT) As Integer

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)> _
    Private Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function

    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    ' Low-Level Keyboard Constants
    Private Const HC_ACTION As Integer = 0
    Private Const LLKHF_EXTENDED As Integer = &H1
    Private Const LLKHF_INJECTED As Integer = &H10
    Private Const LLKHF_ALTDOWN As Integer = &H20
    Private Const LLKHF_UP As Integer = &H80

    ' Virtual Keys
    Public Const VK_TAB = &H9
    Public Const VK_CONTROL = &H11
    Public Const VK_ESCAPE = &H1B
    Public Const VK_DELETE = &H2E
    Public Const VK_LWIN = &H5B

    Private Const WH_KEYBOARD_LL As Integer = 13&
    Public KeyboardHandle As Integer


    ' Implement this function to block as many
    ' key combinations as you'd like
    Public Function IsHooked( _
      ByRef Hookstruct As KBDLLHOOKSTRUCT) As Boolean

        Debug.WriteLine("Hookstruct.vkCode: " & Hookstruct.vkCode)
        Debug.WriteLine(Hookstruct.vkCode = VK_ESCAPE)
        Debug.WriteLine(Hookstruct.vkCode = VK_TAB)

        If (Hookstruct.vkCode = VK_ESCAPE) And _
          CBool(GetAsyncKeyState(VK_CONTROL) _
          And &H8000) Then

            Call HookedState("Ctrl + Esc blocked")
            Return True
        End If

        If (Hookstruct.vkCode = VK_TAB) And _
          CBool(Hookstruct.flags And _
          LLKHF_ALTDOWN) Then

            Call HookedState("Alt + Tab blockd")
            Return True
        End If

        If (Hookstruct.vkCode = VK_ESCAPE) And _
          CBool(Hookstruct.flags And _
            LLKHF_ALTDOWN) Then

            Call HookedState("Alt + Escape blocked")
            Return True
        End If

        'If (Hookstruct.vkCode = VK_LWIN) Then
        '    Call HookedState("Left Windows Logo Key blocked")
        '    Return True
        'End If

        If (Hookstruct.vkCode = Keys.LWin) Then
            Call HookedState("Left Windows Logo Key blocked")
            Return True
        End If

        Return False
    End Function

    Private Sub HookedState(ByVal Text As String)
        Debug.WriteLine(Text)
    End Sub

    Public Function KeyboardCallback(ByVal Code As Integer, _
      ByVal wParam As Integer, _
      ByRef lParam As KBDLLHOOKSTRUCT) As Integer

        If (Code = HC_ACTION) Then
            Debug.WriteLine("Calling IsHooked")

            If (IsHooked(lParam)) Then
                Return 1
            End If

        End If

        Return CallNextHookEx(KeyboardHandle, _
          Code, wParam, lParam)

    End Function


    Public Delegate Function KeyboardHookDelegate( _
      ByVal Code As Integer, _
      ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) _
                   As Integer

    <MarshalAs(UnmanagedType.FunctionPtr)> _
    Private callback As KeyboardHookDelegate

    Public Sub HookKeyboard()
        callback = New KeyboardHookDelegate(AddressOf KeyboardCallback)

        'KeyboardHandle = SetWindowsHookEx( _
        '  WH_KEYBOARD_LL, callback, _
        '  Marshal.GetHINSTANCE( _
        '  [Assembly].GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                KeyboardHandle = SetWindowsHookEx(WH_KEYBOARD_LL, callback, GetModuleHandle(curModule.ModuleName), 0)
            End Using
        End Using

        Call CheckHooked()
    End Sub

    Public Sub CheckHooked()
        If (Hooked()) Then
            Debug.WriteLine("Keyboard hooked")
        Else
            Debug.WriteLine("Keyboard hook failed: " & Err.LastDllError)
        End If
    End Sub

    Private Function Hooked()
        Hooked = KeyboardHandle <> 0
    End Function

    Public Sub UnhookKeyboard()
        If (Hooked()) Then
            Call UnhookWindowsHookEx(KeyboardHandle)
        End If
    End Sub

End Module
