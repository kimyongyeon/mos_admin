Imports Microsoft.win32 'Registry Functions
Imports System.Runtime.InteropServices 'API functions

Public Class frmTaskBarSettings

    Private Declare Function FindWindow Lib "user32.dll" Alias _
    "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Int32 'Find External Window

    Private Declare Function FindWindowEx Lib "user32.dll" Alias _
    "FindWindowExA" (ByVal hWnd1 As Int32, ByVal hWnd2 As Int32, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Int32 'Find Child Window Of External Window

    Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Int32, _
    ByVal nCmdShow As Int32) As Int32 'Show A Window

    Private Declare Function PostMessage Lib "user32.dll" Alias _
    "PostMessageA" (ByVal hwnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, _
    ByVal lParam As Int32) As Int32 'Post Message To Window

    Private Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Int32, _
    ByVal fEnable As Int32) As Int32 'Enable A Window

    Private Declare Function SendMessageSTRING Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Int32, _
    ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As String) As Int32 'SendMessage lParam = String

    Declare Auto Function SendMessageTimeout Lib "User32" ( _
    ByVal hWnd As Integer, _
    ByVal Msg As UInt32, _
    ByVal wParam As Integer, _
    ByVal lParam As Integer, _
    ByVal fuFlags As UInt32, _
    ByVal uTimeout As UInt32, _
    ByRef lpdwResult As IntPtr _
    ) As Long 'Send Message & Wait

    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As _
    Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32 'Normal SendMessage

    Private Declare Function GetDesktopWindow Lib "user32" () As IntPtr 'Get Handle To Desktop

    Private Const WM_WININICHANGE = &H1A 'INI File Update
    Private Const HWND_BROADCAST = &HFFFF& 'Send To All
    Private Const WM_SETTINGCHANGE = &H1A 'Setting Change
    Private Const SMTO_ABORTIFHUNG = &H2 'Stop If Hang
    Private Const WM_COMMAND As Int32 = &H111 'Send Command
    Private Const WM_USER As Int32 = &H400 'User
    Private Const WM_SETTEXT = &HC 'Change Text
    Private Const WM_GETTEXT = &HD 'Get Text

    Private AB As New TBAppBar 'AppBar Object

    Public Sub EnvRefresh() ' Refresh Explorer
        Dim EnvResult As IntPtr 'Result
        SendMessageTimeout(HWND_BROADCAST, _
        Convert.ToUInt32(WM_SETTINGCHANGE), _
        0, 0, _
        Convert.ToUInt32(SMTO_ABORTIFHUNG), _
        Convert.ToUInt32(2000), _
        EnvResult) 'Broadcast A Setting Change To All
    End Sub

    Private Sub btnTaskProp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTaskProp.Click
        Process.Start("rundll32.exe", "shell32.dll,Options_RunDLL 1") 'Taskbar & Start Menu Properties
    End Sub

    Private Sub btnClock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClock.Click
        Select Case btnClock.Text
            Case "Show Clock" 'If Hidden
                Dim TaskBarWin As Long, TrayWin As Long, ClockWin As Long

                TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString) 'Find TaskBar 
                TrayWin = FindWindowEx(TaskBarWin, 0, "TrayNotifyWnd", vbNullString) 'Find Tray Window
                ClockWin = FindWindowEx(TrayWin, 0, "TrayClockWClass", vbNullString) 'Find Clock Window
                ShowWindow(ClockWin, 1) 'Show Clock

                btnClock.Text = "Hide Clock"
            Case "Hide Clock" 'If Shown
                Dim TaskBarWin As Long, TrayWin As Long, ClockWin As Long

                TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString)
                TrayWin = FindWindowEx(TaskBarWin, 0, "TrayNotifyWnd", vbNullString)
                ClockWin = FindWindowEx(TrayWin, 0, "TrayClockWClass", vbNullString)
                ShowWindow(ClockWin, 0) 'Hide Clock

                btnClock.Text = "Show Clock"
        End Select
    End Sub

    Private Sub btnLock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLock.Click
        Dim TaskBarWin As Long

        TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString) 'Find TaskBar
        PostMessage(TaskBarWin, WM_COMMAND, 424, vbNullString) 'Lock TaskBar
    End Sub

    Private Sub btnAutoHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAutoHide.Click

        Select Case btnAutoHide.Text
            Case "Auto Hide TaskBar"
                btnAutoHide.Text = "Auto Hide TaskBar Off"
                AB.AutoHide = True 'Set AutoHide On
            Case "Auto Hide TaskBar Off"
                btnAutoHide.Text = "Auto Hide TaskBar"
                AB.AutoHide = False 'Set AutoHide Off
        End Select
    End Sub

    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Select Case btnTop.Text
            Case "Keep TaskBar on Top"
                AB.AlwaysOnTop = True 'Set Always On Top On
                btnTop.Text = "Don't Keep TaskBar on Top"
            Case "Don't Keep TaskBar on Top"
                AB.AlwaysOnTop = False 'Set Always On Top Off
                btnTop.Text = "Keep TaskBar on Top"
        End Select
    End Sub

    Private Sub btnGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroup.Click
        Select Case btnGroup.Text
            Case "Group Similar TaskBar Buttons"
                Dim GroupRet As Long 'Used With SendMessage
                Dim TaskBarWin As Long

                TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString) 'Find taskbar

                Dim GroupKey As RegistryKey = _
                Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _
                True) 'Open Registry Key

                GroupKey.SetValue("TaskbarGlomming", 1, RegistryValueKind.DWord) 'Set Grouping On
                GroupRet = SendMessage(TaskBarWin, WM_WININICHANGE, 0&, 0&) 'Store New Setting
                btnGroup.Text = "Don't Group Similar TaskBar Buttons"
            Case "Don't Group Similar TaskBar Buttons"
                Dim GroupRet As Long
                Dim TaskBarWin As Long

                TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString)

                Dim GroupKey As RegistryKey = _
                Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _
                True)

                GroupKey.SetValue("TaskbarGlomming", 0, RegistryValueKind.DWord) 'Set Grouping Off
                GroupRet = SendMessage(TaskBarWin, WM_WININICHANGE, 0&, 0&)
                btnGroup.Text = "Group Similar TaskBar Buttons"
        End Select
    End Sub

    Private Sub btnContents_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContents.Click
        Select Case btnContents.Text
            Case "Show Taskbar Contents"
                Dim TaskBarWin As Long, TaskButtonWin As Long

                TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString) 'Find Taskbar
                TaskButtonWin = FindWindowEx(TaskBarWin, 0, "ReBarWindow32", vbNullString) 'Find TaskBar Button Area
                ShowWindow(TaskButtonWin, 1) 'Show Active Buttons
                btnContents.Text = "Hide Taskbar Contents"

            Case "Hide Taskbar Contents"
                Dim TaskBarWin As Long, TaskButtonWin As Long

                TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString)
                TaskButtonWin = FindWindowEx(TaskBarWin, 0, "ReBarWindow32", vbNullString)
                ShowWindow(TaskButtonWin, 0) 'Hide Active Buttons
                btnContents.Text = "Show Taskbar Contents"

        End Select
    End Sub

    Private Sub btnShowDesktop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowDesktop.Click
        'Create Instance Of Shell Class
        'Referenced COM Library "Microsoft Shell Controls And Automation" (shell32.dll)

        '2009.12.16 : SCS
        'Dim objShell As New Shell32.ShellClass()
        'DirectCast(objShell, Shell32.IShellDispatch4).ToggleDesktop() 'Show Desktop


    End Sub

    Private Sub btnFavourites_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFavourites.Click
        Select Case btnFavourites.Text
            Case "Show Favourites In Menu"

                Dim FavKey As RegistryKey = _
                Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _
                True)
                FavKey.SetValue("StartMenuFavorites", 1, RegistryValueKind.DWord) 'Show Favourites Menu

                EnvRefresh() 'Refresh Explorer.exe
                btnFavourites.Text = "Hide Favourites In Menu"
            Case "Hide Favourites In Menu"

                Dim FavKey As RegistryKey = _
                Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _
                True)
                FavKey.SetValue("StartMenuFavorites", 0, RegistryValueKind.DWord) 'Hide Favourites

                EnvRefresh()
                btnFavourites.Text = "Show Favourites In Menu"
        End Select
    End Sub

    Private Sub btnStartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartButton.Click
        Select Case btnStartButton.Text
            Case "Disable Start Button"
                Dim TaskBarWin As Long, StartButtonWin As Long

                TaskBarWin = FindWindowEx(0, 0, "Shell_TrayWnd", Nothing) 'Find TaskBar
                StartButtonWin = FindWindowEx(TaskBarWin, 0, "Button", Nothing) 'Find Start Button
                EnableWindow(StartButtonWin, False) 'Disable Start Button
                btnStartButton.Text = "Enable Start Button"
            Case "Enable Start Button"
                Dim TaskBarWin As Long, StartButtonWin As Long

                TaskBarWin = FindWindowEx(0, 0, "Shell_TrayWnd", Nothing)
                StartButtonWin = FindWindowEx(TaskBarWin, 0, "Button", Nothing)
                EnableWindow(StartButtonWin, True) 'Enable Start Button
                btnStartButton.Text = "Disable Start Button"
        End Select

    End Sub

    Private Sub btnInactiveTrayIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInactiveTrayIcons.Click
        Select Case btnInactiveTrayIcons.Text
            Case "AutoHide Inactive Tray Icons Off"
                Dim TrayWinRet As Long
                Dim TrayWindow As Long

                TrayWindow = FindWindow("Shell_TrayWnd", vbNullString) 'Find TaskBar

                Dim InactiveKey As RegistryKey = _
                Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _
                True)

                InactiveKey.SetValue("EnableAutoTray", 0, RegistryValueKind.DWord) 'Set Enable AutoTray Off
                TrayWinRet = SendMessage(TrayWindow, WM_WININICHANGE, 0&, 0&) 'Store
                btnInactiveTrayIcons.Text = "AutoHide Inactive Tray Icons On"

                EnvRefresh() 'Refresh Explorer.exe
            Case "AutoHide Inactive Tray Icons On"
                Dim TrayWinRet As Long
                Dim TrayWindow As Long

                TrayWindow = FindWindow("Shell_TrayWnd", vbNullString)

                Dim InactiveKey As RegistryKey = _
                Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _
                True)

                InactiveKey.SetValue("EnableAutoTray", 1, RegistryValueKind.DWord) 'Set Enable AutoTray On
                TrayWinRet = SendMessage(TrayWindow, WM_WININICHANGE, 0&, 0&)
                btnInactiveTrayIcons.Text = "AutoHide Inactive Tray Icons Off"

                EnvRefresh()
        End Select
    End Sub

    Private Sub SetStartCaption(ByVal NewStr As String)
        Dim TaskBarWin As Long
        Dim StartWin As Long
        Dim StartText As String

        TaskBarWin = FindWindow("Shell_TrayWnd", vbNullString) 'Find TaskBar
        StartWin = FindWindowEx(TaskBarWin, 0&, "button", vbNullString) 'Find Start Button
        StartText = Microsoft.VisualBasic.Left(NewStr, 5) 'Set StartButton Text
        SendMessageSTRING(StartWin, WM_SETTEXT, 256, StartText) 'Send The Message

        Exit Sub 'Don't Do Anything Else
    End Sub

    Private Sub btnText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnText.Click
        Select Case btnText.Text
            Case "Change Start Button Text"
                Dim NewText As String
                NewText = InputBox("Enter New Text") 'Get New Start Button Text
                SetStartCaption(NewText)
                btnText.Text = "Change Back Start Button Text"
            Case "Change Back Start Button Text"
                SetStartCaption("Start") 'Revert Back To "Start"
                btnText.Text = "Change Start Button Text"
        End Select

    End Sub

    Public Class TBAppBar

        Private Declare Function GetAppBarMessage Lib "shell32" Alias "SHAppBarMessage" _
           (ByVal dwMessage As Integer, ByRef pData As APPBARDATA) As Integer 'Get Message Sent By App Bar

        Private Declare Function SetAppBarMessage Lib "shell32" Alias "SHAppBarMessage" _
           (ByVal dwMessage As Integer, ByRef pData As APPBARDATA) As Integer 'Send Message To App BAr

        Private Structure APPBARDATA 'AppBar Structure
            Dim cbSize As Integer
            Dim hwnd As Integer
            Dim uCallbackMessage As Integer
            Dim uEdge As Integer
            Dim rc As System.Drawing.Rectangle
            Dim lParam As Integer
        End Structure

        Private Const ABM_GETSTATE As Int32 = &H4 'Get Current State
        Private Const ABM_GETTASKBARPOS As Int32 = &H5 'Get TaskBar Position
        Private Const ABM_SETSTATE As Int32 = &HA 'Apply Setting(s)
        Private Const ABS_AUTOHIDE As Int32 = &H1 'Autohide
        Private Const ABS_ALWAYSONTOP As Int32 = &H2 'Always on Top

        Private TBAppBAutoHide As Boolean
        Private TBAppBarAlwaysOnTop As Boolean

        Public Sub New()
            Me.GetState() 'Get Current State
        End Sub

        Private Sub GetState()

            Dim AppBarSetting As New APPBARDATA 'What Setting?

            AppBarSetting.cbSize = Marshal.SizeOf(AppBarSetting) 'Initialise

            Dim AppBarState As Int32 = GetAppBarMessage(ABM_GETSTATE, AppBarSetting) 'Get Current State

            Select Case AppBarState

                Case 0 'Nothing Set
                    TBAppBAutoHide = False
                    TBAppBarAlwaysOnTop = False

                Case ABS_ALWAYSONTOP 'Always On Top
                    TBAppBAutoHide = False
                    TBAppBarAlwaysOnTop = True

                Case Else
                    TBAppBAutoHide = True 'AutoHide

            End Select

        End Sub

        Private Sub SetState() 'Apply Settings

            Dim AppBarSetting As New APPBARDATA 'Setting We Want To Apply
            AppBarSetting.cbSize = Marshal.SizeOf(AppBarSetting) 'Initialise

            If Me.AutoHide Then
                AppBarSetting.lParam = ABS_AUTOHIDE 'AutoHide
            End If

            If Me.AlwaysOnTop Then
                AppBarSetting.lParam = AppBarSetting.lParam Or ABS_ALWAYSONTOP 'Always On Top
            End If

            SetAppBarMessage(ABM_SETSTATE, AppBarSetting)

        End Sub

        Public Property AutoHide() As Boolean 'Autohide
            Get
                Return TBAppBAutoHide
            End Get
            Set(ByVal Value As Boolean)
                TBAppBAutoHide = Value
                Me.SetState()
            End Set
        End Property

        Public Property AlwaysOnTop() As Boolean 'Always On Top
            Get
                Return TBAppBarAlwaysOnTop
            End Get
            Set(ByVal Value As Boolean)
                TBAppBarAlwaysOnTop = Value
                Me.SetState()
            End Set
        End Property

    End Class

    
End Class
