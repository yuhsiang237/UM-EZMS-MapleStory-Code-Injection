Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Module Module2

 
    Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As UInteger, ByVal wMapType As UInteger) As UInteger

    Public Const WM_KEYDOWN = &H100

    Public Const WM_KEYUP = &H101

    Public Const WM_CHAR = &H102

    Public Const VK_0 = &H30
    Public Const VK_1 = &H31
    Public Const VK_2 = &H32
    Public Const VK_3 = &H33
    Public Const VK_4 = &H34
    Public Const VK_5 = &H35
    Public Const VK_6 = &H36
    Public Const VK_7 = &H37
    Public Const VK_8 = &H38
    Public Const VK_9 = &H39
    Public Const VK_A = &H41
    Public Const VK_B = &H42
    Public Const VK_C = &H43
    Public Const VK_D = &H44
    Public Const VK_E = &H45
    Public Const VK_F = &H46
    Public Const VK_G = &H47
    Public Const VK_H = &H48
    Public Const VK_I = &H49
    Public Const VK_J = &H4A
    Public Const VK_K = &H4B
    Public Const VK_L = &H4C
    Public Const VK_M = &H4D
    Public Const VK_N = &H4E
    Public Const VK_O = &H4F
    Public Const VK_P = &H50
    Public Const VK_Q = &H51
    Public Const VK_R = &H52
    Public Const VK_S = &H53
    Public Const VK_T = &H54
    Public Const VK_U = &H55
    Public Const VK_V = &H56
    Public Const VK_W = &H57
    Public Const VK_X = &H58
    Public Const VK_Y = &H59
    Public Const VK_Z = &H5A
    Function MakeKeyLparam(ByVal VirtualKey As Long, ByVal flag As Long) As Long

        Dim s As String

        Dim Firstbyte As String 'lparam參數的24-31位

        If flag = WM_KEYDOWN Then '如果是按下鍵

            Firstbyte = "00"

        Else

            Firstbyte = "C0" '如果是釋放鍵"

        End If

        Dim Scancode As Long

        '獲得鍵的掃描碼

        Scancode = MapVirtualKey(VirtualKey, 0)

        Dim Secondbyte As String 'lparam參數的16-23位元，即虛擬鍵掃描碼

        Secondbyte = Right("00" & Hex(Scancode), 2)

        s = Firstbyte & Secondbyte & "0001" '0001為lparam參數的0-15位，即發送次數和其他擴展資訊"

        MakeKeyLparam = Val("&H" & s)

        Return MakeKeyLparam

    End Function
End Module
