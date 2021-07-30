Imports System.Runtime.InteropServices

Public NotInheritable Class WinVer
    Const WUNKNOWNSTR = ("unknown Windows version")

    Const W95STR = ("Windows 95")
    Const W95SP1STR = ("Windows 95 SP1")
    Const W95OSR2STR = ("Windows 95 OSR2")
    Const W98STR = ("Windows 98")
    Const W98SP1STR = ("Windows 98 SP1")
    Const W98SESTR = ("Windows 98 SE")
    Const WMESTR = ("Windows ME")

    Const WNT351STR = ("Windows NT 3.51")
    Const WNT4STR = ("Windows NT 4")
    Const W2KSTR = ("Windows 2000")
    Const WXPSTR = ("Windows XP")
    Const WVISTASTR = ("Windows Vista")
    Const W7STR = "Windows 7"
    Const W2003STR = ("Windows Server 2003")

    Const WCESTR = ("Windows CE")

    Const WUNKNOWN = 0

    Const W9XFIRST = 1
    Const W95 = 1
    Const W95SP1 = 2
    Const W95OSR2 = 3
    Const W98 = 4
    Const W98SP1 = 5
    Const W98SE = 6
    Const WME = 7

    Const W9XLAST = 99

    Const WNTFIRST = 101
    Const WNT351 = 101
    Const WNT4 = 102
    Const W2K = 103
    Const WXP = 104
    Const W2003 = 105
    Const WVISTA = 106
    Const W7 = 107
    Const WNTLAST = 199

    Const WCEFIRST = 201
    Const WCE = 201
    Const WCELAST = 299


    Const PRODUCT_BUSINESS = &H6    ' Business Edition
    Const PRODUCT_BUSINESS_N = &H10    ' Business Edition
    Const PRODUCT_CLUSTER_SERVER = &H12    ' Cluster Server Edition
    Const PRODUCT_DATACENTER_SERVER = &H8    ' Server Datacenter Edition (full installation)
    Const PRODUCT_DATACENTER_SERVER_CORE = &HC    ' Server Datacenter Edition (core installation)
    Const PRODUCT_ENTERPRISE = &H4    ' Enterprise Edition
    Const PRODUCT_ENTERPRISE_N = &H1B    ' Enterprise Edition
    Const PRODUCT_ENTERPRISE_SERVER = &HA    ' Server Enterprise Edition (full installation)
    Const PRODUCT_ENTERPRISE_SERVER_CORE = &HE    ' Server Enterprise Edition (core installation)
    Const PRODUCT_ENTERPRISE_SERVER_IA64 = &HF    ' Server Enterprise Edition for Itanium-based Systems
    Const PRODUCT_HOME_BASIC = &H2    ' Home Basic Edition
    Const PRODUCT_HOME_BASIC_N = &H5    ' Home Basic Edition
    Const PRODUCT_HOME_PREMIUM = &H3    ' Home Premium Edition
    Const PRODUCT_HOME_PREMIUM_N = &H1A    ' Home Premium Edition
    Const PRODUCT_HOME_SERVER = &H13    ' Home Server Edition
    Const PRODUCT_SERVER_FOR_SMALLBUSINESS = &H18    ' Server for Small Business Edition
    Const PRODUCT_SMALLBUSINESS_SERVER = &H9    ' Small Business Server
    Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM = &H19    ' Small Business Server Premium Edition
    Const PRODUCT_STANDARD_SERVER = &H7    ' Server Standard Edition (full installation)
    Const PRODUCT_STANDARD_SERVER_CORE = &HD    ' Server Standard Edition (core installation)
    Const PRODUCT_STARTER = &HB    ' Starter Edition
    Const PRODUCT_STORAGE_ENTERPRISE_SERVER = &H17    ' Storage Server Enterprise Edition
    Const PRODUCT_STORAGE_EXPRESS_SERVER = &H14    ' Storage Server Express Edition
    Const PRODUCT_STORAGE_STANDARD_SERVER = &H15    ' Storage Server Standard Edition
    Const PRODUCT_STORAGE_WORKGROUP_SERVER = &H16    ' Storage Server Workgroup Edition
    Const PRODUCT_UNDEFINED = &H0    ' An unknown product
    Const PRODUCT_ULTIMATE = &H1    ' Ultimate Edition
    Const PRODUCT_ULTIMATE_N = &H1C    ' Ultimate Edition
    Const PRODUCT_WEB_SERVER = &H11

    Public Shared Function GetMajorVersion()
        Return m_osinfo.dwMajorVersion
    End Function
    Public Shared Function GetMinorVersion() As Integer
        Return m_osinfo.dwMinorVersion
    End Function
    Public Shared Function GetBuildNumber() As Integer
        Return (m_osinfo.dwBuildNumber And &HFFFF) '; }	// needed for 9x
    End Function
    Public Function GetPlatformId() As Integer
        Return (m_osinfo.dwPlatformId)
    End Function
    Public Function GetServicePackString() As String
        Return m_osinfo.szCSDVersion
    End Function
    Shared m_bInitialized As Boolean
    Shared m_osinfo As Win32.Kernel32.OSVERSIONINFOEX
    Shared m_dwVistaProductType As Int32




    '#If VER_PLATFORM_WIN32_WINDOWS Then
    Const VER_PLATFORM_WIN32_WINDOWS = 1
    '#End If
    '#If VER_PLATFORM_WIN32_NT Then
    Const VER_PLATFORM_WIN32_NT = 2
    '#End If
    '#If VER_PLATFORM_WIN32_CE Then
    Const VER_PLATFORM_WIN32_CE = 3
    '#End If
    '// from winnt.h
    '#If VER_NT_WORKSTATION Then
    Const VER_NT_WORKSTATION = &H1
    '#End If
    '#If VER_SUITE_PERSONAL Then
    Const VER_SUITE_PERSONAL = &H200
    '#End If


    Shared Sub New()

        m_dwVistaProductType = 0

        '	ZeroMemory(&m_osinfo, sizeof(m_osinfo));

        m_osinfo.dwOSVersionInfoSize = Marshal.SizeOf(GetType(Win32.Kernel32.OSVERSIONINFO))
        If (Win32.Kernel32.GetVersionEx(m_osinfo)) Then

            m_bInitialized = True

            If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT) AndAlso
   (m_osinfo.dwMajorVersion >= 5)) Then

                '// get extended version info for 2000 and later

                'ZeroMemory(&m_osinfo, sizeof(m_osinfo))

                m_osinfo.dwOSVersionInfoSize = Marshal.SizeOf(GetType(Win32.Kernel32.OSVERSIONINFOEX))

                Win32.Kernel32.GetVersionEx(m_osinfo)

                'TRACE(_T("dwMajorVersion =%u\n"), m_osinfo.dwMajorVersion );
                'TRACE(_T("dwMinorVersion =%u\n"), m_osinfo.dwMinorVersion );
                'TRACE(_T("dwBuildNumber=%u\n"), m_osinfo.dwBuildNumber);
                'TRACE(_T("suite mask=%u\n"), m_osinfo.wSuiteMask);
                'TRACE(_T("product type=%u\n"), m_osinfo.wProductType);
                'TRACE(_T("sp major=%u\n"), m_osinfo.wServicePackMajor);
                'TRACE(_T("sp minor=%u\n"), m_osinfo.wServicePackMinor);
            End If
        End If

    End Sub
    Public Shared Function GetWinVersionString() As String

        Dim strVersion As String = WUNKNOWNSTR

        Dim nVersion As Integer = GetWinVersion()

        Select Case nVersion

            Case W95 : strVersion = W95STR
            Case W95SP1 : strVersion = W95SP1STR
            Case W95OSR2 : strVersion = W95OSR2STR
            Case W98 : strVersion = W98STR
            Case W98SP1 : strVersion = W98SP1STR
            Case W98SE : strVersion = W98SESTR
            Case WME : strVersion = WMESTR
            Case WNT351 : strVersion = WNT351STR
            Case WNT4 : strVersion = WNT4STR
            Case W2K : strVersion = W2KSTR
            Case WXP : strVersion = WXPSTR
            Case W2003 : strVersion = W2003STR
            Case WVISTA : strVersion = WVISTASTR
            Case WCE : strVersion = WCESTR
            Case W7 : strVersion = W7STR
        End Select

        Return strVersion
    End Function
    Public Shared Function GetWinVersion() As Integer

        Dim nVersion As Integer = WUNKNOWN

        Dim dwPlatformId As Integer = m_osinfo.dwPlatformId
        Dim dwMinorVersion As Integer = m_osinfo.dwMinorVersion
        Dim dwMajorVersion As Integer = m_osinfo.dwMajorVersion
        Dim dwBuildNumber As Integer = m_osinfo.dwBuildNumber And &HFFFF ';	// Win 9x needs this

        If ((dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) AndAlso (dwMajorVersion = 4)) Then

            If ((dwMinorVersion < 10) AndAlso (dwBuildNumber = 950)) Then

                nVersion = W95

            ElseIf ((dwMinorVersion < 10) AndAlso
              ((dwBuildNumber > 950) AndAlso (dwBuildNumber <= 1080))) Then

                nVersion = W95SP1

            ElseIf ((dwMinorVersion < 10) AndAlso (dwBuildNumber > 1080)) Then

                nVersion = W95OSR2

            ElseIf ((dwMinorVersion = 10) AndAlso (dwBuildNumber = 1998)) Then

                nVersion = W98

            ElseIf ((dwMinorVersion = 10) AndAlso
              ((dwBuildNumber > 1998) AndAlso (dwBuildNumber < 2183))) Then

                nVersion = W98SP1

            ElseIf ((dwMinorVersion = 10) AndAlso (dwBuildNumber >= 2183)) Then

                nVersion = W98SE

            ElseIf (dwMinorVersion = 90) Then

                nVersion = WME
            End If

        ElseIf (dwPlatformId = VER_PLATFORM_WIN32_NT) Then

            If ((dwMajorVersion = 3) AndAlso (dwMinorVersion = 51)) Then

                nVersion = WNT351

            ElseIf ((dwMajorVersion = 4) AndAlso (dwMinorVersion = 0)) Then

                nVersion = WNT4

            ElseIf ((dwMajorVersion = 5) AndAlso (dwMinorVersion = 0)) Then

                nVersion = W2K

            ElseIf ((dwMajorVersion = 5) AndAlso (dwMinorVersion = 1)) Then

                nVersion = WXP

            ElseIf ((dwMajorVersion = 5) AndAlso (dwMinorVersion = 2)) Then

                nVersion = W2003

            ElseIf ((dwMajorVersion = 6) AndAlso (dwMinorVersion = 0)) Then

                nVersion = WVISTA
                GetVistaProductType()
            ElseIf ((dwMajorVersion = 6) AndAlso (dwMinorVersion = 1)) Then
                nVersion = W7

            End If

        ElseIf (dwPlatformId = VER_PLATFORM_WIN32_CE) Then

            nVersion = WCE
        End If

        Return nVersion
    End Function
    Public Shared Function GetServicePackNT() As Integer

        Dim nServicePack As Integer = 0

        'for i as Integer = 0 to 		 (m_osinfo.szCSDVersion(i) <> "\0") &&
        '		 (i < (sizeof(m_osinfo.szCSDVersion)/sizeof(TCHAR)));
        '	 i++)
        'Dim i As Integer
        'Do
        '    If m_osinfo.szCSDVersion(i) <> "0" AndAlso (i < Marshal.SizeOf(GetType(m_osinfo.szCSDVerison)) / Marshal.SizeOf(GetType(String))) Then
        '        Exit Do
        '    End If
        '    If (_istdigit(m_osinfo.szCSDVersion(i))) Then

        '        nServicePack = _ttoi(m_osinfo.szCSDVersion(i))
        '        Exit Do
        '    End If
        '    i += 1
        'Loop
        For Each c In m_osinfo.szCSDVersion
            If Char.IsDigit(c) Then
                nServicePack = AscW(c)
                Exit For
            End If
        Next
        Return nServicePack
    End Function
    Public Shared Function IsXP() As Boolean

        If (GetWinVersion() = WXP) Then

            Return True
        End If
        Return False
    End Function
    Public Shared Function IsXPHome() As Boolean

        If (GetWinVersion() = WXP) Then

            If (m_osinfo.wSuiteMask And VER_SUITE_PERSONAL) Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Shared Function IsXPPro() As Boolean

        If (GetWinVersion() = WXP) Then

            If ((m_osinfo.wProductType = VER_NT_WORKSTATION) AndAlso Not IsXPHome()) Then
                Return True
            End If
        End If

        Return False
    End Function
    Public Shared Function IsXPSP2() As Boolean

        If (GetWinVersion() = WXP) Then

            If (GetServicePackNT() = 2) Then
                Return True
            End If

        End If

        Return False
    End Function
    '#If SM_MEDIACENTER Then
    Const SM_MEDIACENTER = 87
    '#End If
    Public Shared Function IsMediaCenter() As Boolean

        If (Win32.User32.GetSystemMetrics(SM_MEDIACENTER)) Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function IsWin2003() As Boolean

        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT) AndAlso
  (m_osinfo.dwMajorVersion = 5) AndAlso
  (m_osinfo.dwMinorVersion = 2)) Then
            Return True
        End If

        Return False
    End Function
    'Private Delegate Function lpfnGetProductInfo(ByVal p1 As Int32, ByVal p2 As Int32, ByVal p3 As Int32, ByVal p4 As Int32, ByVal p5 As Int32) As Boolean
    Private Declare Function GetProductInfo Lib "kernel32" _
  (ByVal dwOSMajorVersion As Integer,
   ByVal dwOSMinorVersion As Integer,
   ByVal dwSpMajorVersion As Integer,
   ByVal dwSpMinorVersion As Integer,
  <Out()> ByRef pdwReturnedProductType As Integer) As Integer

    '<DllImport("kernel32.dll")> Private Shared Function GetProcAddress(ByVal ModuleHandle As IntPtr, ByVal ProcName As String) As lpfnGetProductInfo
    'End Function
    Public Shared Function GetVistaProductType() As Integer

        If (m_dwVistaProductType = 0) Then

            'typedef BOOL (FAR PASCAL * lpfnGetProductInfo) (DWORD, DWORD, DWORD, DWORD, PDWORD);

            'Dim hKernel32 As IntPtr = Win32.Kernel32.GetModuleHandle("KERNEL32.DLL")
            'If (hKernel32) Then

            '    Dim pGetProductInfo As lpfnGetProductInfo = GetProcAddress(hKernel32, "GetProductInfo")

            '    If (pGetProductInfo.) Then
            '        pGetProductInfo(6, 0, 0, 0, m_dwVistaProductType)
            '    End If
            'End If
            GetProductInfo(6, 0, 0, 0, m_dwVistaProductType)
        End If

        Return m_dwVistaProductType
    End Function
    Public Shared Function GetVistaProductString() As String

        Dim strProductType As String = ("")

        Select Case (m_dwVistaProductType)

            Case PRODUCT_BUSINESS : strProductType = ("Business Edition")
            Case PRODUCT_BUSINESS_N : strProductType = ("Business Edition")
            Case PRODUCT_CLUSTER_SERVER : strProductType = ("Cluster Server Edition")
            Case PRODUCT_DATACENTER_SERVER : strProductType = ("Server Datacenter Edition (full installation)")
            Case PRODUCT_DATACENTER_SERVER_CORE : strProductType = ("Server Datacenter Edition (core installation)")
            Case PRODUCT_ENTERPRISE : strProductType = ("Enterprise Edition")
            Case PRODUCT_ENTERPRISE_N : strProductType = ("Enterprise Edition")
            Case PRODUCT_ENTERPRISE_SERVER : strProductType = ("Server Enterprise Edition (full installation)")
            Case PRODUCT_ENTERPRISE_SERVER_CORE : strProductType = ("Server Enterprise Edition (core installation)")
            Case PRODUCT_ENTERPRISE_SERVER_IA64 : strProductType = ("Server Enterprise Edition for Itanium-based Systems")
            Case PRODUCT_HOME_BASIC : strProductType = ("Home Basic Edition")
            Case PRODUCT_HOME_BASIC_N : strProductType = ("Home Basic Edition")
            Case PRODUCT_HOME_PREMIUM : strProductType = ("Home Premium Edition")
            Case PRODUCT_HOME_PREMIUM_N : strProductType = ("Home Premium Edition")
            Case PRODUCT_HOME_SERVER : strProductType = ("Home Server Edition")
            Case PRODUCT_SERVER_FOR_SMALLBUSINESS : strProductType = ("Server for Small Business Edition")
            Case PRODUCT_SMALLBUSINESS_SERVER : strProductType = ("Small Business Server")
            Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM : strProductType = ("Small Business Server Premium Edition")
            Case PRODUCT_STANDARD_SERVER : strProductType = ("Server Standard Edition (full installation)")
            Case PRODUCT_STANDARD_SERVER_CORE : strProductType = ("Server Standard Edition (core installation)")
            Case PRODUCT_STARTER : strProductType = ("Starter Edition")
            Case PRODUCT_STORAGE_ENTERPRISE_SERVER : strProductType = ("Storage Server Enterprise Edition")
            Case PRODUCT_STORAGE_EXPRESS_SERVER : strProductType = ("Storage Server Express Edition")
            Case PRODUCT_STORAGE_STANDARD_SERVER : strProductType = ("Storage Server Standard Edition")
            Case PRODUCT_STORAGE_WORKGROUP_SERVER : strProductType = ("Storage Server Workgroup Edition")
            Case PRODUCT_UNDEFINED : strProductType = ("An unknown product")
            Case PRODUCT_ULTIMATE : strProductType = ("Ultimate Edition")
            Case PRODUCT_ULTIMATE_N : strProductType = ("Ultimate Edition")
            Case PRODUCT_WEB_SERVER : strProductType = ("Web Server Edition")

                'default: break;
        End Select

        Return strProductType
    End Function
    Public Shared Function Is7() As Boolean
        Return GetWinVersion() = W7
    End Function
    Public Shared Function IsVista() As Boolean

        If (GetWinVersion() = WVISTA) Then

            Return True
        End If

        Return False
    End Function
    Public Shared Function IsVistaHome() As Boolean

        If (GetWinVersion() = WVISTA) Then

            Select Case (m_dwVistaProductType)

                Case PRODUCT_HOME_BASIC, PRODUCT_HOME_BASIC_N, PRODUCT_HOME_PREMIUM, PRODUCT_HOME_PREMIUM_N, PRODUCT_HOME_SERVER
                    Return True
            End Select
        End If

        Return False
    End Function
    Public Shared Function IsVistaBusiness() As Boolean

        If (GetWinVersion() = WVISTA) Then

            Select Case (m_dwVistaProductType)

                Case PRODUCT_BUSINESS, PRODUCT_BUSINESS_N
                    Return True
            End Select
        End If

        Return False
    End Function
    Public Shared Function IsVistaEnterprise() As Boolean

        If (GetWinVersion() = WVISTA) Then

            Select Case (m_dwVistaProductType)

                Case PRODUCT_ENTERPRISE, PRODUCT_ENTERPRISE_N, PRODUCT_ENTERPRISE_SERVER, PRODUCT_ENTERPRISE_SERVER_CORE, PRODUCT_ENTERPRISE_SERVER_IA64
                    Return True
            End Select
        End If


        Return False
    End Function
    Public Shared Function IsVistaUltimate() As Boolean

        If (GetWinVersion() = WVISTA) Then

            Select Case (m_dwVistaProductType)

                Case PRODUCT_ULTIMATE, PRODUCT_ULTIMATE_N
                    Return True
            End Select
        End If

        Return False
    End Function
    Public Shared Function IsWin2KorLater() As Boolean

        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT) AndAlso
  (m_osinfo.dwMajorVersion >= 5)) Then
            Return True
        End If

        Return False
    End Function
    Public Shared Function IsXPorLater() As Boolean
        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT) AndAlso
  (((m_osinfo.dwMajorVersion = 5) AndAlso (m_osinfo.dwMinorVersion > 0)) OrElse
  (m_osinfo.dwMajorVersion > 5))) Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function IsVistaOrGreater() As Boolean
        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT) AndAlso
(((m_osinfo.dwMajorVersion = 6) AndAlso (m_osinfo.dwMinorVersion >= 0)) OrElse
(m_osinfo.dwMajorVersion > 6))) Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function Is7OrGreater() As Boolean
        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT) AndAlso
(((m_osinfo.dwMajorVersion = 6) AndAlso (m_osinfo.dwMinorVersion >= 1)) OrElse
(m_osinfo.dwMajorVersion > 6))) Then
            Return True
        End If
        Return False
    End Function
    Public Shared Function IsWin95() As Boolean

        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) AndAlso
         (m_osinfo.dwMajorVersion = 4) AndAlso
         (m_osinfo.dwMinorVersion < 10)) Then
            Return True
        End If

        Return False
    End Function
    Public Shared Function IsWin98() As Boolean

        If ((m_osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) AndAlso
         (m_osinfo.dwMajorVersion = 4) AndAlso
         (m_osinfo.dwMinorVersion >= 10)) Then
            Return True
        End If

        Return False
    End Function
    Public Shared Function IsWinCE() As Boolean

        Return (GetWinVersion() = WCE)
    End Function
End Class
