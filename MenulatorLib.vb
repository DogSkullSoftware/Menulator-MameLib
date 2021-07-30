Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Threading
'Imports Microsoft.DirectX
'Imports Microsoft.DirectX.DirectSound
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System

<HideModuleName()>
Public Module Extensions
    <Extension()> Public Function PadRect(ByVal p As System.Windows.Forms.Padding, ByVal r As Rectangle) As Rectangle
        Return New Rectangle(r.X + p.Left, r.Y + p.Top, r.Width - p.Horizontal, r.Height - p.Vertical)
    End Function
    <Extension()> Public Function PadRect(ByVal p As System.Windows.Forms.Padding, ByVal r As RectangleF) As RectangleF
        Return New RectangleF(r.X + p.Left, r.Y + p.Top, r.Width - p.Horizontal, r.Height - p.Vertical)
    End Function
    <Extension()> Public Function OffsetRect(ByVal r As Rectangle, ByVal p As Point) As Rectangle
        Return New Rectangle(r.X + p.X, r.Y + p.Y, r.Width, r.Height)
    End Function
    <Extension()> Public Function OffsetRect(ByVal r As RectangleF, ByVal p As PointF) As RectangleF
        Return New RectangleF(r.X + p.X, r.Y + p.Y, r.Width, r.Height)
    End Function
    <Extension()> Public Function ResizeRect(ByVal r As RectangleF, ByVal x As Single, ByVal y As Single) As RectangleF
        Return New RectangleF(r.X, r.Y, r.Width + x, r.Height + y)
    End Function
    <Extension()> Public Function ResizeRect(ByVal r As Rectangle, ByVal x As Integer, ByVal y As Integer) As Rectangle
        Return New Rectangle(r.X, r.Y, r.Width + x, r.Height + y)
    End Function
    <Extension()> Public Sub DrawImageUnscaledAndClipped(ByVal g As Graphics, ByVal image As Image, ByVal rect As RectangleF)
        If (image Is Nothing) Then
            Throw New ArgumentNullException("image")
        End If
        Dim srcWidth As Single = Math.Min(rect.Width, image.Width)
        Dim srcHeight As Single = Math.Min(rect.Height, image.Height)
        'g.DrawImage(image, rect, 0, 0, srcWidth, srcHeight, GraphicsUnit.Pixel)
        g.DrawImage(image, rect.X, rect.Y, New RectangleF(0, 0, srcWidth, srcHeight), GraphicsUnit.Pixel)

    End Sub
    <Extension()> Public Sub DrawImage(ByVal g As Graphics, ByVal image As Image, ByVal p As PointF, ByVal i As Imaging.ImageAttributes)
        g.DrawImage(image, New PointF() {p, New PointF(p.X + image.Width, p.Y), New PointF(p.X, p.Y + image.Height)}, New RectangleF(0, 0, image.Width, image.Height), GraphicsUnit.Pixel, i)
    End Sub

    Public Function GetEncryptPassword(ByVal value As String) As String
        Dim algorithm As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
        Dim data As Byte() = algorithm.ComputeHash(System.Text.Encoding.UTF8.GetBytes(value))
        Dim md5 As String = ""
        For i As Integer = 0 To data.Length - 1
            md5 &= data(i).ToString("x2").ToLowerInvariant()
        Next
        Return md5
    End Function
    ''' <summary>
    ''' Converts absolute path data to relative path data.
    ''' "c:\test", "c:\test\doc.txt" = "doc.txt"
    ''' "c:\test", "c:\doc.txt" = ".\doc.txt"
    ''' </summary>
    ''' <param name="strRelativePath"></param>
    ''' <param name="strPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AbsolutePathToRelative(ByVal strRelativePath As String, ByVal strPath As String) As String
        Dim s() As String = Split(strPath, ";")
        Dim u As New Uri(strRelativePath)
        For t As Integer = 0 To UBound(s)

            s(t) = Uri.UnescapeDataString(u.MakeRelativeUri(New Uri(u, s(t))).ToString)
            If s(t) = "" Then s(t) = "."
        Next
        Return Join(s, ";")
    End Function
    Public Function AbsolutePathToRelative(ByVal strRelativePath As String, ByVal strPaths() As String) As String()
        Dim u As New Uri(strRelativePath)
        For t As Integer = 0 To UBound(strPaths)
            strPaths(t) = Uri.UnescapeDataString(u.MakeRelativeUri(New Uri(strPaths(t))).ToString)
        Next
        Return strPaths

    End Function
    ''' <summary>
    ''' Converts relative path data to absoulte path data.
    ''' "c:\test", "doc.txt" = "c:\test\doc.txt"
    ''' "c:\test", ".\doc.txt" = "c:\doc.txt"
    ''' </summary>
    ''' <param name="strRelativePath"></param>
    ''' <param name="strPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RelativePathToAbsolute(ByVal strRelativePath As String, ByVal strPath As String) As String
        Dim s() As String = Split(strPath, ";")
        For t As Integer = 0 To UBound(s)
            s(t) = IO.Path.GetFullPath(IO.Path.Combine(strRelativePath, s(t)))
        Next
        Return Join(s, ";")
    End Function
    Public Function RelativePathToAbsolute(ByVal strRelativePath As String, ByVal strPaths() As String) As String()
        For t As Integer = 0 To UBound(strPaths)
            strPaths(t) = IO.Path.GetFullPath(IO.Path.Combine(strRelativePath, strPaths(t)))
        Next
        Return strPaths
    End Function
    Public Function StockIcon(ByVal stockID As Win32.Shell32.SHSTOCKICONID) As Icon
        Dim ps As New Win32.Shell32.SHSTOCKICONINFO
        Dim flags As Win32.Shell32.StockIconInfoFlags = Win32.Shell32.StockIconInfoFlags.SHGSI_ICONLOCATION
        ps.cbSize = Marshal.SizeOf(GetType(Win32.Shell32.SHSTOCKICONINFO))
        'ps.hIcon = IntPtr.Zero
        Win32.Shell32.SHGetStockIconInfo(stockID, flags, ps)
        'StockIcon = Icon.FromHandle(ps.hIcon).ToBitmap
        'DestroyIcon(ps.hIcon)
        Dim z As IntPtr, zz As IntPtr
        Win32.Shell32.SHExtractIconsW(ps.szPath, ps.iIcon, 128, 128, z, zz, 1, 0)
        StockIcon = Icon.FromHandle(z).Clone
        Win32.User32.DestroyIcon(z)
    End Function


    Public Function ShellDisplayText(ByVal filename As String) As String
        Dim ppsi As Win32.Shell32.IShellItem = Nothing
        Try
            If filename <> "" Then filename = IO.Path.GetFullPath(filename)
            Win32.Shell32.SHCreateItemFromParsingName(filename, IntPtr.Zero, Win32.Shell32.IID_IShellItem, ppsi)
        Catch
            Return IO.Path.GetFileName(filename)
        End Try
        Dim iz As IntPtr
        ppsi.GetDisplayName(Win32.Shell32.SIGDN.NORMALDISPLAY, iz)
        ShellDisplayText = Marshal.PtrToStringUni(iz)
        Marshal.ReleaseComObject(ppsi)
    End Function



    Public Function ShellThumbnail(ByVal filename As String, ByVal size As Size) As Bitmap
        Dim ppsi As Win32.Shell32.IShellItem = Nothing
        Dim hbitmap As IntPtr = IntPtr.Zero
        'Dim uuid As New Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")
        If filename <> "" Then filename = IO.Path.GetFullPath(filename)
        Win32.Shell32.SHCreateItemFromParsingName(filename, IntPtr.Zero, Win32.Shell32.IID_IShellItem, ppsi)
        'Dim iz As IntPtr
        'ppsi.GetDisplayName(SIGDN.NORMALDISPLAY, iz)
        'DisplayName = Marshal.PtrToStringUni(iz)

        Dim ret = DirectCast(ppsi, Win32.Shell32.IShellItemImageFactory).GetImage(size,
 Win32.Shell32.SIIGBF.SIIGBF_BIGGERSIZEOK, hbitmap)
        'Dim source As Bitmap = Bitmap.FromHbitmap(hbitmap)
        'source.MakeTransparent()
        'Dim source As Bitmap = Icon.FromHandle(hbitmap).ToBitmap
        'Dim hBitmap As IntPtr = ConvertPixelByPixel(hbitmap)

        Dim source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions)
        Dim stride As Integer = source.PixelWidth * source.Format.BitsPerPixel / 8
        Dim bits(stride * source.PixelHeight * source.Format.BitsPerPixel / 8) As Byte
        source.CopyPixels(bits, stride, 0)

        Dim g As GCHandle = GCHandle.Alloc(bits, GCHandleType.Pinned)
        Dim bitmap As New System.Drawing.Bitmap(
      source.Width, source.Height, stride, Imaging.PixelFormat.Format32bppPArgb, g.AddrOfPinnedObject)
        g.Free()
        'If bitmap.Width > size.Width Or bitmap.Height > size.Height Then
        Dim b2 As New Drawing.Bitmap(size.Width, size.Height, Imaging.PixelFormat.Format32bppArgb)

        Dim r = ScaleRect(bitmap.Size, New RectangleF(0, 0, size.Width, size.Height), False)
        Dim gr As Graphics = Graphics.FromImage(b2)
        'gr.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        'gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        'gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        SetUpGraphics(gr)
        gr.DrawImage(bitmap, r, New Rectangle(Point.Empty, bitmap.Size), GraphicsUnit.Pixel)
        gr.Dispose()
        ShellThumbnail = b2.Clone(New Rectangle(Point.Empty, b2.Size), Imaging.PixelFormat.Format32bppArgb)
        gr.Dispose()
        b2.Dispose()
        'ElseIf bitmap.Width > size.Width Then

        'ElseIf bitmap.Height > size.Height Then
        'Else
        '    ShellThumbnail = bitmap.Clone(New Rectangle(0, 0, bitmap.Width, bitmap.Height), PixelFormat.Format32bppArgb)
        'End If
        bitmap.Dispose()

        Marshal.Release(hbitmap)
        Marshal.ReleaseComObject(ppsi)
        source = Nothing

    End Function

    Public CompositionQuality As System.Drawing.Drawing2D.CompositingQuality = Drawing2D.CompositingQuality.Invalid
    Public InterpolationMode As System.Drawing.Drawing2D.InterpolationMode = Drawing2D.InterpolationMode.Invalid
    Public SmoothingMode As System.Drawing.Drawing2D.SmoothingMode = Drawing2D.SmoothingMode.Invalid
    Public PixelOffsetMode As System.Drawing.Drawing2D.PixelOffsetMode = Drawing2D.PixelOffsetMode.Invalid
    Public TextRenderingHint As System.Drawing.Text.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
    Public Sub SetUpGraphics(ByRef e As Graphics)
        If CompositionQuality <> Drawing2D.CompositingQuality.Invalid Then e.CompositingQuality = CompositionQuality
        If InterpolationMode <> Drawing2D.InterpolationMode.Invalid Then e.InterpolationMode = InterpolationMode
        If SmoothingMode <> Drawing2D.SmoothingMode.Invalid Then e.SmoothingMode = SmoothingMode
        If PixelOffsetMode <> Drawing2D.PixelOffsetMode.Invalid Then e.PixelOffsetMode = PixelOffsetMode
        e.TextRenderingHint = TextRenderingHint

        'e.CompositingMode = Drawing2D.CompositingMode.SourceOver
    End Sub

    Public Function ScaleSize(ByVal source As SizeF, ByVal dest As SizeF) As SizeF
        Dim AspectRate As Single = source.Width / source.Height
        Dim ret As New SizeF
        If AspectRate > dest.Width / dest.Height Then
            'Wide
            ret.Width = dest.Width
            ret.Height = ret.Width / AspectRate
        Else
            'High
            ret.Height = dest.Height
            ret.Width = ret.Height * AspectRate
        End If
        'Left = (dest.Width - ret.Width) / 2
        'Top = (dest.Height - ret.Height) / 2
        Return ret
    End Function
    Public Function ScaleRect(ByVal source As SizeF, ByVal dest As RectangleF, Optional ByVal bStretch As Boolean = True, Optional ByVal bolNoBorder As Boolean = False) As RectangleF
        Return ScaleRect(New RectangleF(PointF.Empty, source), dest, bStretch, bolNoBorder)
    End Function
    Public Function ScaleRect(ByVal source As RectangleF, ByVal dest As RectangleF, Optional ByVal bStretch As Boolean = True, Optional ByVal bolNoBorder As Boolean = False) As RectangleF
        If bStretch = False And source.Width < dest.Width And source.Height < dest.Height Then
            Return New RectangleF((dest.Width / 2) - (source.Width / 2), (dest.Height / 2) - (source.Height / 2), source.Width, source.Height)
        End If
        Dim AspectRate As Single = source.Width / source.Height
        Dim ret As New RectangleF
        If AspectRate > dest.Width / dest.Height Then
            'Wide
            ret.Width = dest.Width
            ret.Height = ret.Width / AspectRate
        Else
            'High
            ret.Height = dest.Height
            ret.Width = ret.Height * AspectRate
        End If
        If Not bolNoBorder Then
            ret.X = dest.X + (dest.Width - ret.Width) / 2
            ret.Y = dest.Y + (dest.Height - ret.Height) / 2
        End If
        ret.Offset(source.Location)
        Return ret
    End Function
    Public Function ScaleImage(ByVal bit As Bitmap, ByVal destWidth As Integer, ByVal destHeight As Integer, Optional ByVal bolNoBorder As Boolean = False, Optional ByVal bolCenterize As Boolean = True) As Bitmap
        'Dim f As Rectangle = Rectangle.Truncate(ScaleRect(bit.Size, New RectangleF(0, 0, destWidth, destHeight)))
        'Dim b As New Bitmap(destWidth, destHeight)
        'Using gr As Graphics = Graphics.FromImage(b)
        '    gr.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        '    gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        '    gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        '    gr.DrawImage(bit, f, New Rectangle(0, 0, bit.Width, bit.Height), GraphicsUnit.Pixel)
        'End Using
        'Return b
        Return ScaleImage(bit, New Size(destWidth, destHeight), bolNoBorder, bolCenterize)
    End Function
    Public Function ScaleImage(ByVal bit As Bitmap, ByVal destSize As Size, Optional ByVal bolNoBorder As Boolean = False, Optional ByVal bolCenterize As Boolean = True) As Bitmap
        'If bit Is Nothing Then Return Nothing

        'Dim f As Rectangle = Rectangle.Ceiling(ScaleRect(bit.Size, New RectangleF(0, 0, destSize.Width, destSize.Height), , bolNoBorder))
        'Dim b As Bitmap
        'If bolNoBorder = False Then
        '    b = New Bitmap(destSize.Width, destSize.Height)
        'Else
        '    If bolCenterize Then
        '        b = New Bitmap(destSize.Width, destSize.Height)
        '        'p = New Point((destSize.Width / 2) - (bit.Width / 2), (destSize.Height / 2) - (bit.Height / 2))
        '        f.Location = New Point((destSize.Width / 2) - (f.Width / 2), (destSize.Height / 2) - (f.Height / 2))
        '    Else
        '        b = New Bitmap(f.Width, f.Height)
        '    End If
        'End If
        'Using gr As Graphics = Graphics.FromImage(b)
        '    'gr.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
        '    'gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        '    'gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        '    SetUpGraphics(gr)
        '    gr.DrawImage(bit, f, New RectangleF(Point.Empty, bit.Size), GraphicsUnit.Pixel)
        'End Using
        'Return b
        Return ScaleImage(bit, destSize, New Rectangle(Point.Empty, bit.Size), bolNoBorder, bolCenterize)
    End Function
    Public Function ScaleImage(ByVal bit As Bitmap, ByVal destSize As Size, ByVal srcRect As Rectangle, Optional ByVal bolNoBorder As Boolean = False, Optional ByVal bolCenterize As Boolean = True) As Bitmap
        If bit Is Nothing Then Return Nothing

        Dim f As Rectangle = Rectangle.Ceiling(ScaleRect(srcRect.Size, New RectangleF(0, 0, destSize.Width, destSize.Height), , bolNoBorder))
        Dim b As Bitmap
        If bolNoBorder = False Then
            b = New Bitmap(destSize.Width, destSize.Height)
        Else
            If bolCenterize Then
                b = New Bitmap(destSize.Width, destSize.Height)
                'p = New Point((destSize.Width / 2) - (bit.Width / 2), (destSize.Height / 2) - (bit.Height / 2))
                f.Location = New Point((destSize.Width / 2) - (f.Width / 2), (destSize.Height / 2) - (f.Height / 2))
            Else
                b = New Bitmap(f.Width, f.Height)
            End If
        End If
        Using gr As Graphics = Graphics.FromImage(b)
            'gr.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            'gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
            'gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            SetUpGraphics(gr)
            gr.DrawImage(bit, f, srcRect, GraphicsUnit.Pixel)
        End Using
        ScaleImage = b.Clone
        b.Dispose()
    End Function
    Public Function FormatSeconds(ByVal o As Object) As String
        If o Is Nothing OrElse o Is DBNull.Value Then Return "0:00"

        Dim ts As TimeSpan = TimeSpan.FromSeconds(o)
        Dim s As New System.Text.StringBuilder
        If ts.Days > 0 Then
            s.Append(ts.Days & ":")
        End If
        If ts.Hours > 0 Then
            s.Append(ts.Hours & ":")
        End If
        s.Append(Format(ts.Minutes, "00") & ":" & Format(ts.Seconds, "00"))
        Return s.ToString
    End Function
    Public Function FormatDate(ByVal d As Object) As String
        If d Is DBNull.Value OrElse d Is Nothing Then Return "Never"
        Return Format(CDate(d), "MM/dd/yy hh:mm tt")
    End Function
    'Public Function FormatBytes(ByVal dwFileSize As Long) As String
    '    Const dwKB As Long = 1024 ';          // Kilobyte

    '    Const dwMB As Long = 1024 * dwKB ';   // Megabyte

    '    Const dwGB As Long = 1024 * dwMB '   // Gigabyte


    '    Dim dwNumber, dwRemainder As Long
    '    Dim strNumber As String = ""

    '    If (dwFileSize < dwKB) Then

    '        strNumber = FormatNumber(dwFileSize, 0, , , TriState.True) + " B"

    '    ElseIf (dwFileSize < dwMB) Then

    '        dwNumber = dwFileSize / dwKB
    '        dwRemainder = (dwFileSize * 100 / dwKB) Mod 100

    '        strNumber = String.Format("{0}.{1:F} KB", FormatNumber(dwNumber, 0, , , TriState.True), dwRemainder)
    '    ElseIf (dwFileSize < dwGB) Then

    '        dwNumber = dwFileSize / dwMB
    '        dwRemainder = (dwFileSize * 100 / dwMB) Mod 100
    '        strNumber = String.Format("{0}.{1:F} MB", FormatNumber(dwNumber, 0, , , TriState.True), dwRemainder)
    '    ElseIf (dwFileSize >= dwGB) Then

    '        dwNumber = dwFileSize / dwGB
    '        dwRemainder = (dwFileSize * 100 / dwGB) Mod 100
    '        strNumber = String.Format("{0}.{1:F} GB", FormatNumber(dwNumber, 0, , , TriState.True), dwRemainder)
    '    End If

    '    '// Display decimal points only if needed

    '    '// another alternative to this approach is to check before calling str.Format, and 

    '    '// have separate cases depending on whether dwRemainder == 0 or not.

    '    strNumber = strNumber.Replace(".00", "")

    '    Return strNumber

    'End Function
    Public Function FormatBytes(ByVal num_bytes As Double) As _
    String
        Const ONE_KB As Double = 1024
        Const ONE_MB As Double = ONE_KB * 1024
        Const ONE_GB As Double = ONE_MB * 1024
        Const ONE_TB As Double = ONE_GB * 1024
        Const ONE_PB As Double = ONE_TB * 1024
        Const ONE_EB As Double = ONE_PB * 1024
        Const ONE_ZB As Double = ONE_EB * 1024
        Const ONE_YB As Double = ONE_ZB * 1024
        'Dim value As Double
        'Dim txt As String

        ' See how big the value is.
        If num_bytes <= 999 Then
            ' Format in bytes.
            FormatBytes = Format$(num_bytes, "0") & " bytes"
        ElseIf num_bytes <= ONE_KB * 999 Then
            ' Format in KB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_KB) & " KB"
        ElseIf num_bytes <= ONE_MB * 999 Then
            ' Format in MB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_MB) & " MB"
        ElseIf num_bytes <= ONE_GB * 999 Then
            ' Format in GB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_GB) & " GB"
        ElseIf num_bytes <= ONE_TB * 999 Then
            ' Format in TB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_TB) & " TB"
        ElseIf num_bytes <= ONE_PB * 999 Then
            ' Format in PB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_PB) & " PB"
        ElseIf num_bytes <= ONE_EB * 999 Then
            ' Format in EB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_EB) & " EB"
        ElseIf num_bytes <= ONE_ZB * 999 Then
            ' Format in ZB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_ZB) & " ZB"
        Else
            ' Format in YB.
            FormatBytes = ThreeNonZeroDigits(num_bytes /
                ONE_YB) & " YB"
        End If
    End Function
    Private Function ThreeNonZeroDigits(ByVal value As Double) _
        As String
        If value >= 100 Then
            ' No digits after the decimal.
            ThreeNonZeroDigits = Format$(CInt(value))
        ElseIf value >= 10 Then
            ' One digit after the decimal.
            ThreeNonZeroDigits = Format$(value, "0.0")
        Else
            ' Two digits after the decimal.
            ThreeNonZeroDigits = Format$(value, "0.00")
        End If
    End Function

    Public Function GetDrawMargins(ByVal srcSize As SizeF, ByVal margins As RectangleF, ByVal destRect As RectangleF) As RectangleF(,)
        Dim r(1, 8) As RectangleF
        r(0, 0) = New RectangleF(destRect.X + 0, destRect.Y + 0, margins.X, margins.Y)
        r(1, 0) = New RectangleF(0, 0, margins.X, margins.Y)

        r(0, 1) = New RectangleF(destRect.Right - margins.Width, destRect.Y + 0, margins.Width, margins.Y)
        r(1, 1) = New RectangleF(srcSize.Width - margins.Width, 0, margins.Width, margins.Y)

        r(0, 2) = New RectangleF(destRect.X, destRect.Bottom - margins.Height, margins.X, margins.Height)
        r(1, 2) = New RectangleF(0, srcSize.Height - margins.Height, margins.X, margins.Height)

        r(0, 3) = New RectangleF(destRect.Right - margins.Width, destRect.Bottom - margins.Height, margins.Width, margins.Height)
        r(1, 3) = New RectangleF(srcSize.Width - margins.Width, srcSize.Height - margins.Height, margins.Width, margins.Height)


        r(0, 4) = New RectangleF(destRect.X, destRect.Y + margins.Y, margins.X, destRect.Height - margins.Bottom)
        r(1, 4) = New RectangleF(0, margins.Y, margins.X, srcSize.Height - margins.Bottom)

        r(0, 5) = New RectangleF(destRect.X + margins.X, destRect.Y, destRect.Width - margins.Right, margins.Y)
        r(1, 5) = New RectangleF(margins.X, 0, srcSize.Width - margins.Right, margins.Y)

        r(0, 6) = New RectangleF(destRect.X + margins.X, destRect.Bottom - margins.Height, destRect.Width - margins.Right, margins.Height)
        r(1, 6) = New RectangleF(margins.X, srcSize.Height - margins.Height, srcSize.Width - margins.Right, margins.Height)

        r(0, 7) = New RectangleF(destRect.Right - margins.Width, destRect.Y + margins.Y, margins.Width, destRect.Height - margins.Bottom)
        r(1, 7) = New RectangleF(srcSize.Width - margins.Width, margins.Y, margins.Width, srcSize.Height - margins.Bottom)



        r(0, 8) = New RectangleF(destRect.X + margins.X, destRect.Y + margins.Y, destRect.Width - margins.Right, destRect.Height - margins.Bottom)
        r(1, 8) = New RectangleF(margins.X, margins.Y, srcSize.Width - margins.Right, srcSize.Height - margins.Bottom)
        Return r
    End Function
    Public Sub DrawMargins(ByRef g As Graphics, ByVal b As Bitmap, ByVal margins(,) As RectangleF)
        With g
            For t As Integer = 0 To 8
                .DrawImage(b, margins(0, t), margins(1, t), GraphicsUnit.Pixel)
            Next
        End With

    End Sub
    Public Sub DrawMargins(ByRef g As Graphics, ByVal b As Bitmap, ByVal margins As RectangleF, ByVal destRect As RectangleF)
        DrawMargins(g, b, GetDrawMargins(b.Size, margins, destRect))
    End Sub

    Public Function ExpandPath(ByVal strTargetPath As String, ByVal strRelativePath As String) As String
        Try
            Dim s() As String = Split(strRelativePath, ";")
            If s Is Nothing Then ReDim s(0) : s(0) = strRelativePath
            If IO.Path.IsPathRooted(s(0)) = False Then

                Return IO.Path.GetFullPath(IO.Path.Combine(IO.Path.GetDirectoryName(strTargetPath), s(0)))
            Else
                Return s(0)
            End If
        Catch ex As ArgumentException
            Return Nothing
        End Try
    End Function
    Public Function ExpandPath(ByVal strTargetPath As String, ByVal strRelativePaths() As String) As String()
        Dim d(UBound(strRelativePaths)) As String
        For t As Integer = 0 To UBound(strRelativePaths)
            d(t) = ExpandPath(strTargetPath, strRelativePaths(t))
        Next
        Return d
    End Function
    Public Function CollapsePath(ByVal strTargetPath As String, ByVal strPath As String) As String
        Dim i As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strPath)) & "\"
        Dim i2 As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strTargetPath)) & "\"

        Return i.Replace(i2, "") & IO.Path.GetFileName(strPath)
    End Function
    Public Function CollapsePath(ByVal strTargetPath As String, ByVal strPaths As String()) As String()
        'Dim i As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strPath)) & "\"
        'Dim i2 As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strMamePath)) & "\"

        'Return i.Replace(i2, "") & IO.Path.GetFileName(strPath)
        Dim d(UBound(strPaths)) As String
        For t As Integer = 0 To UBound(strPaths)
            d(t) = CollapsePath(strTargetPath, strPaths(t))
        Next
        Return d
    End Function
    Public Function CollapseFolder(ByVal strTarget As String, ByVal strPath As String) As String
        If Not IO.Path.IsPathRooted(strPath) Then
            strPath = IO.Path.Combine(strTarget, strPath)
        End If
        Dim i As String = (IO.Path.GetFullPath(strPath)) '& "\"
        Dim i2 As String = (IO.Path.GetFullPath(strTarget)) '& "\"
        If i = i2 Then Return "."
        Return i.Replace(i2 & "\", "") '& IO.Path.GetFileName(strPath)
    End Function
    Public Function CollapseFolder(ByVal strTarget As String, ByVal strPaths() As String) As String()
        Dim d(UBound(strPaths)) As String
        For t As Integer = 0 To UBound(strPaths)
            d(t) = CollapseFolder(strTarget, strPaths(t))
        Next
        Return d
    End Function


    Private Delegate Function ThreadSafeDelegate(ByVal c As Windows.Forms.Control, ByVal target As Object, ByVal ProcName As String, ByVal CallType As CallType, ByVal value As Object) As Object
    ''' <summary>
    ''' Makes a recursive call (if necessary) to any control's public members when executed in another thread.
    ''' </summary>
    ''' <param name="c">Control to query</param>
    ''' <param name="ProcName">The procedure to execute. This can be a property or variable as long as it is declared as public</param>
    ''' <param name="CallType"></param>
    ''' <param name="value">The value used to assign the procedure</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function MakeThreadSafeCall(ByVal c As System.Windows.Forms.Control, ByVal ProcName As String, ByVal CallType As CallType, ByVal value As Object) As Object
        If c.InvokeRequired Then
            Dim d As New ThreadSafeDelegate(AddressOf MakeThreadSafeCall)
            c.Invoke(d, c, Nothing, ProcName, CallType, value)

            Return Nothing
        Else
            Return CallByName(c, ProcName, CallType, value)
        End If
    End Function
    ''' <summary>
    ''' Makes a recursive call (if necessary) to any control's public members when executed in another thread.
    ''' </summary>
    ''' <param name="c">The targets parent control</param>
    ''' <param name="target">Any object that isn't a Control</param>
    ''' <param name="ProcName">The procedure to execute. This can be a property or variable as long as it is declared as public</param>
    ''' <param name="CallType"></param>
    ''' <param name="value">The value used to assign the procedure</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function MakeThreadSafeCall(ByVal c As Windows.Forms.Control, ByVal target As Object, ByVal ProcName As String, ByVal CallType As CallType, ByVal value As Object) As Object
        If c.InvokeRequired Then
            Dim d As New ThreadSafeDelegate(AddressOf MakeThreadSafeCall)
            c.Invoke(d, c, target, ProcName, CallType, value)
            Return Nothing
        Else
            If target Is Nothing Then
                Return CallByName(c, ProcName, CallType, value)
            Else
                Return CallByName(target, ProcName, CallType, value)
            End If

        End If
    End Function


    'Public Function UnzipStream(ByVal strZipFile As String, ByVal extensions() As String, ByVal HeaderSize As Integer) As Byte()
    '    Return UnzipStream(strZipFile, extensions, HeaderSize, 0)
    'End Function
    Public Function UnzipStream(ByVal strZipFile As String, ByVal extensions As String(), ByVal HeaderSize As Integer, ByVal offset As Integer) As Byte()


        Dim si As New IO.FileStream(strZipFile, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
        Dim i As Ionic.Zip.ZipFile '(strZipFile)
        i = Ionic.Zip.ZipFile.Read(si)
        Dim z As Integer = -1
        For t As Integer = 0 To i.Count - 1
            If Array.IndexOf(extensions, IO.Path.GetExtension(i.Item(t).FileName).ToLower) >= 0 Then
                z = t
                Exit For
            End If
        Next
        If z = -1 Then Return Nothing
        Dim a(HeaderSize - 1) As Byte
        'If IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strZipFile), i(z).FileName)) Then
        '    i.Dispose()
        '    si.Close()
        '    si.Dispose()
        '    si = New IO.FileStream(IO.Path.Combine(IO.Path.GetDirectoryName(strZipFile), i(z).FileName), IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
        '    si.Position = offset
        '    si.Read(a, 0, HeaderSize - 1)
        '    si.Close()
        '    si.Dispose()
        'Else
        Try
            i.Item(0).OpenReader.Read(a, offset, HeaderSize - 1)
            'i.Item(z).Extract(si)
            'si.Position = offset
            'si.Read(a, 0, HeaderSize - 1)
        Catch ex As Exception
            Erase a
        Finally
            i.Dispose()
            si.Close()
            si.Dispose()
        End Try
        'End If
        Return a

    End Function
    Public Function UnzipStream(ByVal strZipFile As String, ByVal extensions As String()) As IO.Stream
        'Dim si As New IO.MemoryStream()
        Dim si As New IO.FileStream(strZipFile, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)

        'si.SetLength(headersize)
        Dim i As Ionic.Zip.ZipFile '(strZipFile)
        i = Ionic.Zip.ZipFile.Read(si)
        Dim z As Integer = -1
        For t As Integer = 0 To i.Count - 1
            If Array.IndexOf(extensions, IO.Path.GetExtension(i.Item(t).FileName).ToLower) >= 0 Then
                z = t
                Exit For
            End If
        Next
        If z = -1 Then Return Nothing
        'i.Item(z).Extract(si)


        'UnzipStream = si.ToArray
        'i.Dispose()
        'si.Dispose()
        UnzipStream = New IO.StreamReader(i.Item(z).OpenReader).BaseStream
        i.Dispose()
    End Function
    Public Sub Unzip(ByVal strZipFile As String, Optional ByVal bOverwrite As Boolean = True)
        Dim i As New Ionic.Zip.ZipFile(strZipFile)
        i.ExtractAll(IO.Path.GetDirectoryName(strZipFile), IIf(bOverwrite, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently, Ionic.Zip.ExtractExistingFileAction.DontOverwrite))
        i.Dispose()
    End Sub
    Public Sub Unzip(ByVal strZipFile As String, ByVal targetDir As String, Optional ByVal bOverwrite As Boolean = True)
        Dim i As New Ionic.Zip.ZipFile(strZipFile)
        i.ExtractAll(targetDir, IIf(bOverwrite, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently, Ionic.Zip.ExtractExistingFileAction.DontOverwrite))
        i.Dispose()
    End Sub
    Public Sub Unzip(ByVal strZipFile As String, ByVal targetDir As String, ByVal targetFile As String, Optional ByVal bOverwrite As Boolean = True)
        Dim i As New Ionic.Zip.ZipFile(strZipFile)
        i.Item(targetFile).Extract(IO.Path.GetTempPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
        Dim file As String = IO.Path.Combine(IO.Path.GetTempPath, targetFile)
        IO.File.Copy(file, IO.Path.Combine(targetDir, IO.Path.GetFileName(targetFile)), bOverwrite)
        IO.File.Delete(file)
        i.Dispose()
    End Sub
    'Private Sub FileExists(ByVal sender As Object, ByVal e As SevenZip.FileOverwriteEventArgs)
    '    IO.File.Delete(e.FileName)
    '    e.Cancel = False
    'End Sub
    Public Sub SevenZipExtract(ByVal s7ZipFile As String, ByVal sDestFolder As String)


        'Dim p As New Process
        'With p.StartInfo
        '    .Arguments = "x """ & s7ZipFile & """-o""" & sDestFolder & """"
        '    .CreateNoWindow = True
        '    .FileName = "7za.exe"
        '    '.WorkingDirectory= 
        'End With
        'p.Start()
        'p.WaitForExit()
        'p.Dispose()
        'string sourceName = "ExampleText.txt";
        '    string targetName = "Example.gz";

        '// 1
        '// Initialize process information.
        '//
        Dim p As New ProcessStartInfo()
        p.FileName = "7za.exe"

        '// 2
        '// Use 7-zip
        '// specify a=archive and -tgzip=gzip
        '// and then target file in quotes followed by source file in quotes
        '//
        p.Arguments = "a -tgzip """ & s7ZipFile & """ """ & sDestFolder & """ -mx=9"
        p.WindowStyle = ProcessWindowStyle.Hidden
        p.UseShellExecute = False
        '// 3.
        '// Start process and wait for it to exit
        '//
        Dim x As New Process
        AddHandler x.OutputDataReceived, AddressOf Temp
        p.RedirectStandardOutput = True
        x.EnableRaisingEvents = True
        x = Process.Start(p)

        x.WaitForExit()
        RemoveHandler x.OutputDataReceived, AddressOf Temp
        x.Dispose()
    End Sub

    Private Sub Temp(ByVal sender As Object, ByVal e As Diagnostics.DataReceivedEventArgs)

    End Sub


    Public Function SevenZipList(ByVal s7ZipFile As String) As Collections.ObjectModel.ReadOnlyCollection(Of String)
        'Dim s As New SevenZip.SevenZipExtractor(s7ZipFile)
        'SevenZipList = s.ArchiveFileNames
        's.Dispose()
        'Exit Function

        Dim p As New ProcessStartInfo()
        p.FileName = "7za.exe"

        '// 2
        '// Use 7-zip
        '// specify a=archive and -tgzip=gzip
        '// and then target file in quotes followed by source file in quotes
        '//
        p.Arguments = "l """ & s7ZipFile & """"
        'p.WindowStyle = ProcessWindowStyle.Hidden
        p.UseShellExecute = False
        p.CreateNoWindow = True
        p.RedirectStandardOutput = True
        '// 3.
        '// Start process and wait for it to exit
        '//
        Dim x As New Process
        'x.EnableRaisingEvents = True

        Dim Lines() As String = Nothing
        Dim tmp As String = ""
        x = Process.Start(p)
        Dim sr As New IO.StreamReader(x.StandardOutput.BaseStream)

        Do Until sr.EndOfStream
            Dim c As Char = ChrW(sr.Read)
            Select Case c
                Case vbLf, vbCr
                    'endit
                    If String.IsNullOrEmpty(tmp) Then Continue Do
                    If Lines Is Nothing Then ReDim Lines(0) Else ReDim Preserve Lines(UBound(Lines) + 1)
                    Lines(UBound(Lines)) = tmp
                    tmp = ""
                Case Else
                    tmp &= c
            End Select
        Loop
        x.Dispose()
        '0=7zip info
        '1="listing archive..
        '2=Method
        '3=Solid
        '4=Blocks
        '5=Physical size
        '6=header size
        '7=offset
        '8=table heading
        '9=table sep (---------)

        Dim l As New List(Of String)
        For t As Integer = 10 To UBound(Lines) - 2
            Try
                Dim q As New SevenZipContents(Lines(t))
                l.Add(q.Name)
            Catch
            End Try
        Next
        Return New Collections.ObjectModel.ReadOnlyCollection(Of String)(l)
    End Function

    Public Structure SevenZipContents
        Public [Date] As DateTime
        Public ReadOnly Property IsDirectory() As Boolean
            Get
                Return (Attributes And IO.FileAttributes.Directory)
            End Get
        End Property
        Public Attributes As IO.FileAttributes
        Public Size As Long
        Public CompressedSize As Long
        Public Name As String

        Friend Sub New(ByVal strIn As String)
            Dim i As Integer = InStr(strIn, " ")
            i = InStr(i + 1, strIn, " ")
            [Date] = "#" & strIn.Substring(0, i) & "#"
            'i = InStr(i + 1, strIn, " ")
            For z As Integer = 0 To 4
                Select Case strIn(i + z)
                    Case "D"
                        Attributes = Attributes Or IO.FileAttributes.Directory
                    Case "A"
                        Attributes = Attributes Or IO.FileAttributes.Archive
                    Case "R"
                        Attributes = Attributes Or IO.FileAttributes.ReadOnly
                    Case "H"
                        Attributes = Attributes Or IO.FileAttributes.Hidden
                    Case "C"
                        Attributes = Attributes Or IO.FileAttributes.Compressed
                End Select
            Next
            i += 6
            Dim tmp As String

            tmp = strIn.Substring(i, 12)
            Me.Size = tmp
            i += 12 + 1
            tmp = strIn.Substring(i, 12)
            Me.CompressedSize = Val(tmp)
            i += 12 + 1
            Me.Name = Mid(strIn, i).Trim
        End Sub
    End Structure

    Public Function HiNibble(ByVal b As Byte) As Byte
        Return b >> 4
    End Function
    Public Function LoNibble(ByVal b As Byte) As Byte
        Return b And &HF
    End Function


    '///
    '/// Gets whether or not a given Bitmap is blank.
    '///
    '/// The instance of the Bitmap for this method extension.
    '/// Returns trueif the given Bitmap is blank; otherwise returns false.
    Public Function IsBlank(ByVal bitmap As Bitmap) As Boolean
        Dim stdDev As Double = GetStdDev(bitmap)
        Dim tolerance As Integer = 100000
        Return stdDev < tolerance
    End Function

    ''' Gets the bits per pixel (bpp) for the given .
    '''
    ''' The instance of the for this method extension.
    ''' Returns a representing the bpp for the .
    Friend Function GetBitsPerPixel(ByVal Bitmap As Bitmap) As Byte
        Dim bpp As Byte = &H1

        '//return Regex.Match(Regex.Match(bitmap.PixelFormat.ToString(), @”\dbpp”).Value, @”\d+”).Value;
        Select Case Bitmap.PixelFormat
            Case Imaging.PixelFormat.Format1bppIndexed
                bpp = 1

            Case Imaging.PixelFormat.Format4bppIndexed
                bpp = 4

            Case Imaging.PixelFormat.Format8bppIndexed
                bpp = 8
            Case Imaging.PixelFormat.Format16bppArgb1555,
                Imaging.PixelFormat.Format16bppGrayScale,
                Imaging.PixelFormat.Format16bppRgb555,
                Imaging.PixelFormat.Format16bppRgb565
                bpp = 16
            Case Imaging.PixelFormat.Format24bppRgb
                bpp = 24

            Case Imaging.PixelFormat.Canonical,
                Imaging.PixelFormat.Format32bppArgb,
                Imaging.PixelFormat.Format32bppPArgb,
                Imaging.PixelFormat.Format32bppRgb
                bpp = 32

            Case Imaging.PixelFormat.Format48bppRgb
                bpp = 48
            Case Imaging.PixelFormat.Format64bppArgb, Imaging.PixelFormat.Format64bppPArgb
                bpp = 64

        End Select
        Return bpp
    End Function


    ''' Get the standard deviation of pixel values.
    '''
    ''' The instance of the for this method extension.
    ''' Returns the standard deviation of pixel population of the Bitmap.
    Public Function GetStdDev(ByVal bitmap As Bitmap) As Double
        Dim total As Double = 0
        Dim totalVariance As Double = 0
        Dim count As Integer = 0
        Dim stdDev As Double = 0

        '// First get all the bytes
        Dim bmData As Imaging.BitmapData = bitmap.LockBits(New Rectangle(0, 0, bitmap.Width, bitmap.Height), Imaging.ImageLockMode.ReadOnly, bitmap.PixelFormat)
        Dim stride As Integer = bmData.Stride
        Dim Scan0 As IntPtr = bmData.Scan0

        Dim bitsPerPixel As Byte = GetBitsPerPixel(bitmap)
        Dim bytesPerPixel As Byte = CByte(bitsPerPixel / 8)

        'unsafe {
        'byte* p = (byte*)(void*)Scan0;
        Dim p(bitmap.Width * bitmap.Height) As Byte
        Marshal.Copy(Scan0, p, 0, bitmap.Width * bitmap.Height)
        Dim nOffset As Integer = stride - bitmap.Width * bytesPerPixel
        Dim intCount As Integer = 0
        For y As Integer = 0 To bitmap.Height - 1
            For x As Integer = 0 To bitmap.Width - 1
                count += 1
                If intCount + 2 > UBound(p) Then Exit For
                Dim blue As Byte = p(intCount)
                Dim green As Byte = p(intCount + 1)
                Dim red As Byte = p(intCount + 2)

                Dim pixelValue As Integer = Color.FromArgb(0, red, green, blue).ToArgb()
                total += pixelValue
                Dim avg As Double = total / count
                totalVariance += Math.Pow(pixelValue - avg, 2)
                stdDev = Math.Sqrt(totalVariance / count)
                'p += bytesPerPixel
                intCount = intCount + bytesPerPixel
            Next
            'p += nOffset
            intCount = intCount + nOffset
        Next

        bitmap.UnlockBits(bmData)

        Return stdDev
    End Function

    Private Function XorImageSize(ByVal bmInfoHeader As Win32.Gdi32.BITMAPINFOHEADER) As Integer


        Return (bmInfoHeader.biHeight / 2 * WidthBytes(bmInfoHeader.biWidth * bmInfoHeader.biBitCount * bmInfoHeader.biPlanes))
    End Function
    Private Function WidthBytes(ByVal width As Integer) As Integer

        '// Returns the width of a row in a DIB Bitmap given the
        '// number of bits.  DIB Bitmap rows always align on a 
        '// DWORD boundary.
        Return ((width + 31) / 32) * 4
    End Function
    Private Function XorImageIndex(ByVal bmInfoHeader As Win32.Gdi32.BITMAPINFOHEADER) As Integer

        '// Returns the position of the DIB bitmap bits within a
        '// DIB bitmap array:
        Return Marshal.SizeOf(bmInfoHeader) + dibNumColors(bmInfoHeader) * Marshal.SizeOf(GetType(Win32.Gdi32.RGBQUAD))
    End Function
    Private Function dibNumColors(ByVal bmInfoHeader As Win32.Gdi32.BITMAPINFOHEADER) As Integer

        Dim colorCount As Integer = 0
        If (bmInfoHeader.biClrUsed <> 0) Then

            colorCount = bmInfoHeader.biClrUsed

        Else

            Select Case (bmInfoHeader.biBitCount)

                Case 1
                    colorCount = 2
                Case 4
                    colorCount = 16
                Case 8
                    colorCount = 256

            End Select
        End If
        Return colorCount
    End Function
    Private Function getIconBitmap(ByVal Data As Byte(),
           ByVal mask As Boolean,
           ByVal returnHandle As Boolean,
           ByRef hBmp As IntPtr
           ) As Bitmap

        '// Bitmap to return
        Dim bm As Bitmap = Nothing

        '// Get bitmap info:         
        Dim bmInfoHdr As New Win32.Gdi32.BITMAPINFOHEADER(Data)

        If (mask) Then

            '// extract monochrome mask
            Dim hdc As IntPtr = Win32.Gdi32.CreateCompatibleDC(IntPtr.Zero)
            hBmp = Win32.Gdi32.CreateCompatibleBitmap(hdc, bmInfoHdr.biWidth,
             bmInfoHdr.biHeight / 2)
            Dim hBmpOld As IntPtr = Win32.Gdi32.SelectObject(hdc, hBmp)

            '// Prepare BitmapInfoHeader for mono bitmap:
            Dim rgbQuad As New Win32.Gdi32.RGBQUAD()
            Dim monoBmHdrSize As Integer = bmInfoHdr.biSize + Marshal.SizeOf(rgbQuad) * 2

            Dim bitsInfo As IntPtr = Marshal.AllocCoTaskMem(monoBmHdrSize)
            Marshal.WriteInt32(bitsInfo, Marshal.SizeOf(bmInfoHdr))
            Marshal.WriteInt32(bitsInfo, 4, bmInfoHdr.biWidth)
            Marshal.WriteInt32(bitsInfo, 8, CInt(bmInfoHdr.biHeight / 2))
            Marshal.WriteInt16(bitsInfo, 12, 1)
            Marshal.WriteInt16(bitsInfo, 14, 1)
            Marshal.WriteInt32(bitsInfo, 16, Win32.Gdi32.BI_RGB)
            Marshal.WriteInt32(bitsInfo, 20, 0)
            Marshal.WriteInt32(bitsInfo, 24, 0)
            Marshal.WriteInt32(bitsInfo, 28, 0)
            Marshal.WriteInt32(bitsInfo, 32, 0)
            Marshal.WriteInt32(bitsInfo, 36, 0)
            '// Write the black and white colour indices:
            Marshal.WriteInt32(bitsInfo, 40, 0)
            Marshal.WriteByte(bitsInfo, 44, 255)
            Marshal.WriteByte(bitsInfo, 45, 255)
            Marshal.WriteByte(bitsInfo, 46, 255)
            Marshal.WriteByte(bitsInfo, 47, 0)

            '// Prepare Mask bits:

            Dim maskImageBytes As Integer = bmInfoHdr.biHeight / 2 *
         WidthBytes(bmInfoHdr.biWidth)
            Dim bits As IntPtr = Marshal.AllocCoTaskMem(maskImageBytes)
            Dim _maskImageIndex As Integer = XorImageIndex(bmInfoHdr)
            _maskImageIndex += XorImageSize(bmInfoHdr)
            Marshal.Copy(Data, _maskImageIndex, bits, maskImageBytes)

            Dim success As Integer = Win32.Gdi32.SetDIBitsToDevice(
               hdc,
               0, 0, bmInfoHdr.biWidth, bmInfoHdr.biHeight / 2,
               0, 0, 0, bmInfoHdr.biHeight / 2,
               bits,
               bitsInfo,
             Win32.Gdi32.DIB_RGB_COLORS)

            Marshal.FreeCoTaskMem(bits)
            Marshal.FreeCoTaskMem(bitsInfo)

            Win32.Gdi32.SelectObject(hdc, hBmpOld)
            Win32.Gdi32.DeleteObject(hdc)


        Else

            '// extract colour (XOR) part of image:

            '// Create bitmap:
            Dim hdcDesktop As IntPtr = Win32.Gdi32.CreateDC("DISPLAY", IntPtr.Zero, IntPtr.Zero,
             IntPtr.Zero)
            Dim hdc As IntPtr = Win32.Gdi32.CreateCompatibleDC(hdcDesktop)
            hBmp = Win32.Gdi32.CreateCompatibleBitmap(hdcDesktop, bmInfoHdr.biWidth,
             bmInfoHdr.biHeight / 2)
            Win32.Gdi32.DeleteDC(hdcDesktop)
            Dim hBmpOld As IntPtr = Win32.Gdi32.SelectObject(hdc, hBmp)

            '            // Find the index of the XOR bytes:
            Dim xorIndex As Integer = XorImageIndex(bmInfoHdr)
            Dim _xorImageSize As Integer = XorImageSize(bmInfoHdr)

            '// Get Bitmap info header to a pointer:                        
            Dim bitsInfo As IntPtr = Marshal.AllocCoTaskMem(xorIndex)
            Marshal.Copy(Data, 0, bitsInfo, xorIndex)
            '// fix the height:
            Marshal.WriteInt32(bitsInfo, 8, CInt(bmInfoHdr.biHeight / 2))

            '// Get the XOR bits:            
            Dim bits As IntPtr = Marshal.AllocCoTaskMem(_xorImageSize)
            Marshal.Copy(Data, xorIndex, bits, _xorImageSize)

            Dim success As Integer = Win32.Gdi32.SetDIBitsToDevice(
               hdc,
               0, 0, bmInfoHdr.biWidth, bmInfoHdr.biHeight / 2,
               0, 0, 0, bmInfoHdr.biHeight / 2,
               bits,
               bitsInfo,
              Win32.Gdi32.DIB_RGB_COLORS)

            Marshal.FreeCoTaskMem(bits)
            Marshal.FreeCoTaskMem(bitsInfo)

            Win32.Gdi32.SelectObject(hdc, hBmpOld)
            Win32.Gdi32.DeleteObject(hdc)
        End If

        If (Not returnHandle) Then

            '// the bitmap will own the handle and clear
            '// it up when it is disposed.  Otherwise
            '// need to call DeleteObject on hBmp
            '// returned.
            bm = Bitmap.FromHbitmap(hBmp)
        End If
        Return bm
    End Function
    Public Function BitmapToIcon(ByVal bm As Bitmap) As Icon
        'new
        'setdevice
        '// Initialise the data:
        Dim size As Size = bm.Size
        Dim bmInfoHdr As New Win32.Gdi32.BITMAPINFOHEADER(size, Windows.Forms.ColorDepth.Depth32Bit)
        Dim hIcon As IntPtr
        Dim _maskImageIndex As Integer = XorImageIndex(bmInfoHdr)
        _maskImageIndex += XorImageSize(bmInfoHdr)
        Dim _MaskImageSize As Integer = bmInfoHdr.biHeight / 2 *
         WidthBytes(bmInfoHdr.biWidth)
        Dim data(_maskImageIndex + _MaskImageSize) As Byte

        Dim mw As New IO.MemoryStream(data, 0, data.Length, True)
        Dim bw As New IO.BinaryWriter(mw)
        bmInfoHdr.Write(bw)
        bw.Close()

        'seticon
        Dim hdcc As IntPtr = Win32.Gdi32.CreateDC("DISPLAY", IntPtr.Zero, IntPtr.Zero,
                IntPtr.Zero)
        Dim hdc As IntPtr = Win32.Gdi32.CreateCompatibleDC(hdcc)
        Win32.Gdi32.DeleteDC(hdcc)
        Dim hBmp As IntPtr = bm.GetHbitmap()

        '// Now prepare for GetDIBits call:
        Dim xorIndex As Integer = XorImageIndex(bmInfoHdr)
        Dim xorImageBytes As Integer = XorImageSize(bmInfoHdr)

        '// Get the BITMAPINFO header into the pointer:
        Dim bitsInfo As IntPtr = Marshal.AllocCoTaskMem(xorIndex)
        Marshal.Copy(data, 0, bitsInfo, xorIndex)
        '// fix the height:
        Marshal.WriteInt32(bitsInfo, 8, CInt(bmInfoHdr.biHeight / 2))

        Dim bits As IntPtr = Marshal.AllocCoTaskMem(xorImageBytes)

        Dim success As Integer = Win32.Gdi32.GetDIBits(hdc, hBmp, 0, size.Height, bits,
         bitsInfo, Win32.Gdi32.DIB_RGB_COLORS)

        Marshal.Copy(bits, data, xorIndex, xorImageBytes)

        '// Free memory:
        Marshal.FreeCoTaskMem(bits)
        Marshal.FreeCoTaskMem(bitsInfo)

        Win32.Gdi32.DeleteObject(hBmp)
        Win32.Gdi32.DeleteDC(hdc)

        'createIcon()
        If (hIcon <> IntPtr.Zero) Then

            Win32.User32.DestroyIcon(hIcon)
            hIcon = IntPtr.Zero
        End If

        Dim ii As New Win32.User32.ICONINFO()
        ii.fIcon = Win32.Gdi32.IMAGE_ICON
        getIconBitmap(data, False, True, ii.hBmColor)
        getIconBitmap(data, True, True, ii.hBmMask)

        hIcon = Win32.User32.CreateIconIndirect(ii)

        Win32.Gdi32.DeleteObject(ii.hBmColor)
        Win32.Gdi32.DeleteObject(ii.hBmMask)
        'icon
        Dim i As System.Drawing.Icon =
                System.Drawing.Icon.FromHandle(hIcon)
        BitmapToIcon = i.Clone
        i.Dispose()
        'dispose
        Win32.User32.DestroyIcon(hIcon)
        hIcon = IntPtr.Zero
    End Function

End Module

Public Class ManagedThread
    Implements IDisposable

    'Public MustInherit Class Worker
    'MustOverride Sub DoWork(ByVal o As Object)
    'MustOverride Sub CleanUp()
    Dim context As SynchronizationContext
    Public Sub New()
        context = SynchronizationContext.Current.CreateCopy

    End Sub
    'Public ReadOnly Property ThreadContext() As SynchronizationContext
    '    Get
    '        Return _threadContext
    '    End Get
    'End Property
    Public ReadOnly Property CallerContext() As SynchronizationContext
        Get
            Return context
        End Get
    End Property
    Private Events As New System.ComponentModel.EventHandlerList

    Public Event WorkProgress As SendOrPostCallback
    Public Custom Event AbortProgress As SendOrPostCallback
        AddHandler(ByVal value As SendOrPostCallback)
            Events.AddHandler("AbortProgress", value)
        End AddHandler

        RemoveHandler(ByVal value As SendOrPostCallback)
            Events.RemoveHandler("AbortProgress", value)
        End RemoveHandler

        RaiseEvent(ByVal state As Object)
            Try
                If Events("AbortProgress") IsNot Nothing Then context.Send(Events("AbortProgress"), state)
            Catch
            End Try
        End RaiseEvent
    End Event
    'Public Event Starting As EventHandler
    'Public Event Started As EventHandler
    'Public Event Finished As SendOrPostCallback
    '    AddHandler(ByVal value As SendOrPostCallback)
    '        Events.AddHandler("Finished", value)
    '    End AddHandler

    '    RemoveHandler(ByVal value As SendOrPostCallback)
    '        Events.RemoveHandler("Finished", value)
    '    End RemoveHandler

    '    RaiseEvent(ByVal state As Object)
    '        context.Send(Events("Finished"), state)
    '    End RaiseEvent
    'End Event
    'Public Event DoWork(ByVal sender As Object, ByVal state As Object)
    Public Custom Event CleanUp As SendOrPostCallback
        AddHandler(ByVal value As SendOrPostCallback)
            Events.AddHandler("CleanUp", value)
        End AddHandler

        RemoveHandler(ByVal value As SendOrPostCallback)
            Events.RemoveHandler("CleanUp", value)
        End RemoveHandler

        RaiseEvent(ByVal state As Object)
            Try
                If Events("CleanUp") IsNot Nothing Then context.Send(Events("CleanUp"), state)
            Catch ex As Exception

            End Try
        End RaiseEvent
    End Event
    '    AddHandler(ByVal value As SendOrPostCallback)
    '        Events.AddHandler("CleanUp", value)
    '    End AddHandler

    '    RemoveHandler(ByVal value As SendOrPostCallback)
    '        Events.RemoveHandler("CleanUp", value)
    '    End RemoveHandler

    '    RaiseEvent(ByVal state As Object)
    '        context.Send(Events("CleanUp"), state)
    '    End RaiseEvent
    'End Event
    'Public Event AbortStarted As SendOrPostCallback
    '    AddHandler(ByVal value As SendOrPostCallback)
    '        Events.AddHandler("AbortStarted", value)
    '    End AddHandler

    '    RemoveHandler(ByVal value As SendOrPostCallback)
    '        Events.RemoveHandler("AbortStarted", value)
    '    End RemoveHandler

    '    RaiseEvent(ByVal state As Object)
    ''ReturnToCallerThread(CType(Events("ClickEvent"), EventHandler), Nothing)
    '        context.Send(Events("AbortStarted"), state)
    '    End RaiseEvent
    'End Event
    Public Custom Event OnError As SendOrPostCallback
        AddHandler(ByVal value As SendOrPostCallback)
            Events.AddHandler("OnError", value)
        End AddHandler

        RemoveHandler(ByVal value As SendOrPostCallback)
            Events.RemoveHandler("OnError", value)
        End RemoveHandler

        RaiseEvent(ByVal state As Object)
            Try
                If Events("OnError") IsNot Nothing Then context.Send(Events("OnError"), state)
            Catch
            End Try
        End RaiseEvent
    End Event

    Public Enum States
        Undefined
        Starting
        Started
        CleaningUp
        Aborting
        Stopping
        Stopped
    End Enum
    Dim _state As States
    Dim stateLock As New Object
    Public Property State() As States
        Get
            Return _state
        End Get
        Friend Set(ByVal value As States)
            Monitor.Enter(stateLock)
            _state = value
            Monitor.Exit(stateLock)
            RaiseEvent StateChanged(value)
        End Set
    End Property
    Dim _tp As ThreadPriority = Threading.ThreadPriority.Normal
    Public Property ThreadPriority() As ThreadPriority
        Get
            Return _tp
        End Get
        Set(ByVal value As ThreadPriority)
            _tp = value
        End Set
    End Property
    Public Custom Event StateChanged As Threading.SendOrPostCallback
        AddHandler(ByVal value As Threading.SendOrPostCallback)
            Events.AddHandler("StateChanged", value)
        End AddHandler

        RemoveHandler(ByVal value As Threading.SendOrPostCallback)
            Events.RemoveHandler("StateChanged", value)
        End RemoveHandler

        RaiseEvent(ByVal state As Object)
            Try
                If Events("StateChanged") IsNot Nothing Then context.Send(Events("StateChanged"), state)
            Catch
            End Try
            'DirectCast(Events("StateChanged"), SendOrPostCallback).Invoke(state)
        End RaiseEvent
    End Event
    'End Class
    Dim t As Thread '(AddressOf DoWork)
    Dim wait As New AutoResetEvent(False)
    Dim p, p2 As New AutoResetEvent(False)
    Dim t2 As Thread
    Dim bAbort As Boolean
    Public Function Start(ByVal d As ParameterizedThreadStart, ByVal state As Object) As Boolean
        'RaiseEvent Starting(Me, Nothing)
        If _state = States.Undefined OrElse _state = States.Stopped Then
            bAbort = False
            'Debug.Print("[Starting]")
            Me.State = States.Starting
            bStop = False
            t = New Thread(AddressOf _DoWork)
            t.Priority = _tp
            t.Start(New Object() {d, state})
            'Debug.Print("[Start]")
            'RaiseEvent Started(Me, Nothing)
            Me.State = States.Started
            Return True
        End If
    End Function
    Public Sub Abort()
        Debug.Print(Me.State.ToString)
        Select Case Me.State
            Case States.Aborting
                'this was called earlier
                Return
            Case States.CleaningUp
                'let that go through
                Return
            Case States.Started
                Me.State = States.Aborting
                If t IsNot Nothing AndAlso t.IsAlive Then
                    t.Abort()
                End If
            Case States.Starting
                'this should be tough to get to 
                Me.bStop = True
            Case States.Stopped
                Return
            Case States.Stopping
                'in a shutdown sequence
            Case States.Undefined
        End Select
        'wait.Set()

    End Sub
    'Public Sub Start(ByVal state As Object)
    '    RaiseEvent Starting(Me, Nothing)
    '    bAbort = False
    '    Debug.Print("[Starting]")
    '    bStop = False
    '    t = New Thread(AddressOf _DoWork)
    '    t.Start(state)
    '    Debug.Print("[Start]")
    '    RaiseEvent Started(Me, Nothing)
    'End Sub
    Public Property RequestedStop() As Boolean
        Get
            Return bStop
        End Get
        Set(ByVal value As Boolean)
            bStop = value
        End Set
    End Property
    ''' <summary>
    ''' Cleanly stops the the thread
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RequestStop()
        If State = States.Started Then
            Me.State = States.Stopping
            'bStop = True
            WaitProgress(New Object() {wait})
            Me.State = States.Stopped
        End If
    End Sub
    Public Sub [Stop]()
        If _state = States.Started Then
            bAbort = True
            'RaiseEvent AbortStarted(Nothing)
            'Debug.Print("[Stopping]")
            Me.State = States.Stopping
            'bStop = True
            'wait.WaitOne()

            Dim t3 As New Thread(AddressOf WaitProgress)
            t3.Priority = _tp
            t3.Start(New Object() {wait})

            'AppendText(TextBox1, ".")
            't.Abort()
        End If
    End Sub
    Public Sub [Stop](ByVal bolAsync As Boolean)
        If bolAsync = True Then
            [Stop]()
        Else
            If _state = States.Started Then
                Me.State = States.Stopping
                bAbort = True
                'RaiseEvent AbortStarted(Nothing)
                'Debug.Print("[Stopping]")
                'bStop = True
                'Dim t3 As New Thread(AddressOf WaitProgress)
                't3.Start(New Object() {wait})
                WaitProgress(New Object() {wait})

            End If
        End If
    End Sub
    Private Sub WaitOnThread(ByVal o As Object)
        If UBound(o) = 0 Then
            DirectCast(o(0), AutoResetEvent).WaitOne()
        Else
            AutoResetEvent.WaitAll(DirectCast(o, AutoResetEvent()))
        End If
        AutoResetEvent.SignalAndWait(p, p2)
    End Sub
    Dim _Timeout As Integer = 1000
    Public Property ProgressTimeout() As Integer
        Get
            Return _Timeout
        End Get
        Set(ByVal value As Integer)
            _Timeout = value
        End Set
    End Property
    Private Sub WaitProgress(ByVal o As Object)
        Dim start = Now
        t2 = New Thread(AddressOf WaitOnThread)
        t2.Priority = _tp
        t2.Start(o)
        bStop = True
        Do Until p.WaitOne(_Timeout)
            RaiseEvent AbortProgress(Now - start)
        Loop
        p2.Set()
    End Sub

    Private bStop As Boolean
    Private Sub _DoWork(ByVal o As Object)
        '_threadContext = SynchronizationContext.Current
        'Dim oz() As Object
        'If IsArray(o) Then
        'ReDim oz(UBound(o) - 1)
        'Array.ConstrainedCopy(o, 1, oz, 0, UBound(oz) + 1)
        '
        'Else
        '    oz = o
        'End If
        Try
            'Do Until bStop = True
            '    Thread.Sleep(10000)
            'Loop
            'RaiseEvent DoWork(Me, oz)
            DirectCast(o(0), ParameterizedThreadStart).Invoke(o(1))
        Catch abort As ThreadAbortException
            'RaiseEvent Finished(True)
            bAbort = True
        Catch ex As Exception
            RaiseEvent OnError(ex)
            'Throw ex
        Finally
            State = States.CleaningUp
            RaiseEvent CleanUp(o)
            wait.Set()
            'RaiseEvent Finished(bAbort)
            State = States.Stopped
        End Try
    End Sub
    'Private Sub _CleanUp()
    '    wait.Close()
    '    p.Close()
    '    p2.Close()
    'End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        bStop = True
        Abort()
    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
                [Stop](False)
                'Abort()
                wait.Close()
                p.Close()
                p2.Close()
                context = Nothing
                '_threadContext = Nothing
                Events.Dispose()
                stateLock = Nothing
                wait = Nothing
                p = Nothing
                p2 = Nothing
                t2 = Nothing
                t = Nothing

            End If
            Abort()
            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
Public Class Shortcut
    Implements IDisposable

#Region " COM Interfaces "

    <ComImport(), Guid("000214EE-0000-0000-C000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IShellLinkA

        Sub GetPath(
         <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszFile As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer,
         <[In]()> ByVal pfd As IntPtr,
         <[In]()> ByVal fFlags As SLGP_FLAGS)

        Function GetIDList() As IntPtr

        Sub SetIDList(<[In]()> ByVal pidl As IntPtr)

        Sub GetDescription(
         <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszName As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer)

        Sub SetDescription(
         <[In](), MarshalAs(UnmanagedType.LPStr)> ByVal pszName As String)

        Sub GetWorkingDirectory(
         <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszDir As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer)

        Sub SetWorkingDirectory(
         <[In](), MarshalAs(UnmanagedType.LPStr)>
         ByVal pszDir As String)

        Sub GetArguments(
         <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszArgs As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer)

        Sub SetArguments(
         <[In](), MarshalAs(UnmanagedType.LPStr)> ByVal pszArgs As String)

        Function GetHotkey() As Short

        Sub SetHotkey(
         <[In]()> ByVal wHotkey As Short)

        Function GetShowCmd() As Integer

        Sub SetShowCmd(
         <[In]()> ByVal iShowCmd As Integer)

        Sub GetIconLocation(
         <Out(), MarshalAs(UnmanagedType.LPStr)> ByVal pszIconPath As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer,
         <Out()> ByRef piIcon As Integer)

        Sub SetIconLocation(
         <[In](), MarshalAs(UnmanagedType.LPStr)> ByVal pszIconPath As String,
         <[In]()> ByVal iIcon As Integer)

        Sub SetRelativePath(
         <[In](), MarshalAs(UnmanagedType.LPStr)> ByVal pszPathRel As String,
         <[In]()> ByVal dwReserved As Integer)

        Sub Resolve(
         <[In]()> ByVal hwnd As IntPtr,
         <[In]()> ByVal fFlags As SLR_FLAGS)

        Sub SetPath(
         <[In](), MarshalAs(UnmanagedType.LPStr)> ByVal pszFile As String)

    End Interface

    <ComImport(), Guid("000214F9-0000-0000-C000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IShellLinkW

        Sub GetPath(
         <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer,
            <[In]()> ByVal pfd As IntPtr,
            <[In]()> ByVal fFlags As SLGP_FLAGS)

        Function GetIDList() As IntPtr

        Sub SetIDList(<[In]()> ByVal pidl As IntPtr)

        Sub GetDescription(
         <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer)

        Sub SetDescription(
         <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String)

        Sub GetWorkingDirectory(
         <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer)

        Sub SetWorkingDirectory(
         <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As String)

        Sub GetArguments(
         <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer)

        Sub SetArguments(
         <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As String)

        Function GetHotkey() As Short

        Sub SetHotkey(
         <[In]()> ByVal wHotkey As Short)

        Function GetShowCmd() As Integer

        Sub SetShowCmd(
         <[In]()> ByVal iShowCmd As Integer)

        Sub GetIconLocation(
         <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As System.Text.StringBuilder,
         <[In]()> ByVal cch As Integer,
         <Out()> ByRef piIcon As Integer)

        Sub SetIconLocation(
         <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As String,
         <[In]()> ByVal iIcon As Integer)

        Sub SetRelativePath(
         <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszPathRel As String,
         <[In]()> ByVal dwReserved As Integer)

        Sub Resolve(
         <[In]()> ByVal hwnd As IntPtr,
         <[In]()> ByVal fFlags As SLR_FLAGS)

        Sub SetPath(
         <[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As String)

    End Interface

    <Flags()> Private Enum SLGP_FLAGS
        SLGP_SHORTPATH = &H1
        SLGP_UNCPRIORITY = &H2
        SLGP_RAWPATH = &H4
    End Enum

    <Flags()> Private Enum SLR_FLAGS
        SLR_NO_UI = &H1                    ' don't post any UI durring the resolve
        ' operation, not msgs are pumped
        SLR_ANY_MATCH = &H2                ' no longer used
        SLR_UPDATE = &H4                   ' save the link back to it's file if the 
        ' track made it dirty
        SLR_NOUPDATE = &H8
        SLR_NOSEARCH = &H10                ' don't execute the search heuristics
        SLR_NOTRACK = &H20                 ' don't use NT5 object ID to track the link
        SLR_NOLINKINFO = &H40              ' don't use the net and volume relative info
        SLR_INVOKE_MSI = &H80              ' if we have a darwin link, then call msi to 
        ' fault in the applicaion
        SLR_NO_UI_WITH_MSG_PUMP = &H101    ' SLR_NO_UI + requires an enable modeless site or HWND
    End Enum

#End Region

    Public Enum WindowMode
        Normal = 1
        Minimized = 7
        Maximized = 3
    End Enum

    Private _filename As String
    Private _link As Object
    Private _linkA As IShellLinkA
    Private _linkW As IShellLinkW ' Implemented only on NT based OSes

    Public Sub New(ByVal filename As String)

        Me.New()

        ' If the shortcut doesn't exists create one
        If Not System.IO.File.Exists(filename) Then

            Me.Save(filename)

        Else

            ' Load the shortcut
            Load(filename)

        End If

    End Sub

    Public Sub New()

        ' Get the type for the shell shortcut object
        Dim t As Type = Type.GetTypeFromCLSID(
                           New Guid("00021401-0000-0000-C000-000000000046"))

        ' Create an instance of the shortcut object
        _link = Activator.CreateInstance(t)

        ' Try to get the Unicode interface and
        ' if it fails get the Ansi one.
        If TypeOf _link Is IShellLinkW Then
            _linkW = DirectCast(_link, IShellLinkW)
        Else
            _linkA = DirectCast(_link, IShellLinkA)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the application executable path.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ApplicationPath() As String
        Get

            Dim data As New System.Text.StringBuilder(260)

            If Not _linkW Is Nothing Then
                _linkW.GetPath(data, data.Capacity, IntPtr.Zero, SLGP_FLAGS.SLGP_RAWPATH)
            Else
                _linkA.GetPath(data, data.Capacity, IntPtr.Zero, SLGP_FLAGS.SLGP_RAWPATH)
            End If

            Return data.ToString

        End Get
        Set(ByVal value As String)

            If Not _linkW Is Nothing Then
                _linkW.SetPath(value)
            Else
                _linkA.SetPath(value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get or sets the command line arguments.
    ''' </summary>
    ''' <returns>The command line arguments.</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Arguments() As String
        Get

            Dim data As New System.Text.StringBuilder(260)

            If Not _linkW Is Nothing Then
                _linkW.GetArguments(data, data.Capacity)
            Else
                _linkA.GetArguments(data, data.Capacity)
            End If

            Return data.ToString

        End Get
        Set(ByVal value As String)

            If Not _linkW Is Nothing Then
                _linkW.SetArguments(value)
            Else
                _linkA.SetArguments(value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the shortcut description.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>The description is shown as a tooltip by explorer.</remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Description() As String
        Get

            Dim data As New System.Text.StringBuilder(260)

            If Not _linkW Is Nothing Then
                _linkW.GetDescription(data, data.Capacity)
            Else
                _linkA.GetDescription(data, data.Capacity)
            End If

            Return data.ToString

        End Get
        Set(ByVal value As String)

            If Not _linkW Is Nothing Then
                _linkW.SetDescription(value)
            Else
                _linkA.SetDescription(value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the shortcut hot key.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	03/11/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property HotKey() As Short
        Get

            If Not _linkW Is Nothing Then
                Return _linkW.GetHotkey
            Else
                Return _linkA.GetHotkey
            End If

        End Get
        Set(ByVal value As Short)

            If Not _linkW Is Nothing Then
                _linkW.SetHotkey(value)
            Else
                _linkA.SetHotkey(value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the file that contains the shortcut icon.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property IconFile() As String
        Get

            Dim data As New System.Text.StringBuilder(260)

            If Not _linkW Is Nothing Then
                _linkW.GetIconLocation(data, data.Capacity, 0)
            Else
                _linkA.GetIconLocation(data, data.Capacity, 0)
            End If

            Return data.ToString

        End Get
        Set(ByVal value As String)

            If Not _linkW Is Nothing Then
                _linkW.SetIconLocation(value, Me.IconIndex)
            Else
                _linkA.SetIconLocation(value, Me.IconIndex)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the index of the shortcut icon in the IconFile file.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property IconIndex() As Integer
        Get

            Dim index As Integer

            If Not _linkW Is Nothing Then
                _linkW.GetIconLocation(New System.Text.StringBuilder(0), 0, index)
            Else
                _linkA.GetIconLocation(New System.Text.StringBuilder(0), 0, index)
            End If

            Return index

        End Get
        Set(ByVal value As Integer)

            If Not _linkW Is Nothing Then
                _linkW.SetIconLocation(Me.IconFile, value)
            Else
                _linkA.SetIconLocation(Me.IconFile, value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the initial state of the main window.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	03/11/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property ShowCommand() As WindowMode
        Get

            If Not _linkW Is Nothing Then
                Return CType(_linkW.GetShowCmd, WindowMode)
            Else
                Return CType(_linkA.GetShowCmd, WindowMode)
            End If

        End Get
        Set(ByVal value As WindowMode)

            If Not _linkW Is Nothing Then
                _linkW.SetShowCmd(value)
            Else
                _linkA.SetShowCmd(value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the application working directory.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property WorkingDirectory() As String
        Get

            Dim data As New System.Text.StringBuilder(260)

            If Not _linkW Is Nothing Then
                _linkW.GetWorkingDirectory(data, data.Capacity)
            Else
                _linkA.GetWorkingDirectory(data, data.Capacity)
            End If

            Return data.ToString

        End Get
        Set(ByVal value As String)

            If Not _linkW Is Nothing Then
                _linkW.SetWorkingDirectory(value)
            Else
                _linkA.SetWorkingDirectory(value)
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Releases the shortcut object.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Dispose() Implements System.IDisposable.Dispose

        Do While Marshal.ReleaseComObject(_link) > 0 : Loop

        _linkA = Nothing
        _linkW = Nothing
        _link = Nothing

        GC.SuppressFinalize(Me)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Loads the shortcut file.
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub Load(ByVal filename As String)

        Dim pf As System.Runtime.InteropServices.ComTypes.IPersistFile

        pf = DirectCast(_link, System.Runtime.InteropServices.ComTypes.IPersistFile)
        pf.Load(filename, 0)

        _filename = filename

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Saves the shortcut to a file.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------        
    Public Sub Save(ByVal filename As String)

        Dim pf As System.Runtime.InteropServices.ComTypes.IPersistFile

        pf = DirectCast(_link, System.Runtime.InteropServices.ComTypes.IPersistFile)
        pf.Save(filename, False)

        _filename = filename

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Saves the shortcut to the current file.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Eduardo Morcillo]	11/03/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub Save()

        If _filename Is Nothing Then Throw New ArgumentNullException

        Save(_filename)

    End Sub


    Public Shared Sub CreateShortcut(ByVal shortcutPath As String, ByVal targetPath As String, ByVal workingDirectory As String, Optional ByVal description As String = "")

        Dim cShellLink As New Win32.Shell32.ShellLink()
        Dim iShellLink As Win32.Shell32.IShellLink = DirectCast(cShellLink, Win32.Shell32.IShellLink)
        iShellLink.SetDescription(description)
        iShellLink.SetShowCmd(1)
        iShellLink.SetPath(targetPath)
        iShellLink.SetWorkingDirectory(workingDirectory)
        Dim iPersistFile As ComTypes.IPersistFile = DirectCast(iShellLink, ComTypes.IPersistFile)
        iPersistFile.Save(shortcutPath, False)
        Marshal.ReleaseComObject(iPersistFile)
        iPersistFile = Nothing
        Marshal.ReleaseComObject(iShellLink)
        iShellLink = Nothing
        Marshal.ReleaseComObject(cShellLink)
        cShellLink = Nothing
    End Sub
End Class
Public Class HiResTimer
    Private Shared isPerfCounterSupported As Boolean = False
    Private Shared timerFrequency As Int64 = 0

    ' Windows CE native library with QueryPerformanceCounter().
    'Private Const [lib] As String = "coredll.dll"

    Public Declare Function QueryPerformanceCounter Lib "KERNEL32.dll" _
    (ByRef count As Int64) As Integer

    Public Declare Function QueryPerformanceFrequency Lib "KERNEL32.dll" _
    (ByRef timerFrequency As Int64) As Integer

    Shared Sub New()
        ' Query the high-resolution timer only if it is supported.
        ' A returned frequency of 1000 typically indicates that it is not
        ' supported and is emulated by the OS using the same value that is
        ' returned by Environment.TickCount.
        ' A return value of 0 indicates that the performance counter is
        ' not supported.
        Dim returnVal As Integer = QueryPerformanceFrequency(timerFrequency)

        If returnVal <> 0 AndAlso timerFrequency <> 1000 Then
            ' The performance counter is supported.
            isPerfCounterSupported = True
        Else
            ' The performance counter is not supported. Use
            ' Environment.TickCount instead.
            timerFrequency = 1000
        End If

    End Sub


    Public Shared ReadOnly Property Frequency() As Int64
        Get
            Return timerFrequency
        End Get
    End Property


    Public Shared ReadOnly Property Value() As Int64
        Get
            Dim tickCount As Int64 = 0

            If isPerfCounterSupported Then
                ' Get the value here if the counter is supported.
                QueryPerformanceCounter(tickCount)
                Return tickCount
            Else
                ' Otherwise, use Environment.TickCount
                Return CType(Environment.TickCount, Int64)
            End If
        End Get
    End Property
    Public Const To10thMillisecond As Int64 = 10000
    Public Const ToMillisecond As Int64 = 1000000
    Public Const ToSecond As Int64 = 1000000000
    'Public Const ToNanoSecond As Int64 = 1
    Public Const ToTick As Int64 = 1

    Public Shared ReadOnly Property ValueIn10thMilliseconds() As Int64
        Get
            Return Value * 10000 / Frequency
        End Get
    End Property
    Public Shared ReadOnly Property TimeElapsedInTicks(ByVal startTime As Int64) As Int64
        Get
            Return Value - startTime
        End Get
    End Property
    Public Shared ReadOnly Property TimeElapsedInTicks(ByVal startTime As Int64, ByVal endTime As Int64) As Int64
        Get
            Return endTime - startTime
        End Get
    End Property
    Public Shared ReadOnly Property TimeElapsedIn10thMilliseconds(ByVal startTime As Int64)
        Get
            Return TimeElapsedInTicks(startTime) * 10000 / Frequency
        End Get
    End Property
    Public Shared ReadOnly Property TimeElapsedIn10thMilliseconds(ByVal startTime As Int64, ByVal endTime As Int64) As Int64
        Get
            Return TimeElapsedInTicks(startTime, endTime) * 10000 / Frequency
        End Get
    End Property


    Dim _start As Int64
    Public Sub Start()
        QueryPerformanceCounter(_start)
    End Sub
    Dim _stop As Int64
    Public Sub [Stop]()
        QueryPerformanceCounter(_stop)
    End Sub
    Public Sub Reset()
        _start = 0
        _stop = 0
    End Sub
    Public ReadOnly Property GetStartTime() As Int64
        Get
            Return _start
        End Get
    End Property
    Public ReadOnly Property GetStopTime() As Int64
        Get
            Return _stop
        End Get
    End Property
    Public ReadOnly Property HasStarted() As Boolean
        Get
            Return _start <> 0
        End Get
    End Property

    ''' <summary>
    ''' accepts the number of iterations and returns a duration value. 
    ''' Use this method to calculate the number of ticks between the start and stop values. 
    ''' Next, multiply the result by the frequency multiplier to calculate the duration of all the operations, and then divide by the number of iterations to arrive at the duration per operation value. 
    ''' </summary>
    ''' <param name="iterations"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Duration(ByVal iterations As Integer) As Int64
        Get
            Return ((((_stop - _start) *
        multiplier) /
         Frequency) / iterations)
        End Get
    End Property
    Public ReadOnly Property ElapsedTime() As Int64
        Get
            Return _stop - _start
        End Get
    End Property
    Public ReadOnly Property ElapsedTime(ByVal StartTime As Int64) As Int64
        Get
            Return Value - _start
        End Get
    End Property
    Public ReadOnly Property ElapsedTime(ByVal StartTime As Int64, ByVal EndTime As Int64) As Int64
        Get
            Return EndTime - StartTime
        End Get
    End Property
    Private Shared multiplier As New Decimal(1000000000.0)



    'Shared Sub Main()
    '    Dim timer As New HiResTimer()

    '    ' This example shows how to use the high-resolution counter to 
    '    ' time an operation. 
    '    ' Get counter value before the operation starts.
    '    Dim counterAtStart As Int64 = timer.Value

    '    ' Perform an operation that takes a measureable amount of time.
    '    Dim count As Integer
    '    For count = 0 To 9999
    '        count += 1
    '        count -= 1
    '    Next count

    '    ' Get counter value after the operation ends.
    '    Dim counterAtEnd As Int64 = timer.Value

    '    ' Get time elapsed in tenths of milliseconds
    '    Dim timeElapsedInTicks As Int64 = counterAtEnd - counterAtStart
    '    Dim timeElapseInTenthsOfMilliseconds As Int64 = timeElapsedInTicks * 10000 / timer.Frequency


    '    MessageBox.Show("Time Spent in operation (tenths of ms) " + timeElapseInTenthsOfMilliseconds.ToString + vbLf + "Counter Value At Start: " + counterAtStart.ToString + vbLf + "Counter Value At End : " + counterAtEnd.ToString + vbLf + "Counter Frequency : " + timer.Frequency.ToString)

    'End Sub
End Class
Public Class KeyTimer
    Implements IDisposable

    Dim t As Threading.Timer
    Private iTTL As Integer = 420
    Dim bStart As Boolean = False
    Public Property TTL() As Integer
        Get
            Return iTTL
        End Get
        Set(ByVal value As Integer)
            iTTL = value
        End Set
    End Property
    Public Sub Reset()
        't.Dispose()
        't = New Threading.Timer(remember, Nothing, iTTL, Threading.Timeout.Infinite)
        If t IsNot Nothing Then t.Change(iTTL, Threading.Timeout.Infinite)
    End Sub
    'Shared remember As Threading.TimerCallback
    Public Sub Start(ByVal callback As Threading.TimerCallback)
        'remember = callback
        t = New Threading.Timer(callback, Nothing, iTTL, Threading.Timeout.Infinite)

        bStart = True
    End Sub
    Public Sub [Stop]()
        If t IsNot Nothing Then t.Dispose()
        t = Nothing
        bStart = False
    End Sub
    Public ReadOnly Property IsStarted() As Boolean
        Get
            Return bStart
        End Get
    End Property

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If t IsNot Nothing Then t.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
'Public NotInheritable Class SoundBucket
'    'Implements IDisposable


'    Private Shared ApplicationBuffer As New List(Of SecondaryBuffer)
'    Private Shared ApplicationDevice As Device = Nothing

'    Private Shared hash As New Hashtable
'    Private Delegate Sub Wait(ByVal EventHandle As Threading.EventWaitHandle)
'    'Public Sub New(ByVal dev As Device)
'    '    ApplicationDevice = dev
'    'End Sub
'    'Shared Sub New()
'    '    ApplicationDevice = New Device()
'    ''ApplicationDevice.SetCooperativeLevel(Handle, CooperativeLevel.Normal)
'    'end Sub
'    'Private Shared filepath As New List(Of Object)
'    Private Shared bReady As Boolean = False
'    Public Shared ReadOnly Property Ready() As Boolean
'        Get
'            Return bReady
'        End Get
'    End Property
'    Public Shared Sub Init(ByVal handle As IntPtr)
'        ApplicationDevice = New Device
'        ApplicationDevice.SetCooperativeLevel(handle, CooperativeLevel.Normal)
'        bReady = True
'    End Sub

'    Private Shared Sub Add(ByVal strKey As String, ByVal strPath As String)

'        hash.Add(strKey, hash.Count)
'        'filepath.Add(strPath)
'        Dim desc As New BufferDescription()
'        desc.ControlPositionNotify = True
'        ApplicationBuffer.Add(New SecondaryBuffer(strPath, desc, ApplicationDevice))
'        desc.Dispose()
'    End Sub
'    Private Shared Sub Add(ByVal strKey As String, ByVal strPath As String, ByVal desc As BufferDescription)
'        hash.Add(strKey, hash.Count)
'        'filepath.Add(strPath)
'        ApplicationBuffer.Add(New Microsoft.DirectX.DirectSound.Buffer(strPath, desc, ApplicationDevice))
'    End Sub
'    Private Shared Sub Add(ByVal strKey As String, ByVal stream As IO.Stream)
'        hash.Add(strKey, hash.Count)
'        Dim desc As New BufferDescription
'        desc.ControlPositionNotify = True

'        ApplicationBuffer.Add(New SecondaryBuffer(stream, desc, ApplicationDevice))
'        desc.Dispose()
'    End Sub
'    Private Shared Sub Add(ByVal strKey As String, ByVal stream As IO.Stream, ByVal desc As BufferDescription)
'        hash.Add(strKey, hash.Count)
'        ApplicationBuffer.Add(New SecondaryBuffer(stream, desc, ApplicationDevice))
'    End Sub
'    Public Shared Sub Clear()
'        OnDeviceDestroy()
'        ApplicationBuffer.Clear()
'        hash.Clear()
'    End Sub
'    ''' <summary>
'    ''' Plays a sound by Index and invokes a callback method when the sound is stopped, by the application or end of file.
'    ''' </summary>
'    ''' <param name="index">Index of the sound to be played.</param>
'    ''' <param name="C">Callback routine, fired when the sound stops.</param>
'    ''' <param name="Priority"></param>
'    ''' <param name="Flags"></param>
'    ''' <remarks>The buffer must be created with the ControlPositionNotify flag set.</remarks>
'    ''' <exception cref="ControlUnavailableException">
'    ''' Thrown if the buffer was not created with ControlPositionNotify or the sound device does not support ControlPositionNotify.
'    ''' </exception>
'    Private Shared Sub Play(ByVal index As Integer, ByVal C As AsyncCallback, Optional ByVal Priority As Integer = 0, Optional ByVal Flags As BufferPlayFlags = BufferPlayFlags.Default)
'        If Not SoundBucket.Ready Then Return

'        If ApplicationBuffer(index).Caps.ControlPositionNotify = False Then
'            Throw New Microsoft.DirectX.DirectSound.ControlUnavailableException("ControlPositionNotify is not available on this buffer")
'            Return
'        End If
'        Dim w1 As New Wait(AddressOf Waiter2)

'        Dim nn() As BufferPositionNotify = {New BufferPositionNotify()}

'        Using n As New Notify(ApplicationBuffer(index))
'            Dim w As New Threading.EventWaitHandle(False, Threading.EventResetMode.ManualReset)

'            nn(0).Offset = PositionNotifyFlag.OffsetStop
'            nn(0).EventNotifyHandle = w.SafeWaitHandle.DangerousGetHandle
'            n.SetNotificationPositions(nn)


'            If ApplicationBuffer(index).Status.Playing = True Then
'                ApplicationBuffer(index).Stop()
'            End If
'            ApplicationBuffer(index).Play(Priority, Flags)

'            'Dim t As New Threading.Thread(AddressOf Waiter)
'            't.Start(New Object() {Callback, index, w})


'            w1.BeginInvoke(w, C, index)
'        End Using
'    End Sub
'    Private Shared Sub Play(ByVal Key As String, ByVal C As AsyncCallback, Optional ByVal Priority As Integer = 0, Optional ByVal Flags As BufferPlayFlags = BufferPlayFlags.Default)

'        Try
'            Play(hash(Key), C, Priority, Flags)
'        Catch
'        End Try
'    End Sub
'    Private Shared Sub Play(ByVal index As Integer, Optional ByVal Priority As Integer = 0, Optional ByVal Flags As BufferPlayFlags = BufferPlayFlags.Default)
'        If Not SoundBucket.Ready Then Return

'        If ApplicationBuffer(index).Status.Playing = True Then
'            ApplicationBuffer(index).Stop()
'        End If
'        ApplicationBuffer(index).Play(Priority, Flags)

'    End Sub

'    Private Shared Sub Play(ByVal strKey As String, Optional ByVal Priority As Integer = 0, Optional ByVal Flags As BufferPlayFlags = BufferPlayFlags.Default)
'        If Not SoundBucket.Ready Then Return

'        Dim index As Integer = hash(strKey)
'        If ApplicationBuffer(index).Status.Playing = True Then
'            ApplicationBuffer(index).Stop()
'        End If
'        ApplicationBuffer(index).Play(Priority, Flags)
'    End Sub
'    Private Shared Sub Shit(ByVal ar As IAsyncResult)
'        ar.AsyncState.dispose()
'        Dim s As Runtime.Remoting.Messaging.AsyncResult = DirectCast(ar, Runtime.Remoting.Messaging.AsyncResult)
'        DirectCast(s.AsyncDelegate, Wait).EndInvoke(ar)
'    End Sub



'    Private Shared Sub Play2(ByVal strKey As String)
'        If Not SoundBucket.Ready Then Return

'        If hash.ContainsKey(strKey) = False Then Return
'        'Dim w1 As New Wait(AddressOf Waiter2)
'        'Dim w As New Threading.EventWaitHandle(False, Threading.EventResetMode.ManualReset)
'        'Dim b As New SecondaryBuffer(filepath(hash(strKey)), ApplicationDevice)
'        Dim b As SecondaryBuffer = ApplicationBuffer.Item(hash(strKey)).Clone(ApplicationDevice)
'        'w1.BeginInvoke(w, AddressOf Shit, b)
'        b.Play(0, BufferPlayFlags.Default)
'    End Sub
'    Private Shared Sub Waiter2(ByVal e As Threading.EventWaitHandle)

'        Try
'            e.WaitOne()
'            e.Close()
'        Catch
'        End Try
'    End Sub
'    'Private Shared Sub Waiter(ByVal obj As Object)
'    '    obj(2).WaitOne()
'    '    obj(2).close()
'    '    DirectCast(obj(0), Callback).Invoke(obj(1))
'    'End Sub


'    Public Shared Sub [Stop](ByVal index As Integer)
'        If Not SoundBucket.Ready Then Return

'        ApplicationBuffer(index).Stop()
'    End Sub
'    Public Shared Sub [Stop](ByVal strKey As String)

'        [Stop](hash(strKey))
'    End Sub

'    'Public Function IsPlaying(ByVal Index As Integer) As Boolean
'    '    Return ApplicationBuffer(Index).Status.Playing
'    'End Function
'    'Public Function IsPlaying(ByVal strKey As String) As Boolean
'    '    Return IsPlaying(hash(strKey))
'    'End Function
'    Public Shared ReadOnly Property Status(ByVal Index As Integer) As BufferStatus
'        Get
'            Return ApplicationBuffer(Index).Status
'        End Get
'    End Property
'    Public Shared ReadOnly Property Status(ByVal strKey As String) As BufferStatus
'        Get
'            Return Status(hash(strKey))
'        End Get
'    End Property

'    Public Shared Sub OnDeviceDestroy()
'        For Each s As SecondaryBuffer In ApplicationBuffer
'            s.Dispose()
'        Next
'    End Sub
'    Public Shared Sub OnDeviceLost()
'    End Sub
'    Public Shared Sub OnDeviceRestore()
'        RestoreAll()
'    End Sub
'    Public Shared Sub RestoreAll()
'        For Each s As SecondaryBuffer In ApplicationBuffer
'            If Not s.Disposed Then s.Restore()
'        Next
'    End Sub
'    Public Shared Sub Restore(ByVal Index As Integer)
'        ApplicationBuffer(Index).Restore()
'    End Sub
'    Public Shared Sub Restore(ByVal strKey As String)
'        Restore(hash(strKey))
'    End Sub
'    Public Shared Property Item(ByVal Index As Integer) As SecondaryBuffer
'        Get
'            Return ApplicationBuffer(Index)
'        End Get
'        Set(ByVal value As SecondaryBuffer)
'            ApplicationBuffer(Index) = value
'        End Set
'    End Property
'    Public Shared Property Item(ByVal strKey As String) As SecondaryBuffer
'        Get
'            Return ApplicationBuffer(hash(strKey))

'        End Get
'        Set(ByVal value As SecondaryBuffer)
'            ApplicationBuffer(hash(strKey)) = value
'        End Set
'    End Property

'    '#Region " IDisposable Support "
'    '        Private disposedValue As Boolean = False        ' To detect redundant calls

'    '        ' IDisposable
'    '        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
'    '            If Not Me.disposedValue Then
'    '                If disposing Then
'    '                    ' TODO: free managed resources when explicitly called
'    '                    'OnDeviceDestroy()
'    '                    Clear()
'    '                    'ApplicationDevice.Dispose()
'    '                End If

'    '                ' TODO: free shared unmanaged resources
'    '            End If
'    '            Me.disposedValue = True
'    '        End Sub


'    '        ' This code added by Visual Basic to correctly implement the disposable pattern.
'    '        Public Sub Dispose() Implements IDisposable.Dispose
'    '            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'    '            Dispose(True)
'    '            GC.SuppressFinalize(Me)
'    '        End Sub
'    '#End Region
'    Public Shared Sub BufferSound(ByVal strKey As String, ByVal stream As IO.Stream)
'        Add(strKey, stream)
'    End Sub
'    Public Shared Sub PlayBufferedSound(ByVal strKey As String, ByVal ar As AsyncCallback)
'        Play(strKey, ar)

'    End Sub
'    Public Shared Sub PlayBufferedSoundSync(ByVal strKey As String, Optional ByVal Priority As Integer = 0, Optional ByVal Flags As BufferPlayFlags = BufferPlayFlags.Default)
'        If Not SoundBucket.Ready Then Return

'        Dim index As Integer = hash(strKey)
'        If ApplicationBuffer(index).Caps.ControlPositionNotify = False Then
'            Throw New Microsoft.DirectX.DirectSound.ControlUnavailableException("ControlPositionNotify is not available on this buffer")
'            Return
'        End If
'        Dim w1 As New Wait(AddressOf Waiter2)

'        Dim nn() As BufferPositionNotify = {New BufferPositionNotify()}

'        Using n As New Notify(ApplicationBuffer(index))
'            Dim w As New Threading.EventWaitHandle(False, Threading.EventResetMode.ManualReset)

'            nn(0).Offset = PositionNotifyFlag.OffsetStop
'            nn(0).EventNotifyHandle = w.SafeWaitHandle.DangerousGetHandle
'            n.SetNotificationPositions(nn)


'            If ApplicationBuffer(index).Status.Playing = True Then
'                ApplicationBuffer(index).Stop()
'            End If
'            ApplicationBuffer(index).Play(Priority, Flags)

'            'Dim t As New Threading.Thread(AddressOf Waiter)
'            't.Start(New Object() {Callback, index, w})

'            w1.Invoke(w)
'            'w1.BeginInvoke(w, C, index)
'        End Using
'    End Sub
'    Public Shared Sub PlayBufferedSound(ByVal strKey As String)
'        Play2(strKey)
'    End Sub

'    Private Shared disposedValue As Boolean = False        ' To detect redundant calls
'    Public Shared Sub Dispose()
'        If ApplicationDevice IsNot Nothing Then ApplicationDevice.Dispose()
'        bReady = False
'    End Sub
'    ' IDisposable
'    '        Protected Shared Sub Dispose(ByVal disposing As Boolean)
'    '            If Not disposedValue Then
'    '                If disposing Then
'    '                    ' TODO: free other state (managed objects).
'    '                    Cleanup()
'    '                End If

'    '                ' TODO: free your own state (unmanaged objects).
'    '                ' TODO: set large fields to null.
'    '            End If
'    '            disposedValue = True
'    '        End Sub

'    '#Region " IDisposable Support "
'    '        ' This code added by Visual Basic to correctly implement the disposable pattern.
'    '        Public Sub Dispose() Implements IDisposable.Dispose
'    '            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'    '            Dispose(True)
'    '            GC.SuppressFinalize(Me)
'    '        End Sub
'    '#End Region

'End Class
Namespace Consoles
    Public Interface IHeaderPlugin
        Sub Initialize(ByVal host As IHost)
        ReadOnly Property Name() As String
        ReadOnly Property Author() As String
        ReadOnly Property Version() As String
        ReadOnly Property TargetSystems() As String()
        'ReadOnly Property Gather(ByVal strDir As String, ByVal Recurse As Boolean) As List(Of IBaseRom)
        'ReadOnly Property Gather(ByVal strDir As String, ByVal Recurse As Boolean) As List(Of String)
        ReadOnly Property MakeRom(ByVal strDir As String) As IBaseRom
        Function GetSorter() As Comparison(Of IBaseRom)
        Property Extensions() As String()
        'ReadOnly Property FieldListing() As String()
    End Interface
    Public Interface IHost
        Sub ShowFeedBack(ByVal strFeedBack As String)
    End Interface

    Public Interface IBaseRom
        ReadOnly Property Path() As String
        ReadOnly Property Title() As String
        ReadOnly Property Fields() As Dictionary(Of String, String)
        ReadOnly Property InternalImage() As System.Drawing.Image
    End Interface
    Public Class Plugins


        '  Public Structure AvailablePlugin
        '      Public AssemblyPath As String
        '      Public ClassName As String
        '  End Structure


        '  Public Shared Function FindPlugins(ByVal strPath As String, ByVal strInterface As _
        'String) As AvailablePlugin()
        '      Dim Plugins As ArrayList = New ArrayList()
        '      Dim strDLLs() As String, intIndex As Integer
        '      Dim objDLL As [Assembly]
        '      'Go through all DLLs in the directory, attempting to load them
        '      strDLLs = IO.Directory.GetFileSystemEntries(strPath, "*.dll")
        '      For intIndex = 0 To strDLLs.Length - 1
        '          Try
        '              objDLL = [Assembly].LoadFrom(strDLLs(intIndex))
        '              ExamineAssembly(objDLL, strInterface, Plugins)
        '          Catch e As Exception
        '              'Error loading DLL, we don't need to do anything special
        '          End Try
        '      Next
        '      'Return all plugins found
        '      Dim Results(Plugins.Count - 1) As AvailablePlugin
        '      If Plugins.Count <> 0 Then
        '          Plugins.CopyTo(Results)
        '          For Each a In Results
        '              Dim i As IHeaderPlugin = DirectCast(CreateInstance(a), IHeaderPlugin)
        '              For Each s In i.TargetSystems
        '                  fAvail.Add(s, a)
        '              Next
        '              i = Nothing
        '          Next
        '          Return Results
        '      Else
        '          Return Nothing
        '      End If
        '  End Function
        '  Private Shared fAvail As New Dictionary(Of String, AvailablePlugin)
        '  Public Shared ReadOnly Property AvailablePlugins(ByVal strKey As String) As AvailablePlugin
        '      Get
        '          Return fAvail(strKey)
        '      End Get
        '  End Property
        '  Private Shared Sub ExamineAssembly(ByVal objDLL As [Assembly], _
        ' ByVal strInterface As String, ByVal Plugins As ArrayList)
        '      Dim objType As Type
        '      Dim objInterface As Type
        '      Dim Plugin As AvailablePlugin
        '      'Loop through each type in the DLL
        '      For Each objType In objDLL.GetTypes
        '          'Only look at public types
        '          If objType.IsPublic = True Then
        '              'Ignore abstract classes
        '              If Not ((objType.Attributes And TypeAttributes.Abstract) = _
        '              TypeAttributes.Abstract) Then
        '                  'See if this type implements our interface
        '                  objInterface = objType.GetInterface(strInterface, True)
        '                  If Not (objInterface Is Nothing) Then
        '                      'It does
        '                      Plugin = New AvailablePlugin()
        '                      Plugin.AssemblyPath = objDLL.Location
        '                      Plugin.ClassName = objType.FullName
        '                      Plugins.Add(Plugin)
        '                  End If
        '              End If
        '          End If
        '      Next
        '  End Sub
        '  Public Shared Function CreateInstance(ByVal Plugin As AvailablePlugin) As Object
        '      Dim objDLL As [Assembly]
        '      Dim objPlugin As Object
        '      Try
        '          'Load dll
        '          objDLL = [Assembly].LoadFrom(Plugin.AssemblyPath)
        '          'Create and return class instance
        '          objPlugin = objDLL.CreateInstance(Plugin.ClassName)
        '      Catch e As Exception
        '          Return Nothing
        '      End Try
        '      Return objPlugin
        '  End Function

        'Public Shared Function GetPlugins(Of T)(ByVal folder As String) As List(Of T)
        '    Dim files As String() = IO.Directory.GetFiles(folder, "*.dll")
        '    Dim tList As New List(Of T)
        '    Debug.Assert(GetType(T).IsInterface)
        '    For Each file As String In files
        '        Try

        '            Dim ass As [Assembly] = [Assembly].LoadFile(file)
        '            For Each type As Type In ass.GetTypes()
        '                If (Not type.IsClass OrElse type.IsNotPublic) Then Continue For
        '                Dim interfaces As Type() = type.GetInterfaces()
        '                If interfaces.Contains(GetType(T)) Then
        '                    Dim obj As Object = Activator.CreateInstance(type)
        '                    Dim tz As T = CType(obj, T)
        '                    tList.Add(tz)
        '                End If
        '            Next
        '        Catch ex As Exception
        '            'LogError(ex)
        '        End Try
        '    Next
        '    Return tList
        'End Function
        Public Shared Function FindPlugins(Of t)(ByVal folder As String) As List(Of t)
            If Not IO.Directory.Exists(folder) Then Return Nothing
            Dim strDLLs = IO.Directory.GetFileSystemEntries(folder, "*.dll")
            Dim l As New List(Of t)
            For intIndex = 0 To strDLLs.Length - 1
                Dim asm As [Assembly] = [Assembly].LoadFrom(strDLLs(intIndex))
                'Dim myType As System.Type = asm.GetType() ' asm.GetType) ' + ".Plugin")
                For Each objType In asm.GetTypes
                    Debug.Print(objType.Name)
                    Dim implementsIPlugin As Boolean = GetType(t).IsAssignableFrom(objType)
                    If implementsIPlugin Then
                        l.Add(asm.CreateInstance(objType.FullName))
                    End If

                Next
            Next
            Return l
        End Function

        Public Shared DefaultComparer As New Comparison(Of IBaseRom)(Function(x As IBaseRom, y As IBaseRom) (StrComp(x.Title, y.Title, CompareMethod.Text)))



        Public Shared Function GetFilesFromDir(ByVal strDir As String, ByVal strExtensionList() As String, ByVal bRecurse As Boolean) As List(Of String)
            If Not IO.Directory.Exists(strDir) Then Return Nothing
            Dim l As New List(Of String)
            For Each s In IO.Directory.GetFiles(strDir)
                Dim i As Integer = Array.IndexOf(strExtensionList, IO.Path.GetExtension(s).ToLower)
                If i >= 0 Then
                    l.Add(s)
                End If
            Next
            If bRecurse Then
                For Each d In IO.Directory.GetDirectories(strDir)
                    l.AddRange(GetFilesFromDir(d, strExtensionList, bRecurse))
                Next
            End If
            Return l
        End Function
        Public Shared Function GetRomsFromDir(ByVal Plugin As Consoles.IHeaderPlugin, ByVal strDir As String, ByVal strExtensionList() As String, ByVal bRecurse As Boolean) As List(Of Consoles.IBaseRom)
            If Not IO.Directory.Exists(strDir) Then Return Nothing
            Dim l As New List(Of Consoles.IBaseRom)
            For Each s In IO.Directory.GetFiles(strDir)
                Dim i As Integer = Array.IndexOf(strExtensionList, IO.Path.GetExtension(s).ToLower)
                If i >= 0 Then
                    Dim z = Consoles.Plugins.TryMakeRom(Plugin, s)
                    If z IsNot Nothing Then l.Add(z)
                End If
            Next
            If bRecurse Then
                For Each d In IO.Directory.GetDirectories(strDir)
                    l.AddRange(GetRomsFromDir(Plugin, d, strExtensionList, bRecurse))
                Next
            End If
            Return l
        End Function
        Public Shared Function TryMakeRom(ByVal plugin As Consoles.IHeaderPlugin, ByVal strFile As String) As Consoles.IBaseRom
            Dim z As Consoles.IBaseRom
            Try
                z = plugin.MakeRom(strFile)
            Catch ex As Exception
                Return Nothing
            End Try
            Return z
        End Function

        'Public Shared Function Unzip(ByVal strZipFile As String) As String
        '    Dim myShell As New Shell32.Shell
        '    Dim sourceFolder As Shell32.Folder = myShell.NameSpace(strZipFile)
        '    Dim destinationFolder As Shell32.Folder = myShell.NameSpace(sourceFolder.ParentFolder)
        '    Dim outputFilename As String = sourceFolder.Items().Item(0).Name
        '    If IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strZipFile), outputFilename)) Then Return Nothing
        '    destinationFolder.CopyHere(sourceFolder.Items(), 4 Or 8 Or 16 Or 512 Or 1024)
        '    Return IO.Path.Combine(IO.Path.GetDirectoryName(strZipFile), outputFilename)
        'End Function
        'Public Shared Function Unzip(ByVal strZipFile As String, ByVal destDir As String) As String
        '    Dim myShell As New Shell32.Shell
        '    Dim sourceFolder As Shell32.Folder = myShell.NameSpace(strZipFile)
        '    Dim destinationFolder As Shell32.Folder = myShell.NameSpace(destDir) 'sourceFolder.ParentFolder)
        '    Dim outputFilename As String = sourceFolder.Items().Item(0).Name
        '    destinationFolder.CopyHere(sourceFolder.Items(), 4 Or 8 Or 16 Or 512 Or 1024)
        '    Return IO.Path.Combine(destDir, outputFilename)
        'End Function
    End Class
End Namespace



Namespace Win32
    Public NotInheritable Class Shell32
        Public Enum SIGDN As UInt32
            DESKTOPABSOLUTEEDITING = &H8004C000L
            DESKTOPABSOLUTEPARSING = &H80028000L
            FILESYSPATH = &H80058000L
            NORMALDISPLAY = 0
            PARENTRELATIVEEDITING = &H80031001L
            PARENTRELATIVEFORADDRESSBAR = &H8001C001L
            PARENTRELATIVEPARSING = &H80018001L
            URL = &H80068000L
        End Enum
        Public Enum SHSTOCKICONID
            SIID_DOCNOASSOC = 0
            SIID_DOCASSOC = 1
            SIID_APPLICATION = 2
            SIID_FOLDER = 3
            SIID_FOLDEROPEN = 4
            SIID_DRIVE525 = 5
            SIID_DRIVE35 = 6
            SIID_DRIVEREMOVE = 7
            SIID_DRIVEFIXED = 8
            SIID_DRIVENET = 9
            SIID_DRIVENETDISABLED = 10
            SIID_DRIVECD = 11
            SIID_DRIVERAM = 12
            SIID_WORLD = 13
            SIID_SERVER = 15
            SIID_PRINTER = 16
            SIID_MYNETWORK = 17
            SIID_FIND = 22
            SIID_HELP = 23
            SIID_SHARE = 28
            SIID_LINK = 29
            SIID_SLOWFILE = 30
            SIID_RECYCLER = 31
            SIID_RECYCLERFULL = 32
            SIID_MEDIACDAUDIO = 40
            SIID_LOCK = 47
            SIID_AUTOLIST = 49
            SIID_PRINTERNET = 50
            SIID_SERVERSHARE = 51
            SIID_PRINTERFAX = 52
            SIID_PRINTERFAXNET = 53
            SIID_PRINTERFILE = 54
            SIID_STACK = 55
            SIID_MEDIASVCD = 56
            SIID_STUFFEDFOLDER = 57
            SIID_DRIVEUNKNOWN = 58
            SIID_DRIVEDVD = 59
            SIID_MEDIADVD = 60
            SIID_MEDIADVDRAM = 61
            SIID_MEDIADVDRW = 62
            SIID_MEDIADVDR = 63
            SIID_MEDIADVDROM = 64
            SIID_MEDIACDAUDIOPLUS = 65
            SIID_MEDIACDRW = 66
            SIID_MEDIACDR = 67
            SIID_MEDIACDBURN = 68
            SIID_MEDIABLANKCD = 69
            SIID_MEDIACDROM = 70
            SIID_AUDIOFILES = 71
            SIID_IMAGEFILES = 72
            SIID_VIDEOFILES = 73
            SIID_MIXEDFILES = 74
            SIID_FOLDERBACK = 75
            SIID_FOLDERFRONT = 76
            SIID_SHIELD = 77
            SIID_WARNING = 78
            SIID_INFO = 79
            SIID_ERROR = 80
            SIID_KEY = 81
            SIID_SOFTWARE = 82
            SIID_RENAME = 83
            SIID_DELETE = 84
            SIID_MEDIAAUDIODVD = 85
            SIID_MEDIAMOVIEDVD = 86
            SIID_MEDIAENHANCEDCD = 87
            SIID_MEDIAENHANCEDDVD = 88
            SIID_MEDIAHDDVD = 89
            SIID_MEDIABLURAY = 90
            SIID_MEDIAVCD = 91
            SIID_MEDIADVDPLUSR = 92
            SIID_MEDIADVDPLUSRW = 93
            SIID_DESKTOPPC = 94
            SIID_MOBILEPC = 95
            SIID_USERS = 96
            SIID_MEDIASMARTMEDIA = 97
            SIID_MEDIACOMPACTFLASH = 98
            SIID_DEVICECELLPHONE = 99
            SIID_DEVICECAMERA = 100
            SIID_DEVICEVIDEOCAMERA = 101
            SIID_DEVICEAUDIOPLAYER = 102
            SIID_NETWORKCONNECT = 103
            SIID_INTERNET = 104
            SIID_ZIPFILE = 105
            SIID_SETTINGS = 106
            SIID_DRIVEHDDVD = 132
            SIID_DRIVEBD = 133
            SIID_MEDIAHDDVDROM = 134
            SIID_MEDIAHDDVDR = 135
            SIID_MEDIAHDDVDRAM = 136
            SIID_MEDIABDROM = 137
            SIID_MEDIABDR = 138
            SIID_MEDIABDRE = 139
            SIID_CLUSTEREDDRIVE = 140
            SIID_MAX_ICONS = 174
        End Enum
        Public Enum StockIconInfoFlags As UInteger

            SHGSI_ICONLOCATION = 0
            SHGSI_ICON = &H100
            SHGSI_SYSICONINDEX = &H4000
            SHGSI_LINKOVERLAY = &H8000
            SHGSI_SELECTED = &H10000

            SHGSI_LARGEICON = &H0
            SHGSI_SMALLICON = &H1
            SHGSI_SHELLICONSIZE = &H4
        End Enum
        <Flags()>
        Public Enum SIIGBF
            SIIGBF_BIGGERSIZEOK = 1
            SIIGBF_ICONONLY = 4
            SIIGBF_INCACHEONLY = &H10
            SIIGBF_MEMORYONLY = 2
            SIIGBF_RESIZETOFIT = 0
            SIIGBF_THUMBNAILONLY = 8
        End Enum
        Public Enum ESTRRET As Integer

            eeRRET_WSTR = &H0 ',            // Use STRRET.pOleStr
            STRRET_OFFSET = &H1 ',    // Use STRRET.uOffset to Ansi
            STRRET_CSTR = &H2 '            // Use STRRET.cStr
        End Enum
        Public Enum ESHGDN

            SHGDN_NORMAL = &H0
            SHGDN_INFOLDER = &H1
            SHGDN_FOREDITING = &H1000
            SHGDN_FORADDRESSBAR = &H4000
            SHGDN_FORPARSING = &H8000
        End Enum
        Public Enum ESFGAO As Integer

            SFGAO_CANCOPY = &H1
            SFGAO_CANMOVE = &H2
            SFGAO_CANLINK = &H4
            SFGAO_LINK = &H10000
            SFGAO_SHARE = &H20000
            SFGAO_READONLY = &H40000
            SFGAO_HIDDEN = &H80000
            SFGAO_FOLDER = &H20000000
            SFGAO_FILESYSTEM = &H40000000
            SFGAO_HASSUBFOLDER = &H80000000
        End Enum
        Public Enum SHCONTF
            FOLDERS = 32
            NONFOLDERS = 64
            INCLUDEHIDDEN = 128
            INIT_ON_FIRST_NEXT = 256
            NETPRINTERSRCH = 512
            SHAREABLE = 1024
            STORAGE = 2048

        End Enum

        Public Shared IID_IShellItem As New Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")
        Public Shared IID_IShellFolder As New Guid("000214E6-0000-0000-C000-000000000046")

        Public Shared IID_IExtractImage As New Guid("BB2E617C-0920-11d1-9A0B-00C04FC2D6C1")

        <ComImport(), Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface IShellItem
            Sub BindToHandler(ByVal pbc As IntPtr, <MarshalAs(UnmanagedType.LPStruct)> ByVal bhid As Guid, <MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid, <Out()> ByRef ppv As IntPtr)
            Sub GetParent(<Out()> ByRef ppsi As IShellItem)
            Sub GetDisplayName(ByVal sigdnName As SIGDN, <Out()> ByRef ppszName As IntPtr)
            Sub GetAttributes(ByVal sfgaoMask As UInt32, <Out()> ByRef psfgaoAttribs As UInt32)
            Sub [Compare](ByVal psi As IShellItem, ByVal hint As UInt32, <Out()> ByRef piOrder As Integer)
        End Interface
        <ComImport(), Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface IShellItemImageFactory
            Function GetImage(<[In](), MarshalAs(UnmanagedType.Struct)> ByVal size As Size, <[In]()> ByVal flags As SIIGBF, <Out()> ByRef phbm As IntPtr) As Integer
        End Interface
        <ComImport()>
        <Guid("000214eb-0000-0000-c000-000000000046")>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface IExtractIcon
            <PreserveSig()>
            Function GetIconLocation(ByVal uFlags As Integer,
                    <MarshalAs(UnmanagedType.LPStr)>
                     ByRef szIconFile As String,
                     ByVal cchMax As Integer,
                     ByRef piIndex As Integer,
                     ByRef pwFlags As Integer) As IntPtr


            <PreserveSig()>
            Function Extract(ByVal pszFile As IntPtr,
                     ByVal nIconIndex As Integer,
                     ByVal phiconLarge As IntPtr,
                     ByVal phiconSmall As IntPtr,
                     ByVal nIconSize As Integer) As Integer
        End Interface
        <ComImportAttribute(),
        GuidAttribute("000214E6-0000-0000-C000-000000000046"),
        InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface IShellFolder

            Sub ParseDisplayName(
              ByVal hWnd As IntPtr,
              ByVal pbc As IntPtr,
              ByVal pszDisplayName As String,
              ByRef pchEaten As Integer,
              ByRef ppidl As System.IntPtr,
              ByRef pdwAttributes As Integer)

            Sub EnumObjects(
              ByVal hwndOwner As IntPtr,
              <MarshalAs(UnmanagedType.U4)> ByVal grfFlags As Integer,
              <Out()> ByRef ppenumIDList As IntPtr)

            Sub BindToObject(
              ByVal pidl As IntPtr,
              ByVal pbcReserved As IntPtr,
              ByRef riid As Guid,
              ByRef ppvOut As IShellFolder)

            Sub BindToStorage(
              ByVal pidl As IntPtr,
              ByVal pbcReserved As IntPtr,
              ByRef riid As Guid,
              <Out()> ByVal ppvObj As IntPtr)

            <PreserveSig()>
            Function CompareIDs(
              ByVal lParam As IntPtr,
              ByVal pidl1 As IntPtr,
              ByVal pidl2 As IntPtr) As Integer

            Sub CreateViewObject(
              ByVal hwndOwner As IntPtr,
              ByRef riid As Guid,
              ByVal ppvOut As Object)

            Sub GetAttributesOf(
              ByVal cidl As Integer,
              ByVal apidl As IntPtr,
              <MarshalAs(UnmanagedType.U4)> ByRef rgfInOut As Integer)

            Sub GetUIObjectOf(
              ByVal hwndOwner As IntPtr,
              ByVal cidl As Integer,
              ByRef apidl As IntPtr,
              ByRef riid As Guid,
              <Out()> ByVal prgfInOut As Integer,
              <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef ppvOut As Object)

            Sub GetDisplayNameOf(
              ByVal pidl As IntPtr,
              <MarshalAs(UnmanagedType.U4)> ByVal uFlags As Integer,
              ByRef lpName As STRRET_CSTR)

            Sub SetNameOf(
              ByVal hwndOwner As IntPtr,
              ByVal pidl As IntPtr,
              <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszName As String,
              <MarshalAs(UnmanagedType.U4)> ByVal uFlags As Integer,
              ByRef ppidlOut As IntPtr)

        End Interface
        <GuidAttribute("000214F9-0000-0000-C000-000000000046"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown), ComImport()>
        Public Interface IShellLink
            <PreserveSig()> Sub GetPath(
                                        <MarshalAs(UnmanagedType.LPWStr), Out()> ByVal pszFile As System.Text.StringBuilder,
                                       <[In]()> ByVal cch As Integer,
                                        <[In]()> ByVal pfd As ShellLinkFindData,
                                        <[In]()> ByVal fFlags As UInt32)

            <PreserveSig()> Sub GetIDList(<Out()> ByVal ppidl As IntPtr)
            <PreserveSig()> Sub SetIDList(ByVal pidl As IntPtr)
            <PreserveSig()> Sub GetDescription(<MarshalAs(UnmanagedType.LPWStr), Out()> ByVal pszArgs As String, ByVal cch As Integer)
            <PreserveSig()> Sub SetDescription(<MarshalAs(UnmanagedType.LPWStr, SizeParamIndex:=1)> ByVal pszName As String)
            <PreserveSig()> Sub GetWorkingDirectory(<MarshalAs(UnmanagedType.LPWStr, SizeParamIndex:=1), Out()> ByVal pszDir As String, ByVal cch As Integer)
            <PreserveSig()> Sub SetWorkingDirectory(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszDir As String)
            <PreserveSig()> Sub GetArguments(<MarshalAs(UnmanagedType.LPWStr, SizeParamIndex:=1), Out()> ByVal pszArgs As String, ByVal cch As Integer)
            <PreserveSig()> Sub SetArguments(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszArgs As String)
            <PreserveSig()> Sub GetHotkey(<Out()> ByVal pwHotKey As UShort)
            <PreserveSig()> Sub SetHotkey(ByVal wHotKey As UShort)
            <PreserveSig()> Sub GetShowCmd(<Out()> ByVal piShowCmd As Integer)
            <PreserveSig()> Sub SetShowCmd(ByVal iShowCmd As Integer)
            <PreserveSig()> Sub GetIconLocation(<MarshalAs(UnmanagedType.LPWStr, SizeParamIndex:=1), Out()> ByVal pszIconPath As String, ByVal cch As Integer, <Out()> ByVal piIcon As Integer)
            <PreserveSig()> Sub SetIconLocation(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszIconPath As String, ByVal iIcon As Integer)
            <PreserveSig()> Sub SetRelativePath(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszPathRel As String, ByVal dwReserved As UInteger)
            <PreserveSig()> Sub Resolve(ByVal hwnd As IntPtr, ByVal flags As UInteger)
            <PreserveSig()> Sub SetPath(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFile As String)

        End Interface

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure SHSTOCKICONINFO
            Public cbSize As UInt32
            Public hIcon As IntPtr
            Public iSysImageIndex As Int32
            Public iIcon As Int32
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public szPath As String
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure STRRET_CSTR
            Public uType As Integer
            <FieldOffset(4), MarshalAs(UnmanagedType.LPWStr)>
            Public pOleStr As String
            <FieldOffset(4)>
            Public uOffset As Integer
            <FieldOffset(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=520)>
            Public strName As Byte()
        End Structure
        <StructLayout(LayoutKind.Explicit, Size:=520)>
        Public Structure STRRETinternal

            <FieldOffset(0)>
            Public pOleStr As IntPtr

            <FieldOffset(0)>
            Public pStr As IntPtr ';  // LPSTR pStr;   NOT USED

            <FieldOffset(0)>
            Public uOffset As UInteger

        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure STRRET
            Public uType As UInteger
            Public data As STRRETinternal
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure ShellLinkFindData
            Public dwFileAttributes As Int32
            Public ftCreationTime As Runtime.InteropServices.ComTypes.FILETIME
            Public ftLastAccessTime As Runtime.InteropServices.ComTypes.FILETIME
            Public ftLastWriteTime As Runtime.InteropServices.ComTypes.FILETIME
            Public nFileSizeHigh As Int32
            Public nFileSizeLow As Int32
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
            Public cFileName As String

            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
            Public cAlternateFileName As String

            Public dwReserved0 As Int16

        End Structure

        <DllImport("Shell32")> Public Shared Function SHGetStockIconInfo(ByVal siid As SHSTOCKICONID, ByVal uFlags As UInteger, <[In](), Out()> ByRef psii As SHSTOCKICONINFO) As Integer

        End Function
        <DllImport("shell32")> Public Shared Function SHExtractIconsW(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String, ByVal nIconIndex As Integer, ByVal cxIcon As Integer, ByVal cyIcon As Integer, ByRef phIcon As IntPtr, ByRef pIconId As UInteger, ByVal nIcons As UInteger, ByVal flags As UInteger) As Integer

        End Function

        'Public Shared IID_IShellItem As New Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Public Shared Sub SHCreateItemFromParsingName(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszPath As String, <[In]()> ByVal pbc As IntPtr, <[In](), MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid, <Out()> ByRef ppv As IShellItem)

        End Sub

        '        <ComImport(), Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
        'Public Interface IShellItem
        '            Sub BindToHandler(ByVal pbc As IntPtr, <MarshalAs(UnmanagedType.LPStruct)> ByVal bhid As Guid, <MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid, <Out()> ByRef ppv As IntPtr)
        '            Sub GetParent(<Out()> ByRef ppsi As IShellItem)
        '            Sub GetDisplayName(ByVal sigdnName As SIGDN, <Out()> ByRef ppszName As IntPtr)
        '            Sub GetAttributes(ByVal sfgaoMask As UInt32, <Out()> ByRef psfgaoAttribs As UInt32)
        '            Sub [Compare](ByVal psi As IShellItem, ByVal hint As UInt32, <Out()> ByRef piOrder As Integer)
        '        End Interface
        '        <ComImport(), Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
        '    Public Interface IShellItemImageFactory
        '            Function GetImage(<[In](), MarshalAs(UnmanagedType.Struct)> ByVal size As Size, <[In]()> ByVal flags As SIIGBF, <Out()> ByRef phbm As IntPtr) As Integer
        '        End Interface

        <DllImport("shell32", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Public Shared Function SHGetDesktopFolder(<Out()> ByRef ppshf As IShellFolder) As Integer

        End Function




        <DllImport("Shell32")>
        Public Shared Function SHGetKnownFolderPath(<MarshalAs(UnmanagedType.LPStruct)> ByVal rfid As Guid, ByVal dwFlags As UInteger, ByVal hToken As IntPtr, ByRef pszPath As IntPtr) As Integer
        End Function



        <GuidAttribute("00021401-0000-0000-C000-000000000046"),
           ClassInterfaceAttribute(ClassInterfaceType.None),
           ComImportAttribute()>
        Public Class ShellLink
        End Class
        'Public Enum SIGDN As UInt32
        '    DESKTOPABSOLUTEEDITING = &H8004C000L
        '    DESKTOPABSOLUTEPARSING = &H80028000L
        '    FILESYSPATH = &H80058000L
        '    NORMALDISPLAY = 0
        '    PARENTRELATIVEEDITING = &H80031001L
        '    PARENTRELATIVEFORADDRESSBAR = &H8001C001L
        '    PARENTRELATIVEPARSING = &H80018001L
        '    URL = &H80068000L
        'End Enum
        '    <Flags()> _
        'Public Enum SIIGBF
        '        SIIGBF_BIGGERSIZEOK = 1
        '        SIIGBF_ICONONLY = 4
        '        SIIGBF_INCACHEONLY = &H10
        '        SIIGBF_MEMORYONLY = 2
        '        SIIGBF_RESIZETOFIT = 0
        '        SIIGBF_THUMBNAILONLY = 8
        '    End Enum

        '    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)> _
        'Public Shared Sub SHCreateItemFromParsingName(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal pszPath As String, <[In]()> ByVal pbc As IntPtr, <[In](), MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid, <Out()> ByRef ppv As IShellItem)

        '    End Sub
        'Public Enum SHSTOCKICONID
        '    SIID_DOCNOASSOC = 0
        '    SIID_DOCASSOC = 1
        '    SIID_APPLICATION = 2
        '    SIID_FOLDER = 3
        '    SIID_FOLDEROPEN = 4
        '    SIID_DRIVE525 = 5
        '    SIID_DRIVE35 = 6
        '    SIID_DRIVEREMOVE = 7
        '    SIID_DRIVEFIXED = 8
        '    SIID_DRIVENET = 9
        '    SIID_DRIVENETDISABLED = 10
        '    SIID_DRIVECD = 11
        '    SIID_DRIVERAM = 12
        '    SIID_WORLD = 13
        '    SIID_SERVER = 15
        '    SIID_PRINTER = 16
        '    SIID_MYNETWORK = 17
        '    SIID_FIND = 22
        '    SIID_HELP = 23
        '    SIID_SHARE = 28
        '    SIID_LINK = 29
        '    SIID_SLOWFILE = 30
        '    SIID_RECYCLER = 31
        '    SIID_RECYCLERFULL = 32
        '    SIID_MEDIACDAUDIO = 40
        '    SIID_LOCK = 47
        '    SIID_AUTOLIST = 49
        '    SIID_PRINTERNET = 50
        '    SIID_SERVERSHARE = 51
        '    SIID_PRINTERFAX = 52
        '    SIID_PRINTERFAXNET = 53
        '    SIID_PRINTERFILE = 54
        '    SIID_STACK = 55
        '    SIID_MEDIASVCD = 56
        '    SIID_STUFFEDFOLDER = 57
        '    SIID_DRIVEUNKNOWN = 58
        '    SIID_DRIVEDVD = 59
        '    SIID_MEDIADVD = 60
        '    SIID_MEDIADVDRAM = 61
        '    SIID_MEDIADVDRW = 62
        '    SIID_MEDIADVDR = 63
        '    SIID_MEDIADVDROM = 64
        '    SIID_MEDIACDAUDIOPLUS = 65
        '    SIID_MEDIACDRW = 66
        '    SIID_MEDIACDR = 67
        '    SIID_MEDIACDBURN = 68
        '    SIID_MEDIABLANKCD = 69
        '    SIID_MEDIACDROM = 70
        '    SIID_AUDIOFILES = 71
        '    SIID_IMAGEFILES = 72
        '    SIID_VIDEOFILES = 73
        '    SIID_MIXEDFILES = 74
        '    SIID_FOLDERBACK = 75
        '    SIID_FOLDERFRONT = 76
        '    SIID_SHIELD = 77
        '    SIID_WARNING = 78
        '    SIID_INFO = 79
        '    SIID_ERROR = 80
        '    SIID_KEY = 81
        '    SIID_SOFTWARE = 82
        '    SIID_RENAME = 83
        '    SIID_DELETE = 84
        '    SIID_MEDIAAUDIODVD = 85
        '    SIID_MEDIAMOVIEDVD = 86
        '    SIID_MEDIAENHANCEDCD = 87
        '    SIID_MEDIAENHANCEDDVD = 88
        '    SIID_MEDIAHDDVD = 89
        '    SIID_MEDIABLURAY = 90
        '    SIID_MEDIAVCD = 91
        '    SIID_MEDIADVDPLUSR = 92
        '    SIID_MEDIADVDPLUSRW = 93
        '    SIID_DESKTOPPC = 94
        '    SIID_MOBILEPC = 95
        '    SIID_USERS = 96
        '    SIID_MEDIASMARTMEDIA = 97
        '    SIID_MEDIACOMPACTFLASH = 98
        '    SIID_DEVICECELLPHONE = 99
        '    SIID_DEVICECAMERA = 100
        '    SIID_DEVICEVIDEOCAMERA = 101
        '    SIID_DEVICEAUDIOPLAYER = 102
        '    SIID_NETWORKCONNECT = 103
        '    SIID_INTERNET = 104
        '    SIID_ZIPFILE = 105
        '    SIID_SETTINGS = 106
        '    SIID_DRIVEHDDVD = 132
        '    SIID_DRIVEBD = 133
        '    SIID_MEDIAHDDVDROM = 134
        '    SIID_MEDIAHDDVDR = 135
        '    SIID_MEDIAHDDVDRAM = 136
        '    SIID_MEDIABDROM = 137
        '    SIID_MEDIABDR = 138
        '    SIID_MEDIABDRE = 139
        '    SIID_CLUSTEREDDRIVE = 140
        '    SIID_MAX_ICONS = 174
        'End Enum
        '<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        'Private Structure SHSTOCKICONINFO
        '    Public cbSize As UInt32
        '    Public hIcon As IntPtr
        '    Public iSysImageIndex As Int32
        '    Public iIcon As Int32
        '    <MarshalAs(UnmanagedType.ByValTStr, sizeconst:=260)> Public szPath As String
        'End Structure
        'Public Enum StockIconInfoFlags As UInteger

        '    SHGSI_ICONLOCATION = 0
        '    SHGSI_ICON = &H100
        '    SHGSI_SYSICONINDEX = &H4000
        '    SHGSI_LINKOVERLAY = &H8000
        '    SHGSI_SELECTED = &H10000

        '    SHGSI_LARGEICON = &H0
        '    SHGSI_SMALLICON = &H1
        '    SHGSI_SHELLICONSIZE = &H4
        'End Enum

        '<DllImport("Shell32")> Public Shared Function SHGetStockIconInfo(ByVal siid As SHSTOCKICONID, ByVal uFlags As UInteger, <[In](), Out()> ByRef psii As SHSTOCKICONINFO) As Integer

        'End Function
        '<DllImport("shell32")> Public Shared Function SHExtractIconsW(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String, ByVal nIconIndex As Integer, ByVal cxIcon As Integer, ByVal cyIcon As Integer, ByRef phIcon As IntPtr, ByRef pIconId As UInteger, ByVal nIcons As UInteger, ByVal flags As UInteger) As Integer

        'End Function
    End Class
    Public NotInheritable Class User32
#Region "ICONINFO.  Used by the CreateIconIndirect API function"
        Public Structure ICONINFO

            Public fIcon As Integer
            Public xHotspot As Integer
            Public yHotspot As Integer
            Public hBmMask As IntPtr
            Public hBmColor As IntPtr
        End Structure
#End Region

        <DllImport("user32")>
        Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean

        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Public Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As Integer

        End Function
        <DllImport("user32")>
        Public Shared Function CreateIconIndirect(
                 ByRef piconInfo As ICONINFO) As IntPtr
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Public Shared Function RegisterWindowMessage(ByVal msg As String) As Integer
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Public Shared Function FindWindow(ByVal className As String, ByVal windowName As String) As IntPtr
        End Function
        <DllImport("user32")>
        Public Shared Function ScreenToClient(ByVal hWnd As IntPtr, <[In](), Out()> ByRef lpPoint As Point) As Boolean
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
        Public Shared Function MapWindowPoints(ByVal hWndFrom As IntPtr, ByVal hWndTo As IntPtr, <[In](), Out()> ByRef pt As Point, ByVal cPoints As Integer) As Integer
        End Function
        <DllImport("User32")> Public Shared Function BeginPaint(ByVal hwnd As IntPtr, <Out()> ByRef lpPaint As PAINTSTRUCT) As IntPtr

        End Function
        <DllImport("user32")> Public Shared Function EndPaint(ByVal hwnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As Boolean
        End Function
        <StructLayout(LayoutKind.Sequential)>
        Public Structure PAINTSTRUCT

            Public hdc As IntPtr
            Public fErase As Boolean
            Public rcPaint As Win32.Windows.RECT
            Public fRestore As Boolean
            Public fIncUpdate As Boolean
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> Public rgbReserved As Byte()
        End Structure
        <DllImport("user32")>
        Public Shared Function GetClientRect(ByVal hWnd As IntPtr, ByRef lpRect As Win32.Windows.RECT) As Boolean
        End Function
        <DllImport("user32")>
        Public Shared Function GetScrollInfo(ByVal hwnd As IntPtr, ByVal fnBar As Integer, ByRef lpsi As SCROLLINFO) As Boolean
        End Function
        Public Structure SCROLLINFO
            Dim cbSize As UInteger
            Dim fMask As UInteger
            Dim nMin As Integer
            Dim nMax As Integer
            Dim nPage As UInteger
            Dim nPos As Integer
            Dim nTrackPos As Integer
        End Structure
        <Serializable(), StructLayout(LayoutKind.Sequential)>
        Public Structure MSG
            Public hwnd As IntPtr
            Public message As Integer
            Public wParam As IntPtr
            Public lParam As IntPtr
            Public time As Integer
            Public pt_x As Integer
            Public pt_y As Integer
        End Structure
        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Public Shared Function PeekMessage(<[In](), Out()> ByRef msg As MSG, ByVal hwnd As HandleRef, ByVal msgMin As Integer, ByVal msgMax As Integer, ByVal remove As Integer) As Boolean
        End Function
        <DllImport("user32")> Public Shared Function SetCapture(ByVal hWnd As IntPtr) As IntPtr

        End Function
        <DllImport("user32")> Public Shared Function ReleaseCapture() As Boolean

        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
        Public Shared Function GetCapture() As IntPtr
        End Function
        <DllImport("user32.dll")>
        Public Shared Function ShowScrollBar(ByVal hWnd As IntPtr, ByVal wBar As Integer, ByVal bShow As Integer) As Integer
        End Function
        <DllImport("user32")>
        Public Shared Function GetUpdateRect(ByVal hWnd As IntPtr, <Out()> ByRef lpRect As Win32.Windows.RECT, ByVal bErase As Boolean) As Integer
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
        Public Shared Function InvalidateRect(ByVal hWnd As IntPtr, ByRef rect As Win32.Windows.RECT, ByVal [erase] As Boolean) As Boolean
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
        Public Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal flags As Integer) As Boolean
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
        Public Shared Function UpdateWindow(ByVal hWnd As IntPtr) As Boolean
        End Function
        '        Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
        '(ByVal uAction As Integer, ByVal uParam As Integer, ByRef lpvParam As Integer, ByVal _
        'fuWinIni As Integer) As Integer
        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Public Shared Function SystemParametersInfo(ByVal uAction As Integer, ByVal uParam As Integer, ByRef lpvParam As Integer, ByVal fuWinIni As Integer) As Integer

        End Function
        <DllImport("user32")> Public Shared Function AnimateWindow(ByVal hwnd As IntPtr, ByVal dwTime As Int32, ByVal dwFlags As AnimateWindowFlags) As Boolean

        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function LoadImage(ByVal hinst As IntPtr,
 ByVal lpszName As String, ByVal uType As UInteger, ByVal cxDesired As Integer, ByVal cyDesired As Integer, ByVal fuLoad As UInteger) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function LoadImage(ByVal hinst As IntPtr,
     ByVal lpszName As Integer, ByVal uType As UInteger, ByVal cxDesired As Integer, ByVal cyDesired As Integer, ByVal fuLoad As UInteger) As IntPtr
        End Function
        <DllImport("user32.dll")>
        Public Shared Function LoadString(ByVal hInstance As IntPtr, ByVal uID As UInteger, ByVal lpBuffer As System.Text.StringBuilder, ByVal nBufferMax As Integer) As Integer
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Public Shared Function LoadString(ByVal hInstance As IntPtr, ByVal uID As UInteger, <MarshalAs(UnmanagedType.LPWStr)> ByRef s As String, ByVal nBufferMax As Integer) As Integer
        End Function
        '      <DllImport("user32")> _
        'Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean

        '      End Function

        Public Enum AnimateWindowFlags

            AW_HIDE = &H10000
            AW_ACTIVATE = &H20000
            AW_HOR_POSITIVE = &H1
            AW_HOR_NEGATIVE = &H2
            AW_SLIDE = &H40000
            AW_BLEND = &H80000
            AW_VER_POSITIVE = &H4
            AW_VER_NEGATIVE = &H8
            AW_CENTER = &H10
        End Enum
        Public Const SPI_GETKEYBOARDSPEED = 10&
        Public Const SPI_GETKEYBOARDDELAY = 22&
        'Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
        '(ByVal uAction As Integer, ByVal uParam As Integer, ByRef lpvParam As Integer, ByVal _
        'fuWinIni As Integer) As Integer
        <DllImport("user32.dll")>
        Public Shared Function GetSystemMetrics(ByVal smIndex As Int32) As Integer

        End Function
        Public Enum ExitWindowsFlags As Integer
            EWX_FORCE = 4
            EWX_FORCEIFHUNG = &H10
            EWX_LOGOFF = 0
            EWX_POWEROFF = &H8
            EWX_REBOOT = 2
            EWX_SHUTDOWN = 1
        End Enum
        Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As ExitWindowsFlags, ByVal dwReserved As Integer) As Integer

    End Class
    Public NotInheritable Class Gdi32
#Region "BITMAPINFOHEADER.  me is stored at the start of an icon's data"
        Public Structure BITMAPINFOHEADER

            Public biSize As Integer
            Public biWidth As Integer
            Public biHeight As Integer
            Public biPlanes As Int16
            Public biBitCount As Int16
            Public biCompression As Integer
            Public biSizeImage As Integer
            Public biXPelsPerMeter As Integer
            Public biYPelsPerMeter As Integer
            Public biClrUsed As Integer
            Public biClrImportant As Integer

            Public Sub New(ByVal size As Size, ByVal colorDepth As System.Windows.Forms.ColorDepth)


                Me.biSize = 0
                Me.biWidth = size.Width
                Me.biHeight = size.Height * 2
                Me.biPlanes = 1
                Me.biCompression = BI_RGB
                Me.biSizeImage = 0
                Me.biXPelsPerMeter = 0
                Me.biYPelsPerMeter = 0
                Me.biClrUsed = 0
                Me.biClrImportant = 0
                Select Case (colorDepth)

                    Case System.Windows.Forms.ColorDepth.Depth4Bit
                        Me.biBitCount = 4
                    Case System.Windows.Forms.ColorDepth.Depth8Bit
                        Me.biBitCount = 8
                    Case System.Windows.Forms.ColorDepth.Depth16Bit
                        Me.biBitCount = 16
                    Case System.Windows.Forms.ColorDepth.Depth24Bit
                        Me.biBitCount = 24
                    Case System.Windows.Forms.ColorDepth.Depth32Bit
                        Me.biBitCount = 32
                    Case Else
                        Me.biBitCount = 4
                End Select
                Me.biSize = Marshal.SizeOf(Me.GetType())
            End Sub

            Public Sub Write(ByVal bw As IO.BinaryWriter)

                bw.Write(Me.biSize)
                bw.Write(Me.biWidth)
                bw.Write(Me.biHeight)
                bw.Write(Me.biPlanes)
                bw.Write(Me.biBitCount)
                bw.Write(Me.biCompression)
                bw.Write(Me.biSizeImage)
                bw.Write(Me.biXPelsPerMeter)
                bw.Write(Me.biYPelsPerMeter)
                bw.Write(Me.biClrUsed)
                bw.Write(Me.biClrImportant)
            End Sub

            Public Sub New(ByVal data As Byte())

                Dim ms As New IO.MemoryStream(data, False)
                Dim br As New IO.BinaryReader(ms)
                biSize = br.ReadInt32()
                biWidth = br.ReadInt32()
                biHeight = br.ReadInt32()
                biPlanes = br.ReadInt16()
                biBitCount = br.ReadInt16()
                biCompression = br.ReadInt32()
                biSizeImage = br.ReadInt32()
                biXPelsPerMeter = br.ReadInt32()
                biYPelsPerMeter = br.ReadInt32()
                biClrUsed = br.ReadInt32()
                biClrImportant = br.ReadInt32()
                br.Close()
            End Sub

            Public Overrides Function ToString() As String

                Return String.Format(
                   "biSize: {0}, biWidth: {1}, biHeight: {2}, biPlanes: {3}, biBitCount: {4}, biCompression: {5}, biSizeImage: {6}, biXPelsPerMeter: {7}, biYPelsPerMeter {8}, biClrUsed {9}, biClrImportant {10}",
                   biSize, biWidth, biHeight, biPlanes, biBitCount,
                   biCompression, biSizeImage, biXPelsPerMeter, biYPelsPerMeter,
                   biClrUsed, biClrImportant)
            End Function
        End Structure
#End Region
        Public Const IMAGE_ICON As Int16 = 1
        Public Const BI_RGB As Integer = &H0
        Public Const DIB_RGB_COLORS As Integer = 0 '; //  color table in RGBs
#Region "RQBQUAD. Used to store colours in a paletised icon (2, 4 or 8 bit)"
        Public Structure RGBQUAD

            Public rgbBlue As Byte
            Public rgbGreen As Byte
            Public rgbRed As Byte
            Public rgbReserved As Byte

            Public Sub New(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal alpha As Byte)


                rgbBlue = b
                rgbGreen = g
                rgbRed = r
                rgbReserved = 0 '; //alpha;
            End Sub
            Public Sub New(ByVal c As Color)

                rgbBlue = c.B
                rgbGreen = c.G
                rgbRed = c.R
                rgbReserved = 0 '; //c.A;
            End Sub
            Public Sub Write(ByVal bw As IO.BinaryWriter)

                bw.Write(Me.rgbBlue)
                bw.Write(Me.rgbGreen)
                bw.Write(Me.rgbRed)
                bw.Write(Me.rgbReserved)
            End Sub
            Public Overrides Function ToString() As String

                Return String.Format(
                   "rgbBlue: {0}, rgbGreen: {1}, rgbRed: {2}",
                   rgbBlue, rgbGreen, rgbRed)
            End Function
        End Structure
#End Region
        Public Structure BITMAPINFO
            Public bmiHeader As BITMAPINFOHEADER
            Public bmiColors As RGBQUAD
        End Structure
        <DllImport("gdi32")>
        Public Shared Function SelectObject(
                 ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr

        End Function
        <DllImport("gdi32")>
        Public Shared Function CreateCompatibleBitmap(
                   ByVal hdc As IntPtr,
                   ByVal width As Integer,
                   ByVal height As Integer) As IntPtr
        End Function
        <DllImport("gdi32")>
        Public Shared Function CreateCompatibleDC(
               ByVal hdc As IntPtr) As IntPtr
        End Function
        <DllImport("gdi32")>
        Public Shared Function SetDIBitsToDevice(
            ByVal hdc As IntPtr,
            ByVal X As Integer, ByVal Y As Integer, ByVal dx As Integer, ByVal dy As Integer,
            ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal Scan As Integer, ByVal NumScans As Integer,
            ByVal Bits As IntPtr,
            ByVal BitsInfo As IntPtr,
            ByVal wUsage As Integer) As Integer
        End Function
        <DllImport("gdi32")>
        Public Shared Function DeleteDC(
                 ByVal hdc As IntPtr) As Integer
        End Function
        <DllImport("gdi32")>
        Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Integer
        End Function
        <DllImport("gdi32", CharSet:=CharSet.Auto)>
        Public Shared Function CreateDC(
            <MarshalAs(UnmanagedType.LPTStr)> ByVal _
            lpDriverName As String,
            ByVal lpDeviceName As IntPtr,
            ByVal lpOutput As IntPtr,
            ByVal lpInitData As IntPtr) As IntPtr
        End Function
        <DllImport("gdi32")>
        Public Shared Function GetDIBits(
               ByVal hdc As IntPtr,
               ByVal hBitmap As IntPtr,
               ByVal nStartScan As Integer,
               ByVal nNumScans As Integer,
               ByVal Bits As IntPtr,
               ByVal BitsInfo As IntPtr,
               ByVal wUsage As Integer) As Integer
        End Function
        <DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)>
        Public Shared Function CreateDIBSection(ByVal hdc As IntPtr, ByRef pbmi As BITMAPINFO, ByVal iUsage As UInteger, ByVal ppvBits As Integer, ByVal hSection As IntPtr, ByVal dwOffset As UInteger) As IntPtr
        End Function
        <DllImport("gdi32.dll")>
        Public Shared Function GetObject(ByVal hgdiobj As IntPtr, ByVal cbBuffer As Integer, <Out()> ByRef lpvObject As Win32.Gdi32.BITMAPINFOHEADER) As Integer
        End Function
        <DllImport("gdi32.dll")>
        Public Shared Function BitBlt(ByVal hdc As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As UInteger) As Boolean
        End Function
        <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
        Public Shared Function ReleaseDC(ByVal hdc As IntPtr, ByVal state As Integer) As Integer
        End Function
        <DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)>
        Public Shared Function RestoreDC(ByVal hdc As IntPtr, ByVal nSavedDC As Integer) As Boolean
        End Function
        <DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)>
        Public Shared Function SaveDC(ByVal hdc As IntPtr) As Integer
        End Function
        <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
        Public Shared Function GetDC(ByVal hdc As IntPtr) As IntPtr
        End Function
    End Class
    Public NotInheritable Class Windows
        Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Integer, ByVal DesiredAccess As Integer, ByRef TokenHandle As Integer) As Integer
        Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Integer
        Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Integer, ByVal DisableAllPrivileges As Integer, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Integer, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Integer) As Integer
        Private Const VER_PLATFORM_WIN32_NT = 2
        Private Const TOKEN_ADJUST_PRIVILEGES = &H20
        Private Const TOKEN_QUERY = &H8
        Private Const ANYSIZE_ARRAY = 1
        Private Const SE_PRIVILEGE_ENABLED = &H2
        Public Shared Function IsWinNT() As Boolean
            Dim myOS As New Win32.Kernel32.OSVERSIONINFO
            myOS.dwOSVersionInfoSize = Runtime.InteropServices.Marshal.SizeOf(myOS)
            Win32.Kernel32.GetVersionEx(myOS)
            IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
        End Function
        <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure LUID
            Dim LowPart As Integer
            Dim HighPart As Integer
        End Structure
        <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure LUID_AND_ATTRIBUTES
            Dim pLuid As LUID
            Dim Attributes As Integer
        End Structure
        <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure TOKEN_PRIVILEGES
            Dim PrivilegeCount As Integer
            <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=ANYSIZE_ARRAY)>
            Dim Privileges() As LUID_AND_ATTRIBUTES
        End Structure
        'Public Structure WTA_OPTIONS
        '    Dim dwFlags As WindowsThemeNonClientAttributeFlags
        '    Dim dwMask As WindowsThemeNonClientAttributeFlags
        'End Structure
        Public Structure RECT
            Public left As Integer
            Public top As Integer
            Public right As Integer
            Public bottom As Integer
            Public Sub New(ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer)
                Me.left = left
                Me.top = top
                Me.right = right
                Me.bottom = bottom
            End Sub

            Public Sub New(ByVal r As Rectangle)
                Me.left = r.Left
                Me.top = r.Top
                Me.right = r.Right
                Me.bottom = r.Bottom
            End Sub

            Public Shared Function FromXYWH(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer) As RECT
                Return New RECT(x, y, (x + width), (y + height))
            End Function

            Public ReadOnly Property Size() As Size
                Get
                    Return New Size((Me.right - Me.left), (Me.bottom - Me.top))
                End Get
            End Property
            Public Shared ReadOnly Property Empty() As RECT
                Get
                    Return New RECT
                End Get
            End Property

        End Structure

        Private Shared Sub EnableShutDown()
            Dim hProc As Integer
            Dim hToken As Integer
            Dim mLUID As New LUID
            Dim mPriv As New TOKEN_PRIVILEGES

            Dim mNewPriv As New TOKEN_PRIVILEGES
            hProc = Kernel32.GetCurrentProcess()
            OpenProcessToken(hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken)
            LookupPrivilegeValue("", "SeShutdownPrivilege", mLUID)
            mPriv.PrivilegeCount = ANYSIZE_ARRAY
            ReDim mPriv.Privileges(ANYSIZE_ARRAY)
            mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            mPriv.Privileges(0).pLuid = mLUID
            ' enable shutdown privilege for the current application
            Dim ret As Integer = AdjustTokenPrivileges(hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount))
        End Sub
        Public Shared Sub ExitWindows(ByVal Flags As User32.ExitWindowsFlags)
            'Dim ret As Integer
            If IsWinNT() Then
                EnableShutDown()
            End If
            User32.ExitWindowsEx(Flags, 0)
        End Sub
    End Class
    Public NotInheritable Class Kernel32
        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
        Public Shared Function WideCharToMultiByte(ByVal codePage As Integer, ByVal flags As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal wideStr As String, ByVal chars As Integer, <[In](), Out()> ByVal pOutBytes As Byte(), ByVal bufferBytes As Integer, ByVal defaultChar As IntPtr, ByVal pDefaultUsed As IntPtr) As Integer
        End Function
        Public Declare Function GetCurrentProcess Lib "kernel32" () As Integer

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, ExactSpelling:=True)>
        Public Shared Function MultiByteToWideChar(ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As Byte(), ByVal cchMultiByte As Integer, ByVal lpWideCharStr As Char(), ByVal cchWideChar As Integer) As Integer
        End Function
        <DllImport("kernel32")>
        Public Shared Function CopyMemory(ByVal Destination As IntPtr, ByVal Source As IntPtr, ByVal Length As Integer) As Integer
        End Function
        <DllImport("kernel32", CharSet:=CharSet.Auto)>
        Public Shared Function LoadLibraryEx(<MarshalAs(UnmanagedType.LPTStr)> ByVal lpLibFileName As String, ByVal hFile As IntPtr, ByVal dwFlags As Integer) As IntPtr

        End Function
        <DllImport("kernel32", CharSet:=CharSet.Auto)>
        Public Shared Function LoadLibrary(<MarshalAs(UnmanagedType.LPTStr)> ByVal lpLibFileName As String) As IntPtr

        End Function
        <DllImport("kernel32")>
        Public Shared Function LoadResource(ByVal hInstance As IntPtr, ByVal hResInfo As IntPtr) As IntPtr
        End Function
        <DllImport("kernel32", CharSet:=CharSet.Auto)>
        Public Shared Function FindResource(ByVal hInstance As IntPtr,
            <MarshalAs(UnmanagedType.LPTStr)> ByVal lpName As String,
            <MarshalAs(UnmanagedType.LPTStr)> ByVal lpType As String) As IntPtr
        End Function

        <DllImport("kernel32", CharSet:=CharSet.Auto)>
        Public Shared Function FindResource(ByVal hInstance As IntPtr,
<MarshalAs(UnmanagedType.LPTStr)> ByVal lpName As String, ByVal lpType As IntPtr) As IntPtr
        End Function
        Public Enum ResourceType

            RT_CURSOR = 1
            RT_BITMAP = 2
            RT_ICON = 3
            RT_MENU = 4
            RT_DIALOG = 5
            RT_STRING = 6
            RT_FONTDIR = 7
            RT_FONT = 8
            RT_ACCELERATOR = 9
            RT_RCDATA = 10
            RT_MESSAGETABLE = 11
            RT_VERSION = 16
            RT_DLGINCLUDE = 17
            RT_PLUGPLAY = 19
            RT_VXD = 20
            RT_ANICURSOR = 21
            RT_ANIICON = 22
            RT_HTML = 23
            RT_MANIFEST = 24
            RT_GROUP_CURSOR = ResourceType.RT_CURSOR + DIFFERENCE
            RT_GROUP_ICON = ResourceType.RT_ICON + DIFFERENCE
            CustomDefined = 10
        End Enum
        Public Const DIFFERENCE As Integer = 11
        Public Const DONT_RESOLVE_DLL_REFERENCES As Integer = &H1
        Public Const LOAD_LIBRARY_AS_DATAFILE As Integer = &H2
        Public Const LOAD_WITH_ALTERED_SEARCH_PATH As Integer = &H8
        Public Const LOAD_IGNORE_CODE_AUTHZ_LEVEL As Integer = &H10

        <DllImport("kernel32", CharSet:=CharSet.Auto)>
        Public Shared Function FindResource(ByVal hInstance As IntPtr,
        <MarshalAs(UnmanagedType.LPTStr)> ByVal lpName As String, ByVal lpType As ResourceType) As IntPtr
        End Function
        <DllImport("kernel32")>
        Public Shared Function SizeofResource(ByVal hInstance As IntPtr, ByVal hResInfo As IntPtr) As Integer
        End Function
        <DllImport("kernel32")>
        Public Shared Function FreeResource(ByVal hResData As IntPtr) As Integer
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function LoadLibraryEx(ByVal lpfFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As IntPtr
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function EnumResourceTypes(ByVal hModule As IntPtr, ByVal callback As EnumResTypeProc, ByVal lParam As Long) As Boolean
        End Function
        Public Delegate Function EnumResTypeProc(ByVal hModule As IntPtr,
        <MarshalAs(UnmanagedType.LPStr)> ByVal lpszType As String, ByVal lParam As Long) As Boolean
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function EnumResourceNames(ByVal hModule As IntPtr, ByVal lpszType As ResourceType, ByVal callback As EnumResNameProc, ByVal lParam As Long) As Boolean
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function EnumResourceNames(ByVal hModule As IntPtr, ByVal lpszType As IntPtr, ByVal callback As EnumResNameProc, ByVal lParam As Long) As Boolean
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function EnumResourceNames(ByVal hModule As IntPtr,
        <MarshalAs(UnmanagedType.LPStr)> ByVal lpszType As String, ByVal callback As EnumResNameProc, ByVal lParam As Long) As Boolean
        End Function
        Public Delegate Function EnumResNameProc(ByVal hModule As IntPtr,
<MarshalAs(UnmanagedType.LPStr)> ByVal lpszType As String,
<MarshalAs(UnmanagedType.LPStr)> ByVal lpszName As String, ByVal lParam As Long) As Boolean
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function FindResourceEx(ByVal hModule As IntPtr,
<MarshalAs(UnmanagedType.LPStr)> ByVal lpType As String,
<MarshalAs(UnmanagedType.LPStr)> ByVal lpName As String, ByVal wLanguage As Long) As IntPtr
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function LockResource(ByVal hResData As IntPtr) As IntPtr

        End Function
        <DllImport("kernel32.dll", SetLastError:=True)> Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As Integer

        End Function
        <DllImport("kernel32.dll")> Public Shared Function GetProcAddress(ByVal ModuleHandle As IntPtr, ByVal ProcName As String) As IntPtr
        End Function
        <StructLayout(LayoutKind.Sequential)>
        Public Structure OSVERSIONINFOEX

            Public dwOSVersionInfoSize As Int32
            Public dwMajorVersion As Int32
            Public dwMinorVersion As Int32
            Public dwBuildNumber As Int32
            Public dwPlatformId As Int32
            <MarshalAs(UnmanagedType.LPTStr, SizeConst:=128)> Public szCSDVersion As String
            Public wServicePackMajor As Int16
            Public wServicePackMinor As Int16
            Public wSuiteMask As Int16
            Public wProductType As Byte
            Public wReserved As Byte
        End Structure

        Public Structure OSVERSIONINFO
            Public dwOSVersionInfoSize As Integer
            Public dwMajorVersion As Integer
            Public dwMinorVersion As Integer
            Public dwBuildNumber As Integer
            Public dwPlatformId As Integer
            <MarshalAs(UnmanagedType.LPWStr, SizeConst:=128)> Public szCSDVersion As String
        End Structure
        <DllImport("kernel32")> Public Shared Function GetVersionEx(ByVal lpVersionInformation As OSVERSIONINFO) As Integer

        End Function
        <DllImport("kernel32")> Public Shared Function GetVersionEx(ByVal lpVersionInformation As OSVERSIONINFOEX) As Integer

        End Function
    End Class
End Namespace


Public NotInheritable Class Resources
    Implements IDisposable
    Dim hLibrary As IntPtr = IntPtr.Zero
    '''<summary>Reads string resources from Win32 resource files.</summary>

#Region "Constructors, Destructors"
    Public Sub New()
    End Sub
    Public Sub New(ByVal LibPath As String)
        LoadTheLibrary(LibPath)
    End Sub

    Protected Overrides Sub Finalize()
        UnloadLibrary()
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        UnloadLibrary()
    End Sub
    Public Sub UnloadLibrary()
        If (IntPtr.Zero <> hLibrary) Then
            Win32.Kernel32.FreeLibrary(hLibrary)
        End If
        hLibrary = IntPtr.Zero
    End Sub
#End Region
#Region "Win32 Declarations"
    '//[DllImport("kernel32", CharSet = CharSet.Auto)] 
    '//private extern static IntPtr FindResource(IntPtr hInstance, [MarshalAs(UnmanagedType.LPTStr)] string lpName, IntPtr lpType);

    '//[DllImport("kernel32", CharSet = CharSet.Auto)] 
    '//private extern static IntPtr FindResource(IntPtr hInstance, [MarshalAs(UnmanagedType.LPTStr)] string lpName, IntPtr lpType);

#End Region
#Region "Win32 Constants"
#Region "Enums"
    '//These are actual resource types taken from the Win32 API...


#End Region
    '//Enum types
    '//uType
    'Private Const IMAGE_BITMAP As UInteger = 0
    'Private Const IMAGE_ICON As UInteger = 1
    'Private Const IMAGE_CURSOR As UInteger = 2
    '//fuLoad
    'Private Const LR_DEFAULTCOLOR As UInteger = &H0
    'Private Const LR_MONOCHROME As UInteger = &H1
    'Private Const LR_COLOR As UInteger = &H2
    'Private Const LR_COPYRETURNORG As UInteger = &H4
    'Private Const LR_COPYDELETEORG As UInteger = &H8
    'Private Const LR_LOADFROMFILE As UInteger = &H10
    'Private Const LR_LOADTRANSPARENT As UInteger = &H20
    'Private Const LR_DEFAULTSIZE As UInteger = &H40
    'Private Const LR_VGACOLOR As UInteger = &H80
    'Private Const LR_LOADMAP3DCOLORS As UInteger = &H1000
    'Private Const LR_CREATEDIBSECTION As UInteger = &H2000
    'Private Const LR_COPYFROMRESOURCE As UInteger = &H4000
    'Private Const LR_SHARED As UInteger = &H8000
    '//Languages
    'Private Const LANG_NEUTRAL As Integer = &H0


#End Region
#Region "Private Methods"
    Private Shared Function isIntResource(ByVal value As IntPtr) As Boolean
        If (CUInt(value) > UShort.MaxValue) Then
            Return False
        End If
        Return True
    End Function
    Private Shared Function resourceID(ByVal value As IntPtr) As UInteger
        If (isIntResource(value)) Then
            Return CUInt(value)
        End If
        Throw New System.NotSupportedException("value is not an ID!")
    End Function
    Private Shared Function resourceName(ByVal value As IntPtr) As String
        If (CUInt(value) > UShort.MaxValue) Then
            Return Marshal.PtrToStringAuto(CType(value, IntPtr))
        Else
            Return value.ToString()
        End If
    End Function
    'Private Function EnumerateResourceNames(ByVal hModule As IntPtr, ByVal lpszType As String, ByVal lpszName As String, ByVal lParam As Long) As Boolean
    '    System.Diagnostics.Debug.WriteLine(lpszType + ":" + lpszName)
    '    Return True
    'End Function
    'Private Function EnumerateResourceTypes(ByVal hModule As IntPtr, ByVal lpszType As String, ByVal lParam As Long) As Boolean
    '    System.Diagnostics.Debug.WriteLine("Res Type: " + lpszType)
    '    Return True
    'End Function
    '//lParam is some parameter we may have wanted to pass ourselves!
    '//lpszType and lpszName are both strings b/c in the P/Invoke call, they were
    '//specified with the attribute [MarshalAs(UnmmangedType.LPTStr)] which converts an LPStr
    '//to a string for us. Make sure it's LPTStr and not LPStr or anything else! This ensures
    '//that it's platform-independent! Yeah! (c:
    '//lParam is some parameter we may have wanted to pass ourselves!
#End Region
#Region "Methods"
    '//Courtesy vbAccelerator!
    Public Sub LoadTheLibrary(ByVal libraryFile As String)
        Try
            hLibrary = Win32.Kernel32.LoadLibraryEx(libraryFile, IntPtr.Zero, Win32.Kernel32.LOAD_LIBRARY_AS_DATAFILE)
            If (hLibrary = IntPtr.Zero) Then
                Throw New Exception("Cannot load library: " + libraryFile)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Public Sub EnumerateResourceTypes()
    '    If (hLibrary = IntPtr.Zero) Then
    '        Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
    '    End If
    '    If (EnumResourceTypes(hLibrary, New EnumResTypeProc(AddressOf EnumerateResourceTypes), 0) = False) Then

    '    End If
    'End Sub
    Public Function EnumerateResourceTypes(ByVal e As Win32.Kernel32.EnumResTypeProc) As Boolean
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Return Win32.Kernel32.EnumResourceTypes(hLibrary, e, 0)
    End Function
    'Public Sub EnumerateResources(ByVal CustomResourceType As String)
    '    If (hLibrary = IntPtr.Zero) Then
    '        Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
    '    End If
    '    If (EnumResourceNames(hLibrary, CustomResourceType, New EnumResNameProc(AddressOf EnumerateResourceNames), 0) = False) Then
    '    End If

    'End Sub
    'Public Sub EnumerateResources(ByVal CustomResourceType As ResourceType)
    '    If (hLibrary = IntPtr.Zero) Then
    '        Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
    '    End If
    '    If (EnumResourceNames(hLibrary, CustomResourceType, New EnumResNameProc(AddressOf EnumerateResourceNames), 0) = False) Then
    '    End If

    'End Sub
    Public Function EnumerateResources(ByVal CustomResourceType As String, ByVal e As Win32.Kernel32.EnumResNameProc) As Boolean
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Return Win32.Kernel32.EnumResourceNames(hLibrary, CustomResourceType, e, 0)
    End Function
    Public Function EnumerateResources(ByVal CustomResourceType As Win32.Kernel32.ResourceType, ByVal e As Win32.Kernel32.EnumResNameProc) As Boolean
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Return Win32.Kernel32.EnumResourceNames(hLibrary, CustomResourceType, e, 0)
    End Function
    Public Function EnumerateResources(ByVal CustomResourceType As IntPtr, ByVal e As Win32.Kernel32.EnumResNameProc) As Boolean
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Return Win32.Kernel32.EnumResourceNames(hLibrary, CustomResourceType, e, 0)
    End Function
    'Public Function GetStringResource(ByVal ResourceName As String, ByVal ResourceType As String) As String
    '    Dim tmp As Byte() = GetResource(ResourceName, ResourceType)
    '    If (tmp IsNot Nothing) Then
    '        Return System.Text.Encoding.Default.GetString(tmp)
    '    Else
    '        Return ""
    '    End If
    'End Function
    Public Function GetStringResource(ByVal nIndex As UInteger) As String
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        'Dim s As New System.Text.StringBuilder(256)
        'Dim i = Win32.User32.LoadString(hLibrary, nIndex, s, 255 + 1)

        Dim s As String = ""
        'GetStringResource = s.ToString.Substring(0, i)
        Dim ptrToString As IntPtr = Win32.User32.LoadString(hLibrary, nIndex, s, 0)
        Return s.Substring(0, ptrToString)
    End Function
    'Private Function LoadString(ByVal nID As Integer) As String
    '    '// try fixed buffer first (to avoid wasting space in the heap)
    '    Dim szTemp(256) As Char
    '    Dim nCount As Integer = Marshal.SizeOf(szTemp) / Marshal.SizeOf(szTemp(0))
    '    Dim nLen As Integer = Win32.User32.LoadString(nID, szTemp, nCount)
    '    If (nCount - nLen > 1) Then

    '        '*this = szTemp;
    '        '        Return nLen > 0
    '        Return szTemp
    '    End If

    '    '// try buffer size of 512, then larger size until entire string is retrieved
    '    Dim nSize As Integer = 256
    '    Do

    '        nSize += 256
    '        nLen = Win32.User32.LoadString(nID, GetBuffer(nSize - 1), nSize)
    '    Loop While (nSize - nLen <= 1)
    '    ReleaseBuffer()

    '    Return nLen > 0

    'End Function

    'Public Function GetIconResource(ByVal DesiredSize As Size) As Icon
    '    If (hLibrary = IntPtr.Zero) Then
    '        Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
    '    End If
    '    Dim i = LoadImage(hLibrary, MAKEINTRESOURCE(1000), IMAGE_ICON, DesiredSize.Width, DesiredSize.Height, 0)
    '    If i = IntPtr.Zero Then
    '        Throw New Win32Exception(Marshal.GetLastWin32Error)
    '    End If
    '    Dim iz As Icon = Icon.FromHandle(i)
    '    DestroyIcon(i)
    '    Return iz
    'End Function
    Public ReadOnly Property ResourceExists(ByVal ResourceName As String, ByVal ResourceType As String) As Boolean
        Get
            If (hLibrary = IntPtr.Zero) Then
                Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
            End If
            Dim hRsrc = Win32.Kernel32.FindResource(hLibrary, ResourceName, ResourceType)
            If (hRsrc <> IntPtr.Zero) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property ResourceExists(ByVal ResourceName As String, ByVal ResourceType As Win32.Kernel32.ResourceType) As Boolean
        Get
            If (hLibrary = IntPtr.Zero) Then
                Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
            End If
            Dim hRsrc = Win32.Kernel32.FindResource(hLibrary, ResourceName, ResourceType)
            If (hRsrc <> IntPtr.Zero) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property ResourceExists(ByVal ResourceName As String, ByVal ResourceType As IntPtr) As Boolean
        Get
            If (hLibrary = IntPtr.Zero) Then
                Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
            End If
            Dim hRsrc = Win32.Kernel32.FindResource(hLibrary, ResourceName, ResourceType)
            If (hRsrc <> IntPtr.Zero) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public Function GetResource(ByVal ResourceName As String, ByVal ResourceType As String) As Byte()
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Dim msg As String = ""
        Dim failed As Boolean = False
        Dim hGlobal As IntPtr = IntPtr.Zero
        Dim hRsrc As IntPtr = IntPtr.Zero
        Dim lPtr As IntPtr = IntPtr.Zero
        Dim ret As Byte() = Nothing
        Try
            hRsrc = Win32.Kernel32.FindResource(hLibrary, ResourceName, ResourceType)
            If (hRsrc <> IntPtr.Zero) Then
                Dim size As Integer = Win32.Kernel32.SizeofResource(hLibrary, hRsrc)
                hGlobal = Win32.Kernel32.LoadResource(hLibrary, hRsrc)
                If (hGlobal <> IntPtr.Zero) Then
                    lPtr = Win32.Kernel32.LockResource(hGlobal)
                    If (lPtr <> IntPtr.Zero) Then
                        Dim stuff(size) As Byte
                        Marshal.Copy(lPtr, stuff, 0, size)
                        ret = stuff

                    Else
                        msg = "Can't lock resource for reading."
                        failed = True
                    End If

                Else
                    msg = "Can't load resource for reading."
                    failed = True
                End If

            Else
                msg = "Can't find resource."
                failed = True
            End If

        Catch ex As Exception
            failed = True
            msg = ex.Message

        Finally
            If (hGlobal <> IntPtr.Zero) Then
                Win32.Kernel32.FreeResource(hGlobal)
            End If
            If (failed) Then
                Throw New Exception(msg)
            End If
        End Try
        Return ret
    End Function
    Public Function GetResource(ByVal ResourceName As String, ByVal ResourceType As Win32.Kernel32.ResourceType) As Byte()
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Dim msg As String = ""
        Dim failed As Boolean = False
        Dim hGlobal As IntPtr = IntPtr.Zero
        Dim hRsrc As IntPtr = IntPtr.Zero
        Dim lPtr As IntPtr = IntPtr.Zero
        Dim ret As Byte() = Nothing
        Try
            hRsrc = Win32.Kernel32.FindResource(hLibrary, ResourceName, ResourceType)
            If (hRsrc <> IntPtr.Zero) Then
                Dim size As Integer = Win32.Kernel32.SizeofResource(hLibrary, hRsrc)
                hGlobal = Win32.Kernel32.LoadResource(hLibrary, hRsrc)
                If (hGlobal <> IntPtr.Zero) Then
                    lPtr = Win32.Kernel32.LockResource(hGlobal)
                    If (lPtr <> IntPtr.Zero) Then
                        Dim stuff(size) As Byte
                        Marshal.Copy(lPtr, stuff, 0, size)
                        ret = stuff

                    Else
                        msg = "Can't lock resource for reading."
                        failed = True
                    End If

                Else
                    msg = "Can't load resource for reading."
                    failed = True
                End If

            Else
                msg = "Can't find resource."
                failed = True
            End If

        Catch ex As Exception
            failed = True
            msg = ex.Message

        Finally
            If (hGlobal <> IntPtr.Zero) Then
                Win32.Kernel32.FreeResource(hGlobal)
            End If
            If (failed) Then
                Throw New Exception(msg)
            End If
        End Try
        Return ret
    End Function
    Public Function GetResource(ByVal ResourceName As String, ByVal ResourceType As IntPtr) As Byte()
        If (hLibrary = IntPtr.Zero) Then
            Throw New Exception("You must call LoadLibrary() to initialize the desired resource.")
        End If
        Dim msg As String = ""
        Dim failed As Boolean = False
        Dim hGlobal As IntPtr = IntPtr.Zero
        Dim hRsrc As IntPtr = IntPtr.Zero
        Dim lPtr As IntPtr = IntPtr.Zero
        Dim ret As Byte() = Nothing
        Try
            hRsrc = Win32.Kernel32.FindResource(hLibrary, ResourceName, ResourceType)
            If (hRsrc <> IntPtr.Zero) Then
                Dim size As Integer = Win32.Kernel32.SizeofResource(hLibrary, hRsrc)
                hGlobal = Win32.Kernel32.LoadResource(hLibrary, hRsrc)
                If (hGlobal <> IntPtr.Zero) Then
                    lPtr = Win32.Kernel32.LockResource(hGlobal)
                    If (lPtr <> IntPtr.Zero) Then
                        Dim stuff(size) As Byte
                        Marshal.Copy(lPtr, stuff, 0, size)
                        ret = stuff

                    Else
                        msg = "Can't lock resource for reading."
                        failed = True
                    End If

                Else
                    msg = "Can't load resource for reading."
                    failed = True
                End If

            Else
                msg = "Can't find resource."
                failed = True
            End If

        Catch ex As Exception
            failed = True
            msg = ex.Message

        Finally
            If (hGlobal <> IntPtr.Zero) Then
                Win32.Kernel32.FreeResource(hGlobal)
            End If
            If (failed) Then
                Throw New Exception(msg)
            End If
        End Try
        Return ret
    End Function
    '//Indicates that you just want access to the resources within the library and dont want 
    '//to execute any code it may contain...
    '//Couldn't find anything!
    '//throw new Exception("Could not enumerate library types. Win32 Error: " + Marshal.GetLastWin32Error());
    '//Couldn't find anything with that type of resource!
    '//throw new Exception("Could not enumerate library resources. Win32 Error: " + Marshal.GetLastWin32Error());
    '//The return type of LoadResource is HGLOBAL for backward compatibility, not 
    '//because the function returns a handle to a global memory block. Do not pass 
    '//this handle to the GlobalLock or GlobalFree function. To obtain a pointer to 
    '//the resource data, call the LockResource function.
    '//Clear up handles:
#End Region





    '    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), _
    'Guid("000214E6-0000-0000-C000-000000000046")> _
    'Public Interface IShellFolder

    '        <PreserveSig()> _
    '        Function ParseDisplayName(ByVal hwnd As IntPtr, ByVal pbc As IntPtr, _
    '        <MarshalAs(UnmanagedType.LPWStr)> ByVal pszDisplayName As String, ByRef _
    '        pchEaten As Integer, <Out()> ByRef ppidl As IntPtr, ByRef pdwAttributes As Integer) _
    'As Integer

    '        <PreserveSig()> _
    '        Function EnumObjects(ByVal hwnd As IntPtr, ByVal grfFlags As  _
    '        SHCONTF, <Out()> ByRef ppenumIDList As IntPtr) As Int32

    '        <PreserveSig()> _
    '        Function BindToObject(ByVal pidl As IntPtr, ByVal pbc As IntPtr, <[In]()> ByRef _
    '        riid As Guid, <Out()> ByRef ppv As IShellFolder) As Int32

    '        <PreserveSig()> _
    '        Function BindToStorage(ByVal pidl As IntPtr, ByVal pbc As IntPtr, ByRef _
    '        riid As Guid, ByRef ppv As IntPtr) As Int32

    '        <PreserveSig()> _
    '        Function CompareIDs(ByVal lParam As Int32, ByVal pidl1 As IntPtr, ByVal _
    '        pidl2 As IntPtr) As Int32

    '        <PreserveSig()> _
    '        Function CreateViewObject(ByVal hwndOwner As IntPtr, ByVal riid As Guid, _
    '        ByRef ppv As IntPtr) As Int32

    '        <PreserveSig()> _
    '        Function GetAttributesOf(ByVal cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByRef apidl() As IntPtr, _
    '        ByRef rgfInOut As ESFGAO) As Integer

    '        '<PreserveSig()> _
    '        'Function GetUIObjectOf(ByVal hwndOwner As IntPtr, ByVal cidl As UInt32, _
    '        '<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal apidl() As IntPtr, ByVal riid As Guid, <[In]()> ByRef rgfReserved As UInt32, _
    '        '<Out()> ByRef ppv As STRRET) As Int32
    '        <PreserveSig()> _
    '   Function GetUIObjectOf(ByVal hwndOwner As IntPtr, ByVal cidl As UInt32, _
    '   <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal apidl() As IntPtr, ByVal riid As Guid, <[In]()> ByRef rgfReserved As UInt32, _
    '   <Out()> ByRef ppv As IShellFolder) As Int32

    '        <PreserveSig()> _
    '        Function GetDisplayNameOf(ByVal pidl As IntPtr, ByVal uFlags As ESHGDN, _
    '        <Out()> ByRef pName As STRRET) As Int32

    '        <PreserveSig()> _
    '        Function SetNameOf(ByVal hwnd As IntPtr, ByVal pidl As IntPtr, _
    '        <MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As [String], ByVal uFlags As  _
    '        SHCONTF, <Out()> ByRef ppidlOut As IntPtr) As Int32
    '    End Interface

    'Public Shared Function ShellDisplayText(ByVal filename As String) As String
    '    Dim ppsi As Win32.Shell32.IShellItem = Nothing
    '    Try
    '        If filename <> "" Then filename = IO.Path.GetFullPath(filename)
    '        Win32.Shell32.SHCreateItemFromParsingName(filename, IntPtr.Zero, Win32.Shell32.IID_IShellItem, ppsi)
    '    Catch
    '        Return IO.Path.GetFileName(filename)
    '    End Try
    '    Dim iz As IntPtr
    '    ppsi.GetDisplayName(Win32.Shell32.SIGDN.NORMALDISPLAY, iz)
    '    ShellDisplayText = Marshal.PtrToStringUni(iz)
    '    Marshal.ReleaseComObject(ppsi)
    'End Function

    'Public Shared Function StockIcon(ByVal stockID As Win32.Shell32.SHSTOCKICONID) As Icon
    '    Dim ps As New Win32.Shell32.SHSTOCKICONINFO
    '    Dim flags As Win32.Shell32.StockIconInfoFlags = Win32.Shell32.StockIconInfoFlags.SHGSI_ICONLOCATION
    '    ps.cbSize = Marshal.SizeOf(GetType(Win32.Shell32.SHSTOCKICONINFO))
    '    'ps.hIcon = IntPtr.Zero
    '    Win32.Shell32.SHGetStockIconInfo(stockID, flags, ps)
    '    'StockIcon = Icon.FromHandle(ps.hIcon).ToBitmap
    '    'DestroyIcon(ps.hIcon)
    '    Dim z As IntPtr, zz As IntPtr
    '    Win32.Shell32.SHExtractIconsW(ps.szPath, ps.iIcon, 128, 128, z, zz, 1, 0)
    '    StockIcon = Icon.FromHandle(z).Clone
    '    Win32.User32.DestroyIcon(z)
    'End Function
    '   Public Shared Function ShellThumbnail(ByVal filename As String, ByVal size As Size) As Bitmap
    '       Dim ppsi As Win32.Shell32.IShellItem = Nothing
    '       Dim hbitmap As IntPtr = IntPtr.Zero
    '       'Dim uuid As New Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")
    '       If filename <> "" Then filename = IO.Path.GetFullPath(filename)
    '       Win32.Shell32.SHCreateItemFromParsingName(filename, IntPtr.Zero, Win32.Shell32.IID_IShellItem, ppsi)
    '       'Dim iz As IntPtr
    '       'ppsi.GetDisplayName(SIGDN.NORMALDISPLAY, iz)
    '       'DisplayName = Marshal.PtrToStringUni(iz)

    '       Dim ret = DirectCast(ppsi, Win32.Shell32.IShellItemImageFactory).GetImage(size, _
    'Win32.Shell32.SIIGBF.SIIGBF_BIGGERSIZEOK, hbitmap)
    '       'Dim source As Bitmap = Bitmap.FromHbitmap(hbitmap)
    '       'source.MakeTransparent()
    '       'Dim source As Bitmap = Icon.FromHandle(hbitmap).ToBitmap
    '       'Dim hBitmap As IntPtr = ConvertPixelByPixel(hbitmap)
    '       Dim asd As System.Windows.Media.Imaging.BitmapSizeOptions = System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions

    '       Dim source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, System.Windows.Int32Rect.Empty, asd)
    '       Dim stride As Integer = source.PixelWidth * source.Format.BitsPerPixel / 8
    '       Dim bits(stride * source.PixelHeight * source.Format.BitsPerPixel / 8) As Byte
    '       source.CopyPixels(bits, stride, 0)

    '       Dim g As GCHandle = GCHandle.Alloc(bits, GCHandleType.Pinned)
    '       Dim bitmap As New System.Drawing.Bitmap( _
    '     source.Width, source.Height, stride, Imaging.PixelFormat.Format32bppPArgb, g.AddrOfPinnedObject)
    '       g.Free()
    '       'If bitmap.Width > size.Width Or bitmap.Height > size.Height Then
    '       Dim b2 As New Drawing.Bitmap(size.Width, size.Height, Imaging.PixelFormat.Format32bppArgb)

    '       Dim r = MenulatorLib.ScaleRect(bitmap.Size, New RectangleF(0, 0, size.Width, size.Height), False)
    '       Dim gr As Graphics = Graphics.FromImage(b2)
    '       'gr.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
    '       'gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
    '       'gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
    '       MenulatorLib.SetUpGraphics(gr)
    '       gr.DrawImage(bitmap, r, New Rectangle(Point.Empty, bitmap.Size), GraphicsUnit.Pixel)
    '       gr.Dispose()
    '       ShellThumbnail = b2.Clone(New Rectangle(Point.Empty, b2.Size), Imaging.PixelFormat.Format32bppArgb)
    '       gr.Dispose()
    '       b2.Dispose()
    '       'ElseIf bitmap.Width > size.Width Then

    '       'ElseIf bitmap.Height > size.Height Then
    '       'Else
    '       '    ShellThumbnail = bitmap.Clone(New Rectangle(0, 0, bitmap.Width, bitmap.Height), PixelFormat.Format32bppArgb)
    '       'End If
    '       bitmap.Dispose()

    '       Marshal.Release(hbitmap)
    '       Marshal.ReleaseComObject(ppsi)
    '       source = Nothing

    '   End Function



    '    Private Structure BLENDFUNCTION
    '        Public BlendOp As Byte
    '        Public BlendFlags As Byte
    '        Public SourceConstantAlpha As Byte
    '        Public AlphaFormat As Byte
    '    End Structure
    '    <DllImport("gdi32")> Private Shared Function CreateBitmap( _
    '  ByVal nWidth As Integer, _
    '  ByVal nHeight As Integer, _
    '  ByVal cPlanes As UInteger, _
    '  ByVal cBitsPerPel As UInteger, _
    '  ByVal lpvBits As IntPtr _
    ') As IntPtr

    '    End Function
    '    <DllImport("gdi32")> Private Shared Function CreateCompatibleDC(ByVal hdc As IntPtr) As IntPtr

    '    End Function
    '    <DllImport("Msimg32")> Private Shared Function AlphaBlend( _
    '  ByVal hdcDest As IntPtr, _
    '  ByVal nXOriginDest As Integer, _
    '  ByVal nYOriginDest As Integer, _
    '  ByVal nWidthDest As Integer, _
    '  ByVal nHeightDest As Integer, _
    '  ByVal hdcSrc As IntPtr, _
    '  ByVal nXOriginSrc As Integer, _
    '  ByVal nYOriginSrc As Integer, _
    '  ByVal nWidthSrc As Integer, _
    '  ByVal nHeightSrc As Integer, _
    '  ByVal blendFunction As BLENDFUNCTION _
    '  ) As Boolean

    '    End Function
    '    <DllImport("gdi32")> Private Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr

    '    End Function
    '    <DllImport("gdi32")> Private Shared Function DeleteDC(ByVal hdc As IntPtr) As Boolean
    '    End Function
    '    Private Shared Function ConvertDDBToDIB(ByVal hbm As IntPtr) As IntPtr

    '        '// Is it a valid bitmap?
    '        Dim bm As New WIN32BITMAP
    '        If (GetObject(hbm, Marshal.SizeOf(bm), bm) <> Marshal.SizeOf(bm)) Then
    '            Return 0
    '        End If
    '        '// Is it already a DIB section?
    '        If (bm.bmBits) Then
    '            Return hbm
    '        End If

    '        '// Create the new DIB.
    '        Dim hdc As IntPtr = GetDC(0)
    '        Dim bi As New BITMAPINFO
    '        bi.bmiHeader.biSize = Marshal.SizeOf(GetType(BITMAPINFOHEADER))
    '        bi.bmiHeader.biWidth = bm.bmWidth
    '        bi.bmiHeader.biHeight = -bm.bmHeight '; // top-down
    '        bi.bmiHeader.biPlanes = 1
    '        bi.bmiHeader.biBitCount = bm.bmBitsPixel
    '        bi.bmiHeader.biCompression = 0

    '        If (bi.bmiHeader.biBitCount < 24) Then
    '            bi.bmiHeader.biBitCount = 24
    '        End If
    '        Dim bits As IntPtr = 0
    '        Dim dib As IntPtr = CreateDIBSection(hdc, bi, 0, bits, 0, 0)
    '        If (dib) Then
    '            GetDIBits(hdc, hbm, 0, Math.Abs(bi.bmiHeader.biHeight), bits, bi, 0)
    '        End If
    '        ReleaseDC(0, hdc)
    '        Return dib
    '    End Function
    '    Private Shared Function CreatePARGBFromHICON(ByVal hIcon As IntPtr, ByVal size As UInteger) As Bitmap


    '        '// Create GDI+ bitmap from icon, preserving alpha channel.
    '        Dim ii As New ICONINFO
    '        GetIconInfo(hIcon, ii)
    '        Dim dib As IntPtr = ConvertDDBToDIB(ii.hbmColor)
    '        Dim ds As New DIBSECTION
    '        GetObject(dib, Marshal.SizeOf(ds), ds)
    '        Dim sourceBitmap As Bitmap
    '        If True Then ' (HasAlpha(dib)) Then

    '            sourceBitmap = New Bitmap(ds.dsBm.bmWidth, ds.dsBm.bmWidth, _
    '            ds.dsBm.bmWidthBytes, Imaging.PixelFormat.Format32bppArgb, _
    '            ds.dsBm.bmBits)

    '        Else

    '            sourceBitmap = New Bitmap(hIcon)
    '        End If
    '        DeleteObject(ii.hbmColor)
    '        DeleteObject(ii.hbmMask)

    '        '// Rescale bitmap to target size.
    '        Dim scaledBitmap As New Bitmap(size, size, Imaging.PixelFormat.Format32bppArgb) ' PixelFormat32bppARGB)
    '        Dim g = Graphics.FromImage(scaledBitmap)
    '        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear ' (InterpolationModeHig  hQualityBilinear)
    '        g.DrawImage(sourceBitmap, 0, 0, scaledBitmap.Width, scaledBitmap.Height)

    '        '// Extract and return the GDI bitmap.
    '        Dim hbm As IntPtr = 0
    '        'hbm = scaledBitmap.GetHbitmap(Color.FromArgb(0, 0, 0))
    '        CreatePARGBFromHICON = scaledBitmap
    '        'sourceBitmap.Dispose()
    '        DeleteObject(dib)
    '        'Return hbm
    '    End Function

    '    <DllImport("gdi32.dll")> _
    '     Private Shared Function CreateDIBSection(ByVal hdc As Int32, _
    '       ByRef pbmi As BITMAPINFO, ByVal iUsage As System.UInt32, _
    '       <Out()> ByRef ppvBits As IntPtr, ByVal hSection As Int32, _
    '       ByVal dwOffset As System.UInt32) As Int32
    '    End Function
    '    <DllImport("gdi32")> Private Shared Function GetDIBits( _
    '    ByVal hdc As IntPtr, _
    '    ByVal hbmp As IntPtr, _
    '    ByVal uStartScan As UInteger, _
    '    ByVal cScanLines As UInteger, _
    '    ByRef lpvBits As IntPtr, _
    '    ByRef lpbi As BITMAPINFO, _
    '    ByVal uUsage As UInteger _
    '  ) As Integer
    '    End Function
    '    Private Structure BITMAPINFO
    '        Public bmiHeader As BITMAPINFOHEADER
    '        <MarshalAs(UnmanagedType.ByValArray, sizeconst:=1)> Public bmiColors() As Int32
    '    End Structure
    '    Private Structure BITMAPINFOHEADER
    '        Public biSize As Int32
    '        Public biWidth As Int32
    '        Public biHeight As Int32
    '        Public biPlanes As Int16
    '        Public biBitCount As Int16
    '        Public biCompression As Int32
    '        Public biSizeImage As Int32
    '        Public biXPelsPerMeter As Int32
    '        Public biYPelsPerMeter As Int32
    '        Public biClrUsed As Int32
    '        Public biClrImportant As Int32
    '        Public colors As Int32
    '    End Structure
    '    <DllImport("user32")> Private Shared Function GetIconInfo(ByVal hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Boolean

    '    End Function

    '    Private Structure DIBSECTION
    '        Public dsBm As WIN32BITMAP
    '        Public dsBmih As BITMAPINFOHEADER
    '        <MarshalAs(UnmanagedType.ByValArray, sizeconst:=3)> Public dsBitfields() As Int32
    '        Public dshSection As IntPtr
    '        Public dsOffset As Int32
    '    End Structure
    '    <DllImport("user32")> Private Shared Function ReleaseDC( _
    '  ByVal hWnd As IntPtr, _
    '  ByVal hDC As IntPtr _
    ') As Integer
    '    End Function

    '    Private Structure Pixel
    '        Public b, g, r, a As Byte
    '    End Structure

    'Public Shared Function ConvertWithAlphaBlend(ByVal ipd As IntPtr) As Bitmap

    '    '// get the info about the HBITMAP inside the IPictureDisp
    '    Dim dibsection As New DIBSECTION
    '    GetObjectDIBSection(ipd, Marshal.SizeOf(dibsection), dibsection)
    '    Dim width As Integer = dibsection.dsBm.bmWidth
    '    Dim height As Integer = dibsection.dsBm.bmHeight

    '    '// zero out the RGB values for all pixels with A == 0 
    '    '// (AlphaBlend expects them to all be zero)
    '    'unsafe()
    '    '    {


    '    'Dim pBits(dibsection.dsBm.bmBitsPixel) As RGBQUAD '= dibsection.dsBm.bmBits

    '    For x As Integer = 0 To dibsection.dsBmih.biWidth
    '        For y As Integer = 0 To dibsection.dsBmih.biHeight

    '            Dim offset As Integer = y * dibsection.dsBmih.biWidth + x
    '            Dim pBits As RGBQUAD = Marshal.PtrToStructure(Marshal.ReadIntPtr(dibsection.dsBm.bmBits, offset), GetType(RGBQUAD))
    '            If (pBits.rgbReserved = 0) Then

    '                pBits.rgbRed = 0
    '                pBits.rgbGreen = 0
    '                pBits.rgbBlue = 0
    '            End If
    '        Next
    '    Next
    '    '}

    '    '// create the destination Bitmap object
    '    Dim bitmap As New Bitmap(width, height, PixelFormat.Format32bppArgb)

    '    '// get the HDCs and select the HBITMAP
    '    Dim graphics = Drawing.Graphics.FromImage(bitmap)

    '    Dim hdcDest As IntPtr = graphics.GetHdc()
    '    Dim hdcSrc As IntPtr = CreateCompatibleDC(hdcDest)
    '    Dim hobjOriginal As IntPtr = SelectObject(hdcSrc, ipd)

    '    '// render the bitmap using AlphaBlend
    '    Dim blendfunction As New BLENDFUNCTION(AC_SRC_OVER, 0, &HFF, AC_SRC_ALPHA)
    '    AlphaBlend(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, width, height, blendfunction)

    '    '// clean up
    '    SelectObject(hdcSrc, hobjOriginal)
    '    DeleteDC(hdcSrc)
    '    graphics.ReleaseHdc(hdcDest)
    '    graphics.Dispose()

    '    Return bitmap
    'End Function

    'Public Shared Function ConvertPixelByPixel(ByVal ipd As IntPtr) As Bitmap

    '    '// get the info about the HBITMAP inside the IPictureDisp
    '    Dim dibsection As New DIBSECTION
    '    GetObjectDIBSection(ipd, Marshal.SizeOf(dibsection), dibsection)
    '    Dim width As Integer = dibsection.dsBm.bmWidth
    '    Dim height As Integer = dibsection.dsBm.bmHeight

    '    '// create the destination Bitmap object
    '    Dim bitmap As New Bitmap(width, height, PixelFormat.Format32bppArgb)

    '    'unsafe()
    '    '  {
    '    '// get a pointer to the raw bits
    '    'Dim pBits() As RGBQUAD = dibsection.dsBm.bmBits

    '    '// copy each pixel manually
    '    For x As Integer = 0 To dibsection.dsBmih.biWidth - 1
    '        For y As Integer = 0 To dibsection.dsBmih.biHeight - 1

    '            Dim offset As Integer = y * dibsection.dsBmih.biWidth + x
    '            Dim pBits As RGBQUAD = Marshal.PtrToStructure(Marshal.ReadIntPtr(dibsection.dsBm.bmBits, offset * Marshal.SizeOf(GetType(RGBQUAD))), GetType(RGBQUAD))

    '            If (pBits.rgbReserved <> 0) Then

    '                bitmap.SetPixel(x, y, Color.FromArgb(pBits.rgbReserved, pBits.rgbRed, pBits.rgbGreen, pBits.rgbBlue))

    '            End If
    '        Next
    '    Next


    '    Return bitmap
    'End Function

    '<DllImport("gdi32.dll", EntryPoint:="GdiAlphaBlend")> _
    'Public Shared Function AlphaBlend(ByVal hdcDest As IntPtr, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, _
    '   ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, _
    '   ByVal hdcSrc As IntPtr, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, _
    '   ByVal blendFunction As BLENDFUNCTION) As Boolean
    'End Function

    '<DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)> _
    'Public Shared Function CreateCompatibleDC(ByVal hdc As IntPtr) As IntPtr

    'End Function

    '<DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)> _
    'Public Shared Function DeleteDC(ByVal hdc As IntPtr) As Boolean

    'End Function

    '<DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)> _
    'Public Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr
    'End Function

    '<DllImport("gdi32.dll", ExactSpelling:=True, SetLastError:=True)> _
    'Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    'End Function

    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure BLENDFUNCTION

    '    Public BlendOp As Byte
    '    Public BlendFlags As Byte
    '    Public SourceConstantAlpha As Byte
    '    Public AlphaFormat As Byte

    '    Public Sub New(ByVal op As Byte, ByVal flags As Byte, ByVal alpha As Byte, ByVal format As Byte)

    '        BlendOp = op
    '        BlendFlags = flags
    '        SourceConstantAlpha = alpha
    '        AlphaFormat = format
    '    End Sub
    'End Structure

    'Public Const AC_SRC_OVER As Byte = &H0
    'Public Const AC_SRC_ALPHA As Byte = &H1

    '<DllImport("gdiplus.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)> _
    'Public Shared Function GdipCreateBitmapFromHBITMAP(ByVal hbitmap As IntPtr, _
    '    ByVal hpalette As IntPtr, <Out()> ByRef bitmap As IntPtr) As Integer
    'End Function

    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure RGBQUAD

    '    Public rgbBlue As Byte
    '    Public rgbGreen As Byte
    '    Public rgbRed As Byte
    '    Public rgbReserved As Byte
    'End Structure

    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure win32BITMAP

    '    Public bmType As Int32
    '    Public bmWidth As Int32
    '    Public bmHeight As Int32
    '    Public bmWidthBytes As Int32
    '    Public bmPlanes As Int16
    '    Public bmBitsPixel As Int16
    '    Public bmBits As IntPtr
    'End Structure

    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure BITMAPINFOHEADER

    '    Public biSize As Integer
    '    Public biWidth As Integer
    '    Public biHeight As Integer
    '    Public biPlanes As Int16
    '    Public biBitCount As Int16
    '    Public biCompression As Integer
    '    Public biSizeImage As Integer
    '    Public biXPelsPerMeter As Integer
    '    Public biYPelsPerMeter As Integer
    '    Public biClrUsed As Integer
    '    Public bitClrImportant As Integer
    'End Structure

    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure DIBSECTION

    '    Public dsBm As win32BITMAP
    '    Public dsBmih As BITMAPINFOHEADER
    '    Public dsBitField1 As Integer
    '    Public dsBitField2 As Integer
    '    Public dsBitField3 As Integer
    '    Public dshSection As IntPtr
    '    Public dsOffset As Integer
    'End Structure

    '<DllImport("gdi32.dll", EntryPoint:="GetObject")> _
    'Public Shared Function GetObjectDIBSection(ByVal hObject As IntPtr, ByVal nCount As Integer, ByRef lpObject As DIBSECTION) As Integer

    'End Function


End Class

'Public Class MenulatorLib




'    'Public Shared Sub BufferSound(ByVal strKey As String, ByVal stream As IO.Stream)
'    '    SoundBucket.Add(strKey, stream)
'    'End Sub
'    'Public Shared Sub PlayBufferedSound(ByVal strKey As String, ByVal ar As AsyncCallback)
'    '    SoundBucket.Play(strKey, ar)
'    'End Sub
'    'Public Shared Sub PlayBufferedSound(ByVal strKey As String)
'    '    SoundBucket.Play2(strKey)
'    'End Sub


'End Class

