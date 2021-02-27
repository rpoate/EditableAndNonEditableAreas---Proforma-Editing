Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace mshtml


    <ComImport>
    <Guid("3050F220-98B5-11CF-BB82-00AA00BDCE0B")>
    <InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
    Interface IHTMLTxtRange

        Function duplicate() As IHTMLTxtRange
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1010)>
        Function inRange(
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal range As IHTMLTxtRange) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1011)>
        Function isEqual(
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal range As IHTMLTxtRange) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1012)>
        Sub scrollIntoView(
    <[In]> ByVal Optional fStart As Boolean = True)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1013)>
        Sub collapse(
    <[In]> ByVal Optional Start As Boolean = True)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1014)>
        Function expand(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1015)>
        Function move(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String,
    <[In]> ByVal Optional Count As Integer = 1) As Integer
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1016)>
        Function moveStart(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String,
    <[In]> ByVal Optional Count As Integer = 1) As Integer
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1017)>
        Function moveEnd(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String,
    <[In]> ByVal Optional Count As Integer = 1) As Integer

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1024)>
        Sub [select]()

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1026)>
        Sub pasteHTML(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal html As String)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1006)>
        Function parentElement() As Object ' IHTMLElement
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1001)>
        Sub moveToElementText(
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal element As Object)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1025)>
        Sub setEndPoint(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal how As String,
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal SourceRange As IHTMLTxtRange)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1018)>
        Function compareEndPoints(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal how As String,
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal SourceRange As IHTMLTxtRange) As Integer
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1019)>
        Function findText(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal _String As String,
    <[In]> ByVal Optional Count As Integer = 1073741823,
    <[In]> ByVal Optional Flags As Integer = 0) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1020)>
        Sub moveToPoint(
    <[In]> ByVal x As Integer,
    <[In]> ByVal y As Integer)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1009)>
        Function moveToBookmark(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Bookmark As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1027)>
        Function queryCommandSupported(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1028)>
        Function queryCommandEnabled(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1029)>
        Function queryCommandState(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1030)>
        Function queryCommandIndeterm(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1031)>
        Function queryCommandText(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As String
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1032)>
        Function queryCommandValue(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Object
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1033)>
        Function execCommand(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String,
    <[In]> ByVal Optional showUI As Boolean = False,
    <[Optional]>
    <[In]>
    <MarshalAs(UnmanagedType.Struct)> ByVal Optional value As Object = Nothing) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1034)>
        Function execCommandShowHelp(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
    End Interface
End Namespace

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.HtmlEditControl1.DocumentHTML = "<h1 class='protected'>Non-Editable Heading</h1><div>Editable paragraph</div><p class='protected'>Non-editable Paragraph. </p><div>Editable paragraph <span class='protected'>containing a non-editable inline element. </span></div><p class='protected'>Non-editable Paragraph.</p>"
        Me.HtmlEditControl1.CSSText = "body {font-family: arial;} .protected {background-color: #F0F0F0; user-select: none;}"
        Me.HtmlEditControl1.HideDOMToolbar = True ' prevent element deletion from DOM toolbar And CSS Class name changes by user

    End Sub

    Private Sub HtmlEditControl1_CancellableUserInteraction(ByVal sender As Object, ByVal e As Zoople.CancellableUserInteractionEventsArgs) Handles HtmlEditControl1.CancellableUserInteraction
        If ((e.InteractionType = "onmousedown") _
                    OrElse (e.InteractionType = "onmouseup")) Then
            Return
        End If

        ' let mouse actions happen
        Console.WriteLine(e.InteractionType)
        Dim isAnyPartOfSelectionProtected As Boolean = IsElementProtected()

        If (e.Keys.Control _
                    AndAlso (e.Keys.Keycode = CType(Keys.A, Integer))) Then
            e.Cancel = True
            Return
        End If

        If ((e.Keys.Control _
                    AndAlso (e.Keys.Keycode = CType(Keys.C, Integer))) _
                    OrElse (e.Keys.Control _
                    AndAlso (e.Keys.Keycode = CType(Keys.Z, Integer)))) Then
            e.Cancel = False
            Return
        End If

        If ((e.Keys.Keycode = CType(Keys.Up, Integer)) _
                    OrElse ((e.Keys.Keycode = CType(Keys.Down, Integer)) _
                    OrElse ((e.Keys.Keycode = CType(Keys.Left, Integer)) _
                    OrElse (e.Keys.Keycode = CType(Keys.Right, Integer))))) Then
            e.Cancel = False
            Return
        End If

        If isAnyPartOfSelectionProtected Then
            If ((e.Keys.Keycode = CType(Keys.Delete, Integer)) _
                        OrElse (e.Keys.Keycode = CType(Keys.Back, Integer))) Then
                e.Cancel = True
                Return
            End If

            ' we have found a protected element so prevent this part of the document from being edited
            If ((e.InteractionType = "onkeydown") _
                        OrElse (e.InteractionType = "onkeypress")) Then
                e.Cancel = True
                Return
            End If

        End If

    End Sub

    Private Function IsElementProtected() As Boolean

        Dim retVal As Boolean = False
        Dim COMDoc As Object = HtmlEditControl1.Document.DomDocument

        If (COMDoc.selection.type = "Control") Then
            Return True
        End If

        ' this is for singularly selected images and table element selections etc
        Dim Range As mshtml.IHTMLTxtRange = COMDoc.selection.createrange()

        If (Range.parentElement.hasAttribute("class") AndAlso Range.parentElement.getAttribute("class") = "protected") Then
            Marshal.ReleaseComObject(Range)
            Return True
        End If

        Dim x As Integer = 0

        Do While (x < Range.parentElement.Children.length)

            If (Range.parentElement.Children(x).hasAttribute("class") AndAlso Range.parentElement.Children(x).getAttribute("class") = "protected") Then

                Dim testRange As mshtml.IHTMLTxtRange = COMDoc.selection.createrange()
                testRange.moveToElementText(Range.parentElement.Children(x))

                Dim startPoints = testRange.compareEndPoints("EndToStart", Range)
                Dim endPoints = testRange.compareEndPoints("StartToEnd", Range)

                If ((startPoints <= 0) OrElse (endPoints >= 0)) Then
                    retVal = False
                Else
                    Marshal.ReleaseComObject(testRange)
                    Marshal.ReleaseComObject(Range)
                    Return True
                End If

                Marshal.ReleaseComObject(testRange)

            End If

            x += 1
        Loop

        Marshal.ReleaseComObject(Range)
        Return retVal

    End Function

    Private Sub htmlEditControl1_CommandsToolbarButtonClicked(ByVal sender As Object, ByVal e As Zoople.CommandsToolbarButtonClickedEventArgs)
        If ((e.ButtonIdentifier = "tsbCut") _
                    AndAlso IsElementProtected()) Then
            MessageBox.Show("Selection contains a read only area and cannot be 'Cut'")
            e.Cancelled = True
        End If

    End Sub
End Class
