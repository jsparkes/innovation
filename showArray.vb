'
' 
' showArray.vb This code is licenced under "Creative Commons Attribution Non Commercial 4.0 International"
' See: https://creativecommons.org/licenses/by-nc/4.0/legalcode
'
' This file implements showing card arrays for WinInnovation on the showArray winform

' Conditional compile directives

#Const VERBOSE = True ' FK adding debugging Frame work via c compiler like if defs should be DEBUG but not sure about interference

Option Strict Off
Option Explicit On

' Compiler/Build directives

#Disable Warning IDE1006 'Inherited code with variable names this suppresses: These words must begin with upper case characters
#Disable Warning IDE0054 'Inherited code with assignments (60) this suppresses: Use compound assignment x += 5 vs x = x + 5
#Disable Warning BC40000 'VB compatibility

Friend Class showArray

    Inherits System.Windows.Forms.Form

    Private Sub showArray_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Call initialize_images()
        Call Me.SetBounds(VB6.TwipsToPixelsX(12500), VB6.TwipsToPixelsY(4500), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
    End Sub

    Private Sub initialize_images()
        ' initialize the hand images
        Dim i As Short
        Static Dim initialized As Boolean = False

        If initialized Then
            Return
        End If
        For i = 0 To 40
            If i > imgShowIcon.UBound Then
                lblShowTitle.Load(i)
                imgShowIcon.Load(i)
                imgShowColor.Load(i)
            End If
            imgShowIcon(i).Visible = False
            imgShowColor(i).Visible = False
            lblShowTitle(i).Visible = False

            imgShowIcon(i).SetBounds(imgShowIcon(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgShowIcon(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            imgShowColor(i).SetBounds(imgShowColor(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgShowColor(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            lblShowTitle(i).SetBounds(lblShowTitle(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblShowTitle(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

            If i > 19 Then
                imgShowIcon(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgShowIcon(0).Left) + 2000), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgShowIcon(0).Top) + 360 * (i - 20)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                imgShowColor(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgShowColor(0).Left) + 2000), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgShowColor(0).Top) + 360 * (i - 20)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                lblShowTitle(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lblShowTitle(0).Left) + 2000), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblShowTitle(0).Top) + 360 * (i - 20)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            End If

            imgShowColor(i).SendToBack()
            imgShowColor(i).Tag = -1
            imgShowIcon(i).Tag = -1
            lblShowTitle(i).Tag = -1
        Next i

        initialized = True

    End Sub

    Public Sub load_pictures(ByVal player As Short, ByVal Index As Short, ByVal kind As String)
        Dim i, id As Object
        Dim max_size As Short

        initialize_images()

        If kind = "board" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            max_size = size3(board, player, Index) - 1
            Me.Tag = Index
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            max_size = size2(score_pile, player) - 1
            Me.Tag = 1000
#If VERBOSE Then
            Call Main_Renamed.append_simple("In load_pictures() Piles")
#End If
        End If

        For i = 0 To max_size
            If kind = "board" Then
                id = board(player, Index, i)
            Else
                id = score_pile(player, i)
            End If

            lblShowTitle(i).Text = age(id) & "-" & title(id)
            lblShowTitle(i).Tag = id
            lblShowTitle(i).Visible = True
            lblShowTitle(i).BackColor = Main_Renamed.background_colors(color_lookup(color(id)))

            Call Main_Renamed.set_icon_image(imgShowIcon(i), id, dogma_icon(id))
            Call Main_Renamed.set_color_image(imgShowColor(i), id, color(id))
        Next i

        Me.Controls.Clear()

        For i = 0 To max_size
            imgShowIcon(i).Visible = True
            'imgShowColor(i).Visible = True
            lblShowTitle(i).Visible = True
            Me.Controls.Add(imgShowIcon(i))
            Me.Controls.Add(imgShowColor(i))
            Me.Controls.Add(lblShowTitle(i))
        Next i
        For i = max_size + 1 To imgShowIcon.UBound
            imgShowIcon(i).Visible = False
            imgShowColor(i).Visible = False
            lblShowTitle(i).Visible = False
        Next i
        Me.Visible = True
    End Sub

    Private Sub lblShowTitle_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblShowTitle.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim Index As Short = lblShowTitle.GetIndex(eventSender)
        Call Main_Renamed.imgMouseMove(lblShowTitle(Index), Button, Shift, X, Y)
    End Sub

    Private Sub imgShowColor_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgShowColor.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim Index As Short = imgShowColor.GetIndex(eventSender)
        Call Main_Renamed.imgMouseMove(imgShowColor(Index), Button, Shift, X, Y)
    End Sub

    Private Sub imgShowIcon_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgShowIcon.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim Index As Short = imgShowIcon.GetIndex(eventSender)
        Call Main_Renamed.imgMouseMove(imgShowIcon(Index), Button, Shift, X, Y)
    End Sub
    '
    ' weird VB problem can't add cancel button or the form being full makes a control array that is invalid 
    ' Private Sub Button1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Button1.Click
    'Me.Dispose()
    'End Sub

    Private Sub lblShowTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblShowTitle.Click
        Dim Index As Short = lblShowTitle.GetIndex(eventSender)
        Call process_click(Index)
    End Sub
    Private Sub imgShowColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgShowColor.Click
        Dim Index As Short = imgShowColor.GetIndex(eventSender)
        Call process_click(Index)
    End Sub
    Private Sub imgShowIcon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgShowIcon.Click
        Dim Index As Short = imgShowIcon.GetIndex(eventSender)
        Call process_click(Index)
    End Sub


    Private Sub process_click(ByVal Index As Short)
        If phase = "publications" Then
            imgShowIcon(Index).Visible = False
            imgShowColor(Index).Visible = False
            lblShowTitle(Index).Visible = False
            human_data = human_data + 1
            'MsgBox "Comparing " & human_data & " to " & size3(board, 0, showArray.Tag) & " tag=" & showArray.Tag
            'MsgBox "setting " & size3(board, 0, showArray.Tag) - human_data & " to " & title(imgShowIcon(index).Tag)
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, 0, showArray.Tag). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            board(0, CInt(Me.Tag), size3(board, 0, CShort(Me.Tag)) - human_data) = CShort(imgShowIcon(Index).Tag)
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, 0, showArray.Tag). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If human_data = size3(board, 0, CShort(Me.Tag)) Then
                'UPGRADE_WARNING: Untranslated statement in process_click. Please check source code.
                Call Main_Renamed.update_display()
                Me.Close()
            End If
        ElseIf CDbl(Me.Tag) = 1000 Then
            Call Main_Renamed.process_score_pile_click(Index)
            Me.Close()
        End If
    End Sub
End Class
