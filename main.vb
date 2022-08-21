'
' 
' main.vb This code is licenced under "Creative Commons Attribution Non Commercial 4.0 International"
' See: https://creativecommons.org/licenses/by-nc/4.0/legalcode
'
' This file is the main code for WinInnovation on the main.vb[Design] winform

' Conditional compile directives

#Const VERBOSE = True ' FK adding debugging Frame work via c compiler like if defs should be DEBUG but not sure about interference
#Const VERBOSE2 = False ' For too 2 much info tmi

Option Strict Off
Option Explicit On

Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Imports Microsoft.VisualBasic.Compatibility 'FK doesn't seem to do anything, but documenting anyway
Imports System.Collections.Generic

' Compiler/Build directives

#Disable Warning IDE1006 'Inherited code with variable names this suppresses: These words must begin with upper case characters
#Disable Warning IDE0054 'Inherited code with assignments (60) this suppresses: Use compound assignment x += 5 vs x = x + 5
#Disable Warning IDE0044 'Make field readonly warning messages are suppressed
#Disable Warning BC40000 'VB compatibility warning messages are suppressed

''''''''''''''''''''''''
'
' Main_Renamed - THe main code "container"
'
''''''''''''''''''''''''

Friend Class Main_Renamed
    Inherits System.Windows.Forms.Form
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    'Dim current_score, original_score As Object
    Dim current_score, original_score As Integer
    Dim original_ai_mode As Short
    Dim already_returned(40) As Short
    Dim found(110) As Object
    Dim alert_found As Short
    Dim splay_direction As String
    Dim splay_options(5) As Short
    Dim fake_deck(10, 20) As Object
    Dim fake_hand(4, 100) As Short
    Dim ai_cheat_level As Short

    ' FK Consts added for code readbility
    ' https://www.dotnetheaven.com/article/vb.net-const-and-read-only-member
    ' was using readonly but switched to const

    Const AGECOUNT As Short = 9 ' FK Ages are 1-10 IRL, code uses 0-9
    Const ICONCOUNT As Short = 6 ' FK Icons are 1-6 IRL Leaf. Castle, Lightbulb, Crown, Factory ,Clock, Code has an extra placeholder X
    Const COLORCOUNT As Short = 4 ' FK Colors are 5 Yellow, Red, Purple, Blue, Green
    Const MAXPLAYERS As Short = 3 ' FK 4 players max
    Const TOTACHIEVE As Short = 13 ' FK 0-8 + 5 special achievements = 13

    'Images
    Dim color_images(10) As System.Drawing.Image
    Dim icon_images(10) As System.Drawing.Image
    Dim achievement_images(15) As System.Drawing.Image
    Dim title_image As System.Drawing.Image

    'Background colors
    Public Dim background_colors = New Dictionary(Of String, Color) From {
        {"Yellow", System.Drawing.Color.FromArgb(227, 224, 145)},
        {"Red", System.Drawing.Color.FromArgb(216, 146, 144)},
        {"Purple", System.Drawing.Color.FromArgb(161, 117, 180)},
        {"Blue", System.Drawing.Color.FromArgb(132, 173, 217)},
        {"Green", System.Drawing.Color.FromArgb(129, 187, 126)}
    }

    ''''''''''''''''''''''''
    '
    ' Menu Handlers
    '
    ''''''''''''''''''''''''

    ' File-->Load menu
    Private Sub LoadGameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadGameToolStripMenuItem.Click

    End Sub

    ' File-->Save menu
    Private Sub SaveGameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveGameToolStripMenuItem.Click
        save_game()
    End Sub

    ' File-->RestartNewGame menu
    Private Sub RestartNewGameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestartNewGameToolStripMenuItem.Click
        ' fk nope Main_Load()
        Application.Restart()
    End Sub
    ' File-->exit menu
    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    ' Help-->Rules menu
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        'ShowHelp()
    End Sub
    ' Help-->Shortcuts menu
    Private Sub ShortCutsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShortCutsToolStripMenuItem.Click
        'ShowShortcuts()
    End Sub
    ' Help-->About menu
    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem1.Click
        Dim frmAbout As New AboutBox
        frmAbout.ShowDialog(Me)
        'ShowAbout()
    End Sub

    ''''''''''''''''''''''''
    '
    ' Main Form "handler"
    '
    ''''''''''''''''''''''''
    Private Sub Main_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
#If VERBOSE Then
        Call append_simple("In Main_Load() -----------------------------")
#End If
        alert_found = 0
        num_players = 4
        Call initialize_card_data()
        Call initialize_images()
        Call initialize_game()
        Call start_new_game()
        Call update_display()
    End Sub

    ''''''''''''''''''''''''
    '
    ' Main keybord keys "handler"
    ' Implemented keys:
    ' c - Continue
    ' n - new game
    ' r - restart game
    ' x - exit
    '
    ''''''''''''''''''''''''
    Private Sub Main_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
#If VERBOSE Then
        Call append_simple("In Main_KeyDown +++ ")
#End If
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.R Then
            Call load_game()
            Call ready_for_action()

            ' Clear some random buttons
            cmd2Players.Visible = False
            cmd3Players.Visible = False
            cmd4Players.Visible = False
            cmbCheatLevel.Visible = False
            lblCheatLevel.Visible = False
            lblOtherApps.Visible = False
            imgLarge.SendToBack()
            Line1.BringToFront()
        ElseIf KeyCode = System.Windows.Forms.Keys.N Then
            phase = "new_game"
            Call start_new_game()
        ElseIf KeyCode = System.Windows.Forms.Keys.X Then
            Me.Close()
        ElseIf KeyCode = System.Windows.Forms.Keys.C Then
            Call cmdNext_Click(cmdNext, New System.EventArgs())
        End If
    End Sub
    ''''''''''''''''''''''''
    '
    ' debugging handler - not functioning 
    '
    ''''''''''''''''''''''''
    Private Sub set_debug_values()
#If VERBOSE Then
        Call append_simple("In set_debug_values() ++++++")
#End If
        Dim card_name As String
        card_name = "Sanitation"
        test_card = -1

        Exit Sub
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        test_card = get_id_by_name(card_name)

        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(3, get_id_by_name("Societies"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(0, get_id_by_name("Evolution"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(0, get_id_by_name("Medicine"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(0, get_id_by_name("Medicine"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(0, get_id_by_name("Clothing"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(1, get_id_by_name("A.I."))
        'Call meld(0, get_id_by_name("Canning"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(2, get_id_by_name("Fermenting"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(1, get_id_by_name("Machine Tools"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(1, get_id_by_name("Masonry"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(0, get_id_by_name("Gunpowder"))
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call meld(0, get_id_by_name("Coal"))
        'Call meld(0, get_id_by_name("Software"))

        ' Meld the card we are testing
        'Call meld(0, test_card)
        Call meld(0, test_card)
        Call meld(1, test_card)
        hand(1, 0) = test_card
        vps(2) = 1
        'splayed(0, 1) = "Up"

        Call Push2(score_pile, 0, get_id_by_name("Machine Tools"))
        Call Push2(score_pile, 0, get_id_by_name("Machine Tools"))
        Call Push2(score_pile, 0, get_id_by_name("Machine Tools"))
        Call Push2(score_pile, 0, get_id_by_name("Machine Tools"))
        'Call Push2(score_pile, 0, get_id_by_name("Alchemy"))
        'Call Push2(score_pile, 0, get_id_by_name("Monotheism"))
        'Call Push2(hand, 0, get_id_by_name("Engineering"))
        'Call Push2(hand, 1, get_id_by_name("Engineering"))
        'Call Push2(hand, 1, get_id_by_name("Engineering"))
        'Call Push2(hand, 1, get_id_by_name("Engineering"))
        'Call Push2(hand, 1, get_id_by_name("Engineering"))
        Call Push2(score_pile, 1, get_id_by_name("Oars"))
        Call Push2(score_pile, 0, get_id_by_name("Alchemy"))
        Call Push2(score_pile, 1, get_id_by_name("Monotheism"))
        Call Push2(score_pile, 2, get_id_by_name("Philosophy"))
        Call Push2(score_pile, 2, get_id_by_name("Compass"))
        Call Push2(score_pile, 3, get_id_by_name("Philosophy"))
        Call Push2(score_pile, 3, get_id_by_name("Compass"))
        'Call Push2(hand, 0, get_id_by_name("A.I."))
        'Call Push2(hand, 0, get_id_by_name("A.I."))
        'Call Push2(hand, 0, get_id_by_name("A.I."))
        'Call Push2(hand, 0, get_id_by_name("A.I."))
        Call Push2(hand, 1, get_id_by_name("Oars"))
        Call Push2(hand, 1, get_id_by_name("Oars"))
        Call Push2(hand, 1, get_id_by_name("A.I."))

    End Sub

    Private Function get_id_by_name(ByVal str_Renamed As String) As Object
        Dim i As Short
        For i = 0 To num_cards - 1

            'MsgBox "comparing -" & title(i) & "- to -" & str & "-"
#If VERBOSE Then
            Call append(0, "In get_id_by_name() ------ comparing - " & title(i) & " - to - " & str_Renamed & " - ")
#End If
            If title(i) = str_Renamed Then
                'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                get_id_by_name = i
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_id_by_name = -1
    End Function

    Private Sub cmbCards_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCards.SelectedIndexChanged
        Dim i As Short
        Dim txt As String
        txt = cmbCards.SelectedItem.ToString()
        For i = 0 To num_cards - 1
            If txt = title(i) Then
                Call load_picture(i)
            End If
        Next i
    End Sub

    Public Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
#If VERBOSE Then
        Call append_simple("In cmdCancel_Click() -----------------------------")
#End If
        Dim i, player As Short
        player = 0
        If phase = "pottery" Then
            If human_data > 0 Then
                Call draw_and_score(player, human_data)
            End If
        ElseIf phase = "currency" Then
            For i = 0 To AGECOUNT
                If already_returned(i) = 1 Then Call draw_and_score(player, 2)
            Next i
        ElseIf phase = "democracy" Then
            If human_data >= democracy_minimum Then
                Call draw_and_score(player, 8)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
                democracy_minimum = human_data + 1
            End If
        ElseIf phase = "evolution" Then
            If score_pile(player, 0) = -1 Then
                Call draw_num(player, 1)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, age(score_pile(player, size2(score_pile, player) - 1)) + 1)
            End If
        ElseIf phase = "lighting" Then
            For i = 0 To 9
                If already_returned(i) = 1 Then Call draw_and_score(player, 7)
            Next i
        ElseIf phase = "antibiotics" Then
            For i = 0 To AGECOUNT
                If already_returned(i) = 1 Then
                    Call draw_num(player, 8)
                    Call draw_num(player, 8)
                End If
            Next i
        ElseIf phase = "collaboration" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(human_data, find_card(human_data, already_returned(0)))
            'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_in_hand(human_data, active_player, find_card(human_data, already_returned(1)))
            'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(active_player, find_card(active_player, already_returned(1)))
        ElseIf phase = "suburbia" Then
            For i = 1 To human_data
                Call draw_and_score(player, 1)
            Next i

        End If
        Call resume_dogma()
    End Sub

    Private Sub resume_dogma()

        cmdCancel.Visible = False
        cmdYes.Visible = False
        cmdYes.Text = "Yes"
        cmdOther.Visible = False
        'MsgBox "changing player from " & dogma_player & " to " & dogma_player + 1
#If VERBOSE Then
        Call append_simple("In resume_dogma() -----------------------------")
        Call append(0, "In resume_dogma() -----changing player from " & dogma_player & " to " & dogma_player + 1)
#End If
        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_player = dogma_player + 1
        phase = "playing game"

        ' If we are doing a normal dogma, proceed with that.
        ' Otherwise if we are solo-activating one, do that
        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If solo_dogma_id > -1 Then
            'alert "resuming solo"
            'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call perform_solo_dogma_effect(solo_dogma_player, solo_dogma_id)
        Else
            'alert "resuming normal"
            Call perform_dogma_effect_by_level(dogma_id)
        End If

    End Sub

    Private Sub cmdDogma_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDogma.Click
        Dim index As Short = cmdDogma.GetIndex(eventSender)
        Dim id As Short
        id = board(0, index, 0)
        Call perform_dogma_effect(id)
    End Sub

    Private Sub cmdNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNext.Click
        If cmdNext.Visible = True Then
            cmdNext.Visible = False
            Call play_game()
        End If
    End Sub

    Private Sub cmdOddball_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOddball.Click
#If VERBOSE Then
        Call append_simple("In cmdOddball_Click() -----------------------------")
#End If
        Dim index As Short = cmdOddball.GetIndex(eventSender)
        Dim player, id As Object
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        player = 0

        If phase = "empiricism" Then
            human_data = human_data - 1
            cmdOddball(index).Visible = False
            If human_data = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                id = draw_num(player, 9)
                If cmdOddball(color(id)).Visible = False Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call meld_from_hand(player, find_card(player, id))
                    For i = 0 To 4
                        cmdOddball(i).Visible = False
                    Next i
                    human_data = color(id)
                    Call show_prompt("empiricism2", "You may splay " & color_lookup(color(id)) & " Up.", 2)
                Else
                    For i = 0 To 4
                        cmdOddball(i).Visible = False
                    Next i
                    Call resume_dogma()
                End If
            End If
        ElseIf phase = "massMedia" Then
            Call return_from_all_score_piles(index + 1)
            For i = 0 To AGECOUNT
                cmdOddball(i).Visible = False
            Next i
            Call resume_dogma()
        End If

    End Sub

    Private Sub cmdScore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdScore.Click
        Call showArray.load_pictures(0, 0, "score")
    End Sub

    Private Sub cmdStack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStack.Click
#If VERBOSE Then
        Call append_simple("In cmdStack_Click() -----------------------------")
#End If
        Dim index As Short = cmdStack.GetIndex(eventSender)
        Call showArray.load_pictures(0, index, "board")
    End Sub

    Private Sub cmdYes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdYes.Click
#If VERBOSE Then
        Call append_simple("In cmdYes_Click() -----------------------------")
#End If
        'Dim player As Object
        Dim i, player As Short
        player = 0
        If phase = "codeOfLawsA" Then
            Call splay(0, human_data, "Left")
        ElseIf phase = "canalBuilding" Then
            Call exchange_canal_building(player)
        ElseIf phase = "engineering" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 1) And active_player <> player Then dogma_copied = 1
            Call splay(player, 1, "Left")
        ElseIf phase = "feudalism" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 0) And active_player <> player Then dogma_copied = 1
            Call splay(player, 0, "Left")
        ElseIf phase = "machinery1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 1) And active_player <> player Then dogma_copied = 1
            Call splay(player, 1, "Left")
        ElseIf phase = "paper" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 4) And active_player <> player Then dogma_copied = 1
            Call splay(player, 4, "Left")
        ElseIf phase = "enterprise" Or phase = "banking2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 4) And active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call splay(player, 4, "Right")
        ElseIf phase = "printingPress2" Or phase = "chemistry" Or phase = "atomicTheory" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 3) And active_player <> player Then dogma_copied = 1
            Call splay(player, 3, "Right")
        ElseIf phase = "reformation" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 0) And active_player <> player Then dogma_copied = 1
            Call splay(player, 0, "Right")
        ElseIf phase = "coal2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 1) And active_player <> player Then dogma_copied = 1
            Call splay(player, 1, "Right")
        ElseIf phase = "statistics2" Or phase = "canning2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 0) And active_player <> player Then dogma_copied = 1
            Call splay(player, 0, "Right")
        ElseIf phase = "canning" Then
            Call draw_and_tuck(player, 6)
            For i = 0 To 4
                If board(player, i, 0) > -1 Then
                    If Not card_has_symbol(board(player, i, 0), 5) Then
                        Call score_top_card(player, i)
                    End If
                End If
            Next i
        ElseIf phase = "emancipation2" Or phase = "industrialization2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 1) And active_player <> player Then dogma_copied = 1
            Call splay(player, 1, "Right")
        ElseIf phase = "metricSystem2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 4) And active_player <> player Then dogma_copied = 1
            Call splay(player, 4, "Right")
        ElseIf phase = "bicycle" Then
            Call exchange_bicycle(player)
        ElseIf phase = "evolution" Then
            Call draw_and_score(player, 8)
            lblPrompt.Text = "Choose a card in your score pile to return."
            phase = "evolution2"
            cmdCancel.Visible = False
            cmdYes.Visible = False
            Exit Sub
        ElseIf phase = "publications2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 0) And active_player <> player Then dogma_copied = 1
            Call splay(player, 0, "Up")
        ElseIf phase = "empiricism2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, human_data) And active_player <> player Then dogma_copied = 1
            Call splay(player, human_data, "Up")
        ElseIf phase = "flight2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 1) And active_player <> player Then dogma_copied = 1
            Call splay(player, 1, "Up")
        ElseIf phase = "massMedia2" Or phase = "satellites2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 2) And active_player <> player Then dogma_copied = 1
            Call splay(player, 2, "Up")
        ElseIf phase = "collaboration" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(human_data, find_card(human_data, already_returned(1)))
            'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_in_hand(human_data, active_player, find_card(human_data, already_returned(0)))
            'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(active_player, find_card(active_player, already_returned(0)))
        ElseIf phase = "computers" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 1) And active_player <> player Then dogma_copied = 1
            Call splay(player, 1, "Up")
        ElseIf phase = "specialization2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 0) And active_player <> player Then dogma_copied = 1
            Call splay(player, 0, "Up")
        ElseIf phase = "stemCells" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = size2(hand, player) - 1 To 0 Step -1
                Call score_from_hand(player, 0)
            Next i
        ElseIf phase = "theInternet" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 4) And active_player <> player Then dogma_copied = 1
            Call splay(player, 4, "Up")
        End If

        Call resume_dogma()
    End Sub

    Private Sub cmdOther_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOther.Click
#If VERBOSE Then
        Call append_simple("In cmdOther_Click() -----------------------------")
#End If
        Dim player As Short
        player = 0
        If phase = "feudalism" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 2) And active_player <> player Then dogma_copied = 1
            Call splay(player, 2, "Left")
        ElseIf phase = "paper" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 3) And active_player <> player Then dogma_copied = 1
            Call splay(player, 3, "Left")
        ElseIf phase = "reformation" Or phase = "emancipation2" Or phase = "industrialization2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 2) And active_player <> player Then dogma_copied = 1
            Call splay(player, 2, "Right")
        ElseIf phase = "publications2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 3) And active_player <> player Then dogma_copied = 1
            Call splay(player, 3, "Up")
        ElseIf phase = "computers" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 4) And active_player <> player Then dogma_copied = 1
            Call splay(player, 4, "Up")
        ElseIf phase = "specialization2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, 3) And active_player <> player Then dogma_copied = 1
            Call splay(player, 3, "Up")
        End If
        Call resume_dogma()
    End Sub



    Public Sub play_game()
#If VERBOSE Then
        Call append(0, "In play_game() -- before phase = " & phase & " ai_mode =  " & ai_mode & " break_dogma_loop = " & break_dogma_loop)
#End If
        Dim i, gss_count As Short
        'MsgBox "before" & phase & " " & ai_mode & " " & break_dogma_loop
        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If phase = "new_game" Or ai_mode = 1 Or break_dogma_loop = 1 Then Exit Sub
        'MsgBox "after"
        ' If there are no actions remaining, move on to the next player
        Call update_display()
        If actions_remaining = 0 Then
            'alert ("out of actions")
            For i = 0 To num_players - 1
                winners(i) = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object scored_this_turn(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                scored_this_turn(i) = 0
                tucked_this_turn(i) = 0
            Next i
            gss_count = 1
            'For i = 0 To 2
            '    lblPlayerDetail(i) = score_game_individual(i + 1)
            'Next i
            cmdNext.Visible = True
            cmdNext.Focus()
            current_turn = current_turn + 1
            active_player = (active_player + 1) Mod num_players
            actions_remaining = 2
            If current_turn = 2 And num_players = 4 Then actions_remaining = 1
#If VERBOSE Then
            Call append_simple("In play_game?() -----------------------------")
#Else
            Call append_simple("-----------------------------")
#End If


            If active_player = 0 Then Call save_game()
        End If

        If active_player = 0 Then
            Call ready_for_action()
        Else
            Call wait_for_next_action()
            Call perform_AI_action()
        End If
    End Sub

    Private Sub perform_AI_action()
        If actions_remaining = 1 Then
#If VERBOSE Then
            Call append_simple("In perform_AI_action() -----------------------------")
#Else
            Call append_simple("-----------------------------")
#End If
        End If
        Dim choices(9) As Object
        'Dim cheat As Object
        Dim cheat, num As Short
        ai_mode = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object gss_count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gss_count = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object gs_count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gs_count = 1
        'MsgBox "looking to pick a move"
#If VERBOSE Then
        Call append_simple("perform_AI_action() looking to pick a move ----")
#End If
        ' Fake data depending on cheat level
        cheat = 1
        num = Int(100 * Rnd())
        If num >= (30 + 14 * ai_cheat_level) Then
            cheat = 0
            'MsgBox "not cheating " & num & " vs " & 30 + 14 * ai_cheat_level
#If VERBOSE Then
            Call append(0, "perform_AI_action() Not cheating " & num & " vs " & 30 + 14 * ai_cheat_level)
#End If
            Call fake_data(active_player)
        Else
            'MsgBox "cheating " & num & " vs " & 30 + 14 * ai_cheat_level
#If VERBOSE Then
            Call append(0, "In perform_AI_action() cheating " & num & " vs " & 30 + 14 * ai_cheat_level)
#End If
        End If

        Call pick_ai_move(active_player, choices, 0)
        ai_mode = 0
        'If choices(0) > 0 Then MsgBox "choice: " & choices(0)
        'MsgBox "picked " & choices(0) & " and " & choices(1) & " ai_mode = " & ai_mode
#If VERBOSE Then
        Call append(0, "In perform_AI_action() picked " & choices(0) & " and " & choices(1) & " ai_mode = " & ai_mode)
#End If
        ' Unfake data
        If cheat = 0 Then
            Call unfake_data(active_player)
        End If

        Call make_ai_move(active_player, choices, 0)
        'MsgBox "Made move " & choices(0) & " ai_mode = " & ai_mode
#If VERBOSE Then
        Call append(0, "In perform_AI_action() Made move " & choices(0) & " ai_mode = " & ai_mode)
#End If
        ' Automatically continue the game if a non-dogma action was taken, or
        ' if the human player is next
        'If (actions_remaining > 0 And choices(0) <> 2) Or (actions_remaining = 0 And active_player = num_players - 1) Then Call play_game
        'Call wait_for_next_action
    End Sub

    Private Sub ready_for_action()
#If VERBOSE Then
        Call append_simple("In ready_for_action() ++++++")
#End If
        Dim i As Short

        phase = "waiting for action"

        lblActionsRemaining.Text = "Actions Remaining: " & actions_remaining
        lblActionsRemaining.Visible = True

        lblPrompt.Visible = True
        lblPrompt.Text = "Please choose an action.  To Achieve, click on the unclaimed achievement you would like to claim."

        cmdDraw.Visible = True
        'UPGRADE_WARNING: Couldn't resolve default property of object find_next_draw(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        cmdDraw.Text = "Draw a " & find_next_draw(0)
        cmdNext.Visible = False

        For i = 0 To COLORCOUNT
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, 0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size3(board, 0, i) > 0 Then cmdDogma(i).Visible = True
        Next i
    End Sub

    Private Sub disable_actions()
        Dim i As Short
        lblPrompt.Visible = True
        lblActionsRemaining.Visible = False
        cmdDraw.Visible = False
        For i = 0 To 4
            cmdDogma(i).Visible = False
        Next i
    End Sub


    Private Sub cmd2Players_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd2Players.Click
        Call launch_game(2)
    End Sub
    Private Sub cmd3Players_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd3Players.Click
        Call launch_game(3)
    End Sub
    Private Sub cmd4Players_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd4Players.Click
        Call launch_game(4)
    End Sub

    Private Sub cmdDraw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDraw.Click
        draw((0))
        'MsgBox "pre actions remaining: " & actions_remaining
#If VERBOSE Then
        Call append(0, "In cmdDraw_Click() ++++ pre actions remaining: " & actions_remaining)
#End If
        actions_remaining = actions_remaining - 1
        Call play_game()
    End Sub


    Private Sub launch_game(ByRef players As Object)
        ' FK below is an extra initialize game?? nope removing first one
        ' Call initialize_game()
#If VERBOSE Then
        Call append_simple("In launch_game() ++++++")
#End If
        ' Clear some random buttons
        ai_cheat_level = cmbCheatLevel.SelectedIndex
        cmd2Players.Visible = False
        cmd3Players.Visible = False
        cmd4Players.Visible = False
        lblOtherApps.Visible = False
        cmbCheatLevel.Visible = False
        lblCheatLevel.Visible = False
        imgLarge.SendToBack()
        Line1.BringToFront()

        'UPGRADE_WARNING: Couldn't resolve default property of object players. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        num_players = players

        'Dim i As Object FK
        Dim i, j As Short

        ' Hide all information for players not playing the game
        For i = 1 To num_players - 1
            lblPlayer(i).Visible = True
            lblPlayerDetail(i - 1).Visible = True
            For j = 0 To 5
                lblIcon(i + j * 4).Visible = True
            Next j
            lblVP(i).Visible = True
            lblScore(i).Visible = True
        Next i
        For i = num_players To 3
            lblPlayer(i).Visible = False
            lblPlayerDetail(i - 1).Visible = False
            For j = 0 To 5
                lblIcon(i + j * 4).Visible = False
            Next j
            lblVP(i).Visible = False
            lblScore(i).Visible = False
        Next i

        ' Remove the top card from each deck 1-9 IRL these are cards are used for scoring
        For i = 0 To 8
            deck(i, size2(deck, i) - 1) = -1
        Next i

        ' Give every player 2 cards
        For i = 0 To num_players - 1
            Call draw_num(i, 1)
            Call draw_num(i, 1)
        Next i

        ' Read the debug file and update the board state accordingly
        'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim filename As Object
        Dim str_Renamed = ""
#If VERBOSE Then
        Call append_simple("In launch_game() ++++++  Read the debug file and update the board state accordingly /testing.txt")
#End If
        Dim id, pos As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        filename = My.Application.Info.DirectoryPath & "/testing.txt"
        'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(filename) <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(1, filename, OpenMode.Input)
#If VERBOSE Then
            Call append_simple("In BROKEN launch_game() found /testing.txt ++++++")
#End If
            While Not EOF(1)
                Input(1, str_Renamed)
                If Mid(str_Renamed, 1, 1) <> "'" Then
                    pos = InStr(str_Renamed, ":")
                    If pos > 0 Then
                        'MsgBox Mid(str, 1, pos - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object get_id_by_name(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        id = get_id_by_name(Mid(str_Renamed, 1, pos - 1))

                        pos = InStr(str_Renamed, "hand")
                        If pos > 0 Then
                            i = Val(Mid(str_Renamed, pos + 4, 1)) - 1
                            Call Push2(hand, i, id)
                        End If

                        pos = InStr(str_Renamed, "score")
                        If pos > 0 Then
                            i = Val(Mid(str_Renamed, pos + 5, 1)) - 1
                            Call Push2(score_pile, i, id)
                        End If

                        pos = InStr(str_Renamed, "topcard")
                        If pos > 0 Then
                            i = Val(Mid(str_Renamed, pos + 7, 1)) - 1
                            Call Unshift3(board, i, color(id), id)
                        End If

                        pos = InStr(str_Renamed, "tuck")
                        If pos > 0 Then
                            i = Val(Mid(str_Renamed, pos + 4, 1)) - 1
                            Call Push3(board, i, color(id), id)
                        End If
                    End If
                End If
            End While
            FileClose(1)
        Else
            Call append_simple(" ") ' FK not sure if this required don't want blank endif ifdef isn't there?
#If VERBOSE Then
            Call append_simple("In launch_game() /testing.txt ++++++ NOT FOUND")
#End If

        End If

        Call update_icon_total()

        phase = "initial_meld"
        lblPrompt.Text = "Choose a card to meld by clicking on a card in your hand.  The person melding the lowest card alphabetically will play first."

        Call load_picture(hand(0, 0))

        Call set_debug_values()
#If VERBOSE Then
        Call append_simple("In launch_game() set_debug_values() DISABLED ")
        ' this call below was being used all the time but not much was being done
        ' Call set_debug_values()
#End If
        ' Update the display

        Call update_display()
    End Sub

    Private Sub start_new_game()
        ' Choose the number of players
#If VERBOSE Then
        Call append_simple("In start_new_game() ++++++ Choose the number of players")
#End If
        lblPrompt.Text = "Welcome to WinInnovation AI.  How many AI players would you like in this game?"
        lblActionsRemaining.Visible = False
        cmd2Players.Visible = True
        cmd3Players.Visible = True
        cmd4Players.Visible = True
        lblOtherApps.Visible = True

        'FK stoping from displaying doesn't seem to do anything leaving code in for now
        'cmbCheatLevel.Visible = True
        'lblCheatLevel.Visible = True
        cmbCheatLevel.Visible = False
        lblCheatLevel.Visible = False
        cmdNext.Visible = False
        cmdDraw.Visible = False
        imgLarge.Image = title_image
        imgLarge.BringToFront()
        phase = "new_game"
        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        break_dogma_loop = 1
    End Sub

    Private Sub initialize_game()
#If VERBOSE Then
        Call append_simple("In initialize_game() ++++++")
#End If
        Dim i, j As Short

        ' Set all scores to 0
        For i = 0 To MAXPLAYERS
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            scores(i) = 0
            vps(i) = 0
        Next i

        ' Empty hands, score piles, and boards
        For i = 0 To MAXPLAYERS
            hand(i, 0) = -1
            score_pile(i, 0) = -1
            For j = 0 To 4
                board(i, j, 0) = -1
                splayed(i, j) = ""
            Next j
        Next i

        ' Create new decks and shuffle them
        ' FK Why 11 decks? is he using 0 as an index??
        For i = 0 To 10
            deck(i, 0) = -1
        Next i
        'For i = 0 To num_cards - 1
        For i = num_cards - 1 To 0 Step -1
            Call Push2(deck, age(i) - 1, i)
        Next i
        For i = 0 To AGECOUNT
            Call randomize_array2(deck, i)
        Next i

        ' FK not sure what line is doing below but it "empties" textbox1 ????
        ' txtLog.Text = ""

        ' Unclaim all achievements
        For i = 0 To TOTACHIEVE
            achievements_scored(i) = 0
            lblAchievement(i).Visible = True
        Next i

        Call update_display()

    End Sub

    Private Sub fake_data(ByVal player As Short)
#If VERBOSE Then
        Call append_simple("In fake_data() ++++++")
#End If
        ' FK converted
        ' Dim j, i, r As Object
        Dim j, i, r, yes As Short

        ' Fake the decks
        For i = 0 To AGECOUNT
            Call copy_array2(deck, fake_deck, i)
            Call randomize_array2(deck, i)
        Next i

        ' Fake player hands
        For i = 0 To num_players - 1
            If i <> player Then
                Call copy_array2(hand, fake_hand, i)
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = 0 To size2(hand, i) - 1
                    yes = 1
                    While (yes)
                        r = Int(num_cards * Rnd())
                        If age(r) = age(hand(i, j)) Then
                            hand(i, j) = r
                            yes = 0
                        End If
                    End While
                Next j
            End If
        Next i
    End Sub

    Private Sub unfake_data(ByVal player As Short)
#If VERBOSE Then
        Call append_simple("In unfake_data() -----------------------------")
#End If
        Dim i As Short

        ' UnFake the decks
        For i = 0 To AGECOUNT
            Call copy_array2(fake_deck, deck, i)
        Next i

        ' UnFake player hands
        For i = 0 To num_players - 1
            If i <> player Then
                Call copy_array2(fake_hand, hand, i)
            End If
        Next i
    End Sub

    Private Sub perform_solo_dogma_effects(ByVal player As Short, ByVal id As Short)
        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        solo_dogma_id = id
        solo_dogma_level = -1
        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        solo_dogma_player = player
        Call perform_solo_dogma_effect(player, id)
    End Sub

    Private Sub perform_solo_dogma_effect(ByVal player As Short, ByVal id As Short)
#If VERBOSE Then
        Call append_simple("In perform_solo_dogma_effect() -----------------------------")
#End If
        ' Loop through each player that is affected by this dogma.
        Dim quit, solo_dogma_id As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        break_dogma_loop = 0
        'solo_dogma_id = id
        'alert ("going solo for player " & player + 1)

        ' Increment the level
        solo_dogma_level = solo_dogma_level + 1
        'alert "Trying " & title(id) & " for player " & player + 1 & " id = " & solo_dogma_id

        ' If the solo-dogma effect is complete, then resume the normal dogma
        quit = 0
        If solo_dogma_level > 2 Then
            quit = 1
        ElseIf Len(dogma(id, solo_dogma_level)) < 2 Then
            quit = 1
        End If

        If quit = 1 Then
            solo_dogma_id = -1
            'alert "quitting"
            Call resume_dogma()
            Exit Sub
        End If
        'alert "not quitting"

        ' If there is another dogma effect, call that now.
        If is_demand(id, solo_dogma_level) = 0 Then
            If solo_dogma_level = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call append(solo_dogma_player, "activates the first effect of " & title(id))
            ElseIf solo_dogma_level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call append(solo_dogma_player, "activates the second effect of " & title(id))
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call append(solo_dogma_player, "activates the third effect of " & title(id))
            End If
            'alert ("Beginning " & title(id))
            Call activate_card(player, id, solo_dogma_level + 1)
            'alert ("Finishing " & title(id))
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If break_dogma_loop = 1 Then Exit Sub

        ' Automatically move on to the next level
        Call perform_solo_dogma_effect(player, id)
    End Sub


    Public Sub perform_dogma_effect(ByVal id As Short)
        ' Dim i, j As Object
        Dim i, j, symbol As Short

        dogma_id = id
        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_level = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_player = (active_player + 1) Mod num_players
        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_copied = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        demand_met = 0
        democracy_minimum = 1

        'alert ("setting dogma player to " & dogma_player)
        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        solo_dogma_id = -1
        solo_dogma_level = -1

        ' Determing which players will be affected by this dogma
        symbol = dogma_icon(id)
        For i = 0 To num_players - 1
            For j = 0 To 2
                'UPGRADE_WARNING: Couldn't resolve default property of object affected_by_dogma(i, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                affected_by_dogma(i, j) = 0
                If is_demand(id, j) = 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object affected_by_dogma(i, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If icon_total(i, symbol) < icon_total(active_player, symbol) Then affected_by_dogma(i, j) = 1
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object affected_by_dogma(i, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If icon_total(i, symbol) >= icon_total(active_player, symbol) Then affected_by_dogma(i, j) = 1
                End If
            Next j
        Next i
        If title(id) = "Canal Building" Then
            For i = 0 To num_players - 1
                'alert ("affected " & i + 1 & ":" & affected_by_dogma(i, 0))
#If VERBOSE Then

                Call append(0, "In perform_dogma_effect() Canal Building ----affected " & i + 1 & ":" & affected_by_dogma(i, 0))
#End If
            Next i
        End If

        Call append(active_player, "activates the first effect of " & title(id))
        perform_dogma_effect_by_level((id))
    End Sub

    Private Sub perform_dogma_effect_by_level(ByVal id As Short)
        ' Loop through each player that is affected by this dogma.
        ' Dim player, i, loop_end As Object
        Dim i, loop_end, player, original_player As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        break_dogma_loop = 0

        'MsgBox "Dogma player: " & dogma_player & " Activating level " & dogma_level
#If VERBOSE Then
        Call append_simple("In perform_dogma_effect_by_level() -----------------------------")
        Call append(0, "In perform_dogma_effect_by_level() ----Dogma player: " & dogma_player & " Activating level " & dogma_level)
#End If
        loop_end = active_player + num_players
        If active_player = num_players - 1 Then loop_end = active_player
        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = dogma_player To loop_end
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            player = i Mod num_players
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object affected_by_dogma(player, dogma_level - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If affected_by_dogma(player, dogma_level - 1) = 1 Then
                'alert (title(id) & " is activating for player " & player + 1 & " with dogma_player = " & dogma_player & " with level " & dogma_level & " value = " & affected_by_dogma(player, dogma_level - 1))
                original_player = active_player
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call activate_card(player, id, dogma_level)
                'alert (title(id) & " is done")
                ' If we reach a human interaction, break the loop.  We will return later.
                'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If break_dogma_loop = 1 Or original_player <> active_player Then
                    'alert "loop broken"
                    Exit Sub
                End If
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dogma_player = i + 1
        Next i

        ' If there is another dogma effect, call that now.

        'MsgBox "Reached end of loop.  Next value = " & dogma(id, dogma_level)
#If VERBOSE Then
        Call append_simple("In perform_dogma_effect_by_level() ---If there is another dogma effect, call that now")
        Call append(player, "In perform_dogma_effect_by_level() ---Reached end of loop.  Next value = " & dogma(id, dogma_level))
#End If
        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If dogma_level < 3 And Len(dogma(id, dogma_level)) > 1 Then
            'alert (dogma(id, dogma_level))
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dogma_level = dogma_level + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If dogma_level = 2 Then
                Call append(active_player, "activates the second effect of " & title(id))
            Else
                Call append(active_player, "activates the third effect of " & title(id))
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dogma_player = (active_player + 1) Mod num_players
            'alert ("another effect... setting dogma player to " & dogma_player)
            Call perform_dogma_effect_by_level(id)
        Else
            ' Draw a card if our dogma was shared at all
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If dogma_copied = 1 Then
                Call append_simple("Dogma effect was shared.")
                draw((active_player))
            End If

            ' Move on to the next action
            If active_player = 0 Then
                actions_remaining = actions_remaining - 1
                Call play_game()
                'ElseIf actions_remaining > 0 Then
                '    Call play_game
            Else
                Call wait_for_next_action()
            End If
        End If
    End Sub

    Private Sub wait_for_next_action()
        Call disable_actions()
        cmdNext.Visible = True
        cmdNext.Focus()
        '& ", Actions Remaining: " & actions_remaining
        lblPrompt.Text = "Current Player: " & active_player + 1 & vbNewLine & "Click 'Continue' or hit 'C' to advance to the next player's turn."
    End Sub

    Private Sub achieve_special(ByVal player As Short, ByVal index As Short, ByVal str_Renamed As String)
        If achievements_scored(index) = 0 Then
            achievements_scored(index) = 1
            vps(player) = vps(player) + 1
            alert(("Player " & player + 1 & " claimed the '" & str_Renamed & "' achievement."))
            Call append(player, "claimed the '" & str_Renamed & "' achievement")
            Call update_display()
            Call end_game_points()
        End If
    End Sub

    Private Function get_highest_card_in_hand(ByVal player As Short) As Object
        ' Dim i, max As Object
        Dim i, max_index As Short
        Dim max As Integer
        max = -1
        max_index = -1
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            If age(hand(player, i)) > max Then
                max = age(hand(player, i))
                max_index = i
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_highest_card_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_highest_card_in_hand = max_index
    End Function

    Private Function get_highest_card_in_score_pile(ByVal player As Short) As Object
        'Dim i, max As Object
        Dim i, max_index As Short
        Dim max As Integer
        max = -1
        max_index = -1
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(score_pile, player) - 1
            If age(score_pile(player, i)) > max Then
                max = age(score_pile(player, i))
                max_index = i
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_highest_card_in_score_pile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_highest_card_in_score_pile = max_index
    End Function

    Private Function is_highest_card_in_score_pile(ByVal player As Short, ByVal index As Short) As Object
        'Dim i As Object
        Dim i, value As Short
        value = age(score_pile(player, index))
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(score_pile, player) - 1
            If age(score_pile(player, i)) > value Then
                'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_score_pile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                is_highest_card_in_score_pile = TriState.False
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_score_pile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        is_highest_card_in_score_pile = TriState.True
    End Function

    Private Function is_highest_card_in_hand(ByVal player As Short, ByVal index As Short) As Object
        'Dim i As Object
        Dim i, value As Short
        value = age(hand(player, index))
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            If age(hand(player, i)) > value Then
                'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                is_highest_card_in_hand = TriState.False
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        is_highest_card_in_hand = TriState.True
    End Function

    Private Function get_lowest_card_in_hand(ByVal player As Short) As Object
        ' Dim i, max As Object
        Dim i, max_index As Short
        Dim max As Integer
        max = 11
        max_index = -1
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            If age(hand(player, i)) < max Then
                max = age(hand(player, i))
                max_index = i
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_lowest_card_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_lowest_card_in_hand = max_index
    End Function

    Private Function is_lowest_card_in_hand(ByVal player As Short, ByVal index As Short) As Object
        Dim i As Object
        Dim value As Short
        value = age(hand(player, index))
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(hand(player, i)) < value Then
                'UPGRADE_WARNING: Couldn't resolve default property of object is_lowest_card_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                is_lowest_card_in_hand = TriState.False
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object is_lowest_card_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        is_lowest_card_in_hand = TriState.True
    End Function

    Private Sub return_from_all_score_piles(ByVal num As Short)
        Dim i As Object
        Dim j As Short
        For i = 0 To num_players - 1
            'MsgBox "trying " & i
            'MsgBox size2(score_pile, i)
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For j = size2(score_pile, i) - 1 To 0 Step -1
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If age(score_pile(i, j)) = num Then Call return_from_score_pile(i, j)
            Next j
        Next i
    End Sub

    Private Sub return_from_hand(ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = hand(player, index)
        Call remove_card_from_hand(player, index)
        Call Push2(deck, age(card) - 1, card)
        Call append(player, "returned a " & age(card))
    End Sub

    Private Sub return_from_board(ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = board(player, index, 0)
        Call remove_card_from_board(player, index)
        Call Push2(deck, age(card) - 1, card)
        Call append(player, "returned " & title(card))
        Call update_icon_total()
    End Sub

    Private Sub return_from_score_pile(ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = score_pile(player, index)
        Call remove_card_from_score_pile(player, index)
        Call Push2(deck, age(card) - 1, card)
        Call append(player, "returned a " & age(card))
    End Sub

    Private Sub tuck(ByVal player As Short, ByVal id As Short)
        Call Push3(board, player, color(id), id)
        Call update_icon_total()
        Call update_display()
        tucked_this_turn(player) = tucked_this_turn(player) + 1
        If tucked_this_turn(player) > 5 Then
            Call achieve_special(player, 9, "Monument")
        End If
    End Sub

    Private Sub tuck_from_hand(ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = hand(player, index)
        Call remove_card_from_hand(player, index)
        Call tuck(player, card)
        Call append(player, "tucks a " & color_lookup(color(card)) & " card from hand")
    End Sub

    Private Function can_splay(ByVal player As Short, ByVal index As Short) As Object
        'MsgBox index & " size: " & size3(board, player, index)
        'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If size3(board, player, index) > 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object can_splay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            can_splay = TriState.True
            Exit Function
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object can_splay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        can_splay = TriState.False
    End Function


    Private Function splay(ByVal player As Short, ByVal index As Short, ByVal direction As String) As Object
        ' alert ("trying to splay player " & player & " pile " & Index & " " & direction)
        'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If can_splay(player, index) Then
            splayed(player, index) = direction
            'UPGRADE_WARNING: Couldn't resolve default property of object splay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            splay = TriState.True
            Call check_for_achievements()
            Call update_icon_total()
            Call append(player, "splays " & color_lookup(index) & " " & direction)
            Call update_display()
            'alert ("splaying")
            Exit Function
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object splay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        splay = TriState.False
    End Function

    Private Function perform_best_splay(ByVal player As Short, ByRef splay_options As Object, ByVal direction As String) As Object
        'Dim max_score, max_index As Object
        Dim max_score, max_index As Integer
        Dim found, i As Short

        If player = 0 And ai_mode = 0 Then
            found = 0
            For i = 0 To 4
                'UPGRADE_WARNING: Couldn't resolve default property of object splay_options(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If splay_options(i) = 1 And can_splay(player, i) And splayed(player, i) <> direction Then found = 1
            Next i
            If found = 1 Then
                splay_direction = direction
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call show_prompt("splay", dogma(dogma_id, dogma_level - 1) & vbNewLine & "(click a top card to splay its pile)", 1)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object perform_best_splay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            perform_best_splay = TriState.False
            Exit Function
        End If


        'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        max_score = score_game_individual(player)
        max_index = -1
        For i = 0 To 4
            'UPGRADE_WARNING: Couldn't resolve default property of object splay_options(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If splay_options(i) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If can_splay(player, i) Then
                    Call save(player)
                    Call splay(player, i, direction)
                    Call restore(player)
                    If current_score > max_score Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object current_score. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object max_score. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        max_score = current_score
                        'UPGRADE_WARNING: Couldn't resolve default property of object max_index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        max_index = i
                    End If
                End If
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object max_index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If max_index > -1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object max_index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call splay(player, max_index, direction)
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object max_index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object perform_best_splay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        perform_best_splay = max_index
    End Function

    Private Sub transfer_card_in_hand(ByVal player_from As Short, ByVal player_to As Short, ByVal index As Short)
        Dim card As Short
        card = hand(player_from, index)
        Call remove_card_from_hand(player_from, index)
        Call Push2(hand, player_to, card)
        Call append(player_from, "transfers a " & age(card) & " from his hand to player " & player_to + 1 & "'s hand")
        Call update_display()
    End Sub

    Private Sub transfer_hand_to_score(ByVal player_from As Short, ByVal player_to As Short, ByVal index As Short)
        Dim card As Short
        card = hand(player_from, index)
        Call remove_card_from_hand(player_from, index)
        Call Push2(score_pile, player_to, card)
        Call append(player_from, "transfers a " & age(card) & " from his hand to player " & player_to + 1 & "'s score pile")
        Call update_display()
    End Sub

    Private Sub transfer_card_from_score(ByVal player_from As Short, ByVal player_to As Short, ByVal index As Short)
        Dim card As Short
        card = score_pile(player_from, index)
        'MsgBox "Index = " & index
        'MsgBox title(card) & index
        Call remove_card_from_score_pile(player_from, index)
        'alert title(score_pile(1, 0))
        Call Push2(score_pile, player_to, card)
        Call append(player_from, "transfers a " & age(card) & " from his score pile to player " & player_to + 1 & "'s pile")
        'Call append(player_from, "transfers a " & title(card) & " from his score pile to player " & player_to + 1 & "'s pile")
        Call update_display()

    End Sub

    Private Sub transfer_card_on_board(ByVal player_from As Short, ByVal player_to As Short, ByVal index As Short)
        Dim card As Short
        card = board(player_from, index, 0)
        If card = -1 Then Exit Sub
        Call shift3(board, player_from, index)
        Call Unshift3(board, player_to, index, card)
        Call append(player_from, "transfers " & title(card) & " from his board to player " & player_to + 1 & "'s board")
        Call check_for_achievements()
        Call update_icon_total()
        Call update_display()
    End Sub

    Private Sub transfer_board_to_score(ByVal player_from As Short, ByVal player_to As Short, ByVal index As Short)
        Dim card As Short
        card = board(player_from, index, 0)
        Call shift3(board, player_from, index)
        Call Push2(score_pile, player_to, card)
        Call append(player_from, "transfers " & title(card) & " from his board to player " & player_to + 1 & "'s score pile")
        Call check_for_achievements()
        Call update_icon_total()
        Call update_display()
    End Sub

    Private Sub transfer_board_to_hand(ByVal player_from As Short, ByVal player_to As Short, ByVal index As Short)
        Dim card As Short
        'MsgBox "looking for " & player_from + 1 & " index=" & Index
        card = board(player_from, index, 0)
        Call shift3(board, player_from, index)
        Call Push2(hand, player_to, card)
        Call append(player_from, "transfers " & title(card) & " from his board to player " & player_to + 1 & "'s hand")
        Call check_for_achievements()
        Call update_icon_total()
        Call update_display()
    End Sub

    Private Sub transfer_score_to_hand(ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = score_pile(player, index)
        Call remove_card_from_score_pile(player, index)
        Call Push2(hand, player, card)
        Call append(player, "transfers a " & age(card) & " from his score pile to hand")
        Call check_for_achievements()
        Call update_icon_total()
        Call update_display()
    End Sub

    Private Sub transfer_score_to_other_hand(ByVal player_from As Short, ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = score_pile(player_from, index)
        Call remove_card_from_score_pile(player_from, index)
        Call Push2(hand, player, card)
        Call append(player_from, "transfers a " & age(card) & " from his score pile to " & player + 1 & " 's hand")
        Call check_for_achievements()
        Call update_icon_total()
        Call update_display()
    End Sub

    Private Sub score_top_card(ByVal player As Short, ByVal index As Short)
        Dim card As Short
        card = board(player, index, 0)
        Call shift3(board, player, index)
        Call score(player, card)
        Call check_for_achievements()
        Call update_icon_total()
        Call update_display()
    End Sub


    Public Function card_has_symbol(ByVal id As Short, ByVal num As Short) As Object
        Dim i As Short
        If id = -1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            card_has_symbol = TriState.False
            Exit Function
        End If

        For i = 0 To 4
            If icons(id, i) = num Then
                'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                card_has_symbol = TriState.True
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        card_has_symbol = TriState.False
    End Function

    Private Sub exchange_bicycle(ByVal player As Short)
        Dim temp(40) As Object
        Dim count As Object
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        count = 0

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If active_player <> player Then dogma_copied = 1

        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = size2(hand, player) - 1 To 0 Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object temp(count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp(count) = hand(player, i)
            'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            count = count + 1
            hand(player, i) = -1
        Next i

        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = size2(score_pile, player) - 1 To 0 Step -1
            Call Push2(hand, player, score_pile(player, i))
            score_pile(player, i) = -1
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To count - 1
            Call Push2(score_pile, player, temp(i))
        Next i
        Call append(player, "chose to swapped")
        Call update_scores()
        Call update_display()
    End Sub

    Private Sub exchange_canal_building(ByVal player As Short)
        Dim temp(40) As Object
        Dim count As Object
        Dim high_in_hand = -1
        Dim high_in_score, i As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If active_player <> player Then dogma_copied = 1

        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        count = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If size2(hand, player) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object high_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            high_in_hand = age(hand(player, size2(hand, player) - 1))
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If size2(score_pile, player) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            high_in_score = age(score_pile(player, size2(score_pile, player) - 1))
        End If

        'alert ("Hand: " & high_in_hand & " and score = " & high_in_score)
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = size2(hand, player) - 1 To 0 Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object high_in_hand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If high_in_hand = age(hand(player, i)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temp(count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temp(count) = hand(player, i)
                'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                count = count + 1
                hand(player, i) = -1
            End If
        Next i
        'alert (count & " cards found in hand for player " & player + 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = size2(score_pile, player) - 1 To 0 Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, i)) = high_in_score Then
                Call Push2(hand, player, score_pile(player, i))
                score_pile(player, i) = -1
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To count - 1
            Call Push2(score_pile, player, temp(i))
        Next i
        Call append(player, "chose to swap")
        Call update_scores()
        Call update_display()
    End Sub

    Private Sub exchange_machinery(ByVal player As Short)
        Dim temp(40) As Object
        Dim count As Object
        Dim high_in_hand, i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        count = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, active_player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If size2(hand, active_player) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            high_in_hand = age(hand(active_player, size2(hand, active_player) - 1))
        End If
        'alert ("Hand: " & high_in_hand & " and score = " & high_in_score)
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, active_player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = size2(hand, active_player) - 1 To 0 Step -1
            If high_in_hand = age(hand(active_player, i)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temp(count). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temp(count) = hand(active_player, i)
                'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                count = count + 1
                hand(active_player, i) = -1
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = size2(hand, player) - 1 To 0 Step -1
            Call Push2(hand, active_player, hand(player, i))
            hand(player, i) = -1
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To count - 1
            Call Push2(hand, player, temp(i))
        Next i
        Call append(player, "got swapped")
        Call update_display()
    End Sub

    Private Sub collect_all_cards(ByVal player As Short, ByVal my_color As Short)
        ' Dim i As Object
        Dim i, j As Short
        For i = 0 To num_players - 1
            If i <> player Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = size2(hand, i) - 1 To 0 Step -1
                    If color(hand(i, j)) = my_color Then
                        Call transfer_card_in_hand(i, player, j)
                    End If
                Next j
            End If
        Next i
    End Sub

    Private Sub save(ByVal player As Short)
        'If ai_mode = 1 Then MsgBox "Lost in a save"
        'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object original_score. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        original_score = score_game_individual(player)
        original_ai_mode = ai_mode
        ai_mode = 1
        Call copy_game_state(0)
    End Sub

    Private Sub restore(ByVal player As Short)
        'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object current_score. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        current_score = score_game_individual(player)
        Call restore_game_state(0)
        ai_mode = original_ai_mode
    End Sub

    ''''''''''''''''''''''''
    '
    ' activate_card - This is where most of the card logic is defined
    '
    ''''''''''''''''''''''''

    Private Sub activate_card(ByVal player As Short, ByVal id As Short, ByVal level As Short)
#If VERBOSE Then
        Call append(player, "In activate_card() ACTIVATING = " & title(id))
#End If
        Dim card_name As String
        'Dim min, max, card_to_meld, index, card, found, max_index, min_index As Object
        Dim min, card_to_meld, index, card, found, min_index As Short
        Dim max_count As Short
        'Dim total, valid, max_score, count As Object
        Dim total, valid, count As Short
        Dim draw_id, max_index2, i, j As Short
        Dim max, max_index, max_score As Integer

        card_name = title(id)
        human_data = 0
        For i = 0 To COLORCOUNT '4 
            splay_options(i) = 0
        Next i
        For i = 0 To 10
            already_returned(i) = 0
        Next i

        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Age 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
        Call append(player, "In activate_card() AGE 1")
#End If
        Dim colors_found(3) As Short
        Dim age1 As Object
        Dim age2 As Short
        If card_name = "Agriculture" Then
#If VERBOSE Then
            Call append(player, "In activate_card() Agriculture")
#End If
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("agriculture", dogma(id, level - 1), 1)
                Else
                    ' Return your highest card
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    index = size2(hand, player) - 1
                    card = hand(player, index)
                    If card = -1 Then MsgBox("bad card found")
                    Call return_from_hand(player, index)
                    Call draw_and_score(player, age(card) + 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                End If
            End If
        ElseIf card_name = "Archery" Then
#If VERBOSE Then
            Call append(player, "In activate_card() Archery")
#End If
            Call draw_num(player, 1)
            If player = 0 And ai_mode <> 1 Then
                Call show_prompt("archery", dogma(id, level - 1), 0)
            Else
                ' Return your highest card
                'UPGRADE_WARNING: Couldn't resolve default property of object get_highest_card_in_hand(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                index = get_highest_card_in_hand(player)
                If index > -1 Then
                    Call transfer_card_in_hand(player, active_player, index)
                End If
            End If
        ElseIf card_name = "City States" Then
            If icon_total(player, 2) > 3 Then
                ' Find the castle with the lowest value
                min = 11
                min_index = -1
                For i = 0 To 4
                    'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If card_has_symbol(board(player, i, 0), 2) Then
                        If age(board(player, i, 0)) < min Then
                            min = age(board(player, i, 0))
                            min_index = i
                        End If
                    End If
                Next i

                If min_index > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("cityStates", dogma(id, level - 1), 0)
                    Else
                        Call transfer_card_on_board(player, active_player, min_index)
                        Call draw_num(player, 1)
                    End If
                End If
                demand_met = 1
            End If

        ElseIf card_name = "Clothing" Then
            If level = 1 Then
                ' Meld a player's highest card
                'UPGRADE_WARNING: Couldn't resolve default property of object card_to_meld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                card_to_meld = -1
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = size2(hand, player) - 1 To 0 Step -1
                    If board(player, color(hand(player, i)), 0) = -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_to_meld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        card_to_meld = i
                        i = 0
                    End If
                Next i
                'UPGRADE_WARNING: Couldn't resolve default property of object card_to_meld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If card_to_meld > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("clothing1", dogma(id, level - 1), 0)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_to_meld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(player, card_to_meld)
                    End If
                End If
            Else
                For j = 0 To 4
                    found = 0
                    For i = 0 To num_players - 1
                        If i <> player And board(i, j, 0) > -1 Then
                            found = 1
                        ElseIf i = player And board(i, j, 0) = -1 Then
                            found = 1
                        End If
                    Next i
                    If found = 0 Then
                        Call draw_and_score(player, 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                    End If
                Next j
            End If

        ElseIf card_name = "Code of Laws" Then
            If hand(player, 0) = -1 Then Exit Sub
            If player = 0 And ai_mode <> 1 Then
                Call show_prompt("codeOfLaws", dogma(id, level - 1), 1)
            Else
                'Exit Sub
                ' Try tucking each card in hand and see if any are worth splaying
                'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                original_score = score_game_individual(player)
                max_score = original_score - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = 0 To size2(hand, player) - 1
                    valid = 0
                    For i = 0 To 4
                        If board(player, i, 0) > -1 Then
                            If color(board(player, i, 0)) = color(hand(player, j)) Then
                                valid = 1
                            End If
                        End If
                    Next i
                    If valid = 1 Then
                        Call save(player)
                        draw_id = hand(player, j)
                        Call tuck_from_hand(player, j)
                        Call splay(player, color(draw_id), "Left")
                        Call restore(player)
                        If current_score > max_score Then
                            max_score = current_score
                            max_index = j
                        End If
                    End If
                Next j
                If max_score > original_score Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    draw_id = hand(player, max_index)
                    Call tuck_from_hand(player, max_index)
                    Call splay(player, color(draw_id), "Left")
                End If
            End If

        ElseIf card_name = "Domestication" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                If player = 0 And ai_mode = 0 Then
                    Call show_prompt("domestication", dogma(id, level - 1), 0)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object get_lowest_card_in_hand(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    id = get_lowest_card_in_hand(player)
                    Call meld_from_hand(player, id)
                    Call draw_num(player, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                End If
            Else
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
            End If

        ElseIf card_name = "Masonry" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                If player = 0 And ai_mode = 0 Then
                    Call show_prompt("masonry", dogma(id, level - 1), 1)
                Else
                    total = 0
                    count = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = 0 To size2(hand, player) - 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, i), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(hand(player, i), 2) Then total = total + 1
                    Next i
                    'alert ("total = " & total & " for player " & player + 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, i), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(hand(player, i), 2) Then
                            If total < 4 Then
                                Call save(player)
                                Call meld_from_hand(player, index)
                                Call restore(player)
                            End If
                            If total >= 4 Or current_score > original_score Then
                                Call meld_from_hand(player, i)
                                count = count + 1
                                If count = 4 Then Call achieve_special(player, 9, "Monument")
                                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If active_player <> player Then dogma_copied = 1
                            End If
                        End If
                    Next i
                End If
            End If

        ElseIf card_name = "Metalworking" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            count = 0
            While count = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                silent = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                id = draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                silent = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(id, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If card_has_symbol(id, 2) Then
                    Call append(player, "draws and scores " & title(id))
                    'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    silent = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call remove_card_from_hand(player, find_card(player, id))
                    Call score(player, id)
                    'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    silent = 0
                Else
                    Call append(player, "draws " & title(id))
                    count = 1
                End If
            End While

        ElseIf card_name = "Mysticism" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            id = draw_num(player, 1)
            If board(player, color(id), 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call meld_from_hand(player, find_card(player, id))
                Call draw_num(player, 1)
            End If

        ElseIf card_name = "Oars" Then
            If level = 2 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If demand_met = 0 Then
                    Call draw_num(player, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                End If
                Exit Sub
            End If

            'UPGRADE_WARNING: Couldn't resolve default property of object found. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            found = 0
            min = 99
            min_index = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 0 To size2(hand, player) - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, i), 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If card_has_symbol(hand(player, i), 4) Then
                    found = 1
                    If age(hand(player, i)) < min Then
                        min = age(hand(player, i))
                        min_index = i
                    End If
                End If
            Next i
            If found = 1 Then
                If player = 0 And ai_mode = 0 Then
                    Call show_prompt("oars", dogma(id, level - 1), 0)
                Else
                    id = hand(player, min_index)
                    Call append(active_player, "steals " & title(id) & " from player " & player + 1)
                    Call remove_card_from_hand(player, min_index)
                    Call Push2(score_pile, active_player, id)
                    Call draw_num(player, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    demand_met = 1
                End If
            End If

        ElseIf card_name = "Pottery" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("pottery", dogma(id, level - 1), 1)
                    Else
                        ' Return worst 3 cards
                        count = 0
                        For i = 0 To 2
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If size2(hand, player) > 0 Then
                                Call return_from_hand(player, 0)
                                count = count + 1
                            End If
                        Next i
                        Call draw_and_score(player, count)
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                    End If

                End If
            Else
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
            End If

        ElseIf card_name = "Sailing" Then
            Call draw_and_meld(player, 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1

        ElseIf card_name = "The Wheel" Then
            Call draw_num(player, 1)
            Call draw_num(player, 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1

        ElseIf card_name = "Tools" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 2 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("tools1", dogma(id, level - 1), 1)
                    Else
                        ' Return 3 cards worst
                        Call save(player)
                        For i = 0 To 2
                            Call return_from_hand(player, 0)
                        Next i
                        Call draw_and_meld(player, 3)
                        Call restore(player)
                        If current_score > original_score Then
                            For i = 0 To 2
                                Call return_from_hand(player, 0)
                            Next i
                            Call draw_and_meld(player, 3)
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If active_player <> player Then dogma_copied = 1
                        End If
                    End If

                End If
            Else
                ' Return a 3 to draw 3 1's
                found = -1
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = 0 To size2(hand, player) - 1
                    If age(hand(player, i)) = 3 Then found = i
                Next i
                If found > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("tools2", dogma(id, level - 1), 1)
                    Else
                        Call save(player)
                        Call return_from_hand(player, found)
                        Call draw_num(player, 1)
                        Call draw_num(player, 1)
                        Call draw_num(player, 1)
                        Call restore(player)
                        If current_score > original_score Then
                            Call return_from_hand(player, found)
                            Call draw_num(player, 1)
                            Call draw_num(player, 1)
                            Call draw_num(player, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If active_player <> player Then dogma_copied = 1
                        End If
                    End If
                End If
            End If

        ElseIf card_name = "Writing" Then
            Call draw_num(player, 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 2
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 2")
#End If
        ElseIf card_name = "Calendar" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(score_pile, player) > size2(hand, player) Then
                Call draw_num(player, 3)
                Call draw_num(player, 3)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
            End If

        ElseIf card_name = "Canal Building" Then
            If player = 0 And ai_mode <> 1 Then
                Call show_prompt("canalBuilding", dogma(id, level - 1), 2)
            Else
                Call save(player)
                Call exchange_canal_building(player)
                Call restore(player)
                If current_score > original_score Then
                    Call exchange_canal_building(player)
                End If
            End If

        ElseIf card_name = "Currency" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                For i = 0 To 9
                    already_returned(i) = 0
                Next i
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("currency", dogma(id, level - 1), 1)
                Else
                    ' Return a card from each age
                    count = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        If already_returned(age(hand(player, i))) = 0 Then
                            already_returned(age(hand(player, i))) = 1
                            Call return_from_hand(player, i)
                            count = count + 1
                        End If
                    Next i
                    For i = 0 To count - 1
                        Call draw_and_score(player, 2)
                    Next i
                End If
            End If

        ElseIf card_name = "Construction" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("construction", dogma(id, level - 1), 0)
                    Else
                        Call transfer_card_in_hand(player, active_player, 0)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If size2(hand, player) > 0 Then Call transfer_card_in_hand(player, active_player, 0)
                    End If
                End If
            Else
                found = 1
                For i = 0 To 4
                    If board(player, i, 0) = -1 Then found = 0
                Next i
                If found = 1 Then
                    For j = 0 To num_players - 1
                        If player <> j Then
                            found = 1
                            For i = 0 To 4
                                If board(j, i, 0) = -1 Then found = 0
                            Next i
                            If found = 1 Then Exit Sub
                        End If
                    Next j
                    Call achieve_special(player, 10, "Empire")
                End If
            End If

        ElseIf card_name = "Fermenting" Then
            If icon_total(player, 1) > 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                For i = 1 To Int(icon_total(player, 1) / 2)
                    Call draw_num(player, 2)
                Next i
            End If

        ElseIf card_name = "Mapmaking" Then
            If level = 1 Then
                If score_pile(player, 0) > -1 Then
                    If age(score_pile(player, 0)) = 1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        demand_met = 1
                        If player = 0 And ai_mode <> 1 Then
                            Call show_prompt("mapmaking", dogma(id, level - 1), 0)
                        Else
                            Call transfer_card_from_score(player, active_player, 0)
                        End If
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If demand_met = 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    Call draw_and_score(player, 1)
                End If
            End If

        ElseIf card_name = "Mathematics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                If player = 0 And ai_mode <> 1 Then 'BUG FK? Mathematics second dogma not happening?
                    Call show_prompt("mathematics", dogma(id, level - 1), 1)
                Else
                    Call save(player)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    id = hand(player, size2(hand, player) - 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call return_from_hand(player, size2(hand, player) - 1)
                    Call draw_and_meld(player, age(id) + 1)
                    Call restore(player)
                    If current_score > original_score Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call return_from_hand(player, size2(hand, player) - 1)
                        Call draw_and_meld(player, age(id) + 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                    End If
                End If
            End If

        ElseIf card_name = "Monotheism" Then
            If level = 1 Then
                found = 0
                For i = 0 To 4
                    already_returned(i) = 0
                    min = 99
                    If board(player, i, 0) > -1 And board(active_player, i, 0) = -1 Then
                        found = 1
                        If age(board(player, i, 0)) < min Then
                            min_index = i
                            min = board(player, i, 0)
                        End If
                        already_returned(i) = 1
                    End If
                Next i
                If found > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("monotheism", dogma(id, level - 1), 0)
                    Else
                        Call transfer_board_to_score(player, active_player, min_index)
                        Call draw_and_tuck(player, 1)
                    End If
                End If
            Else
                Call draw_and_tuck(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
            End If

        ElseIf card_name = "Philosophy" Then
            If level = 1 Then
                found = 0
                For i = 0 To 4
                    splay_options(i) = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If can_splay(player, i) Then
                        found = 1
                        splay_options(i) = 1
                    End If
                Next i
                If found = 1 Then
                    Call perform_best_splay(player, splay_options, "Left")
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("philosophy2", dogma(id, level - 1), 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call score_from_hand(player, size2(hand, player) - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                    End If
                End If
            End If

        ElseIf card_name = "Road Building" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("roadBuilding1", dogma(id, level - 1), 1)
                Else
                    ' Meld the 1st card
                    max_index = -1
                    'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    max = score_game_individual(player)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = 0 To size2(hand, player) - 1
                        Call save(player)
                        Call meld_from_hand(player, i)
                        Call restore(player)
                        If current_score > max_score Then
                            max_score = current_score
                            max_index = i
                        End If
                    Next i
                    If max_index > -1 Then
                        Call meld_from_hand(player, max_index)
                        max_index = -1
                        max_index2 = -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If size2(hand, player) > 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            For i = size2(hand, player) - 1 To 0 Step -1
                                For j = 0 To num_players - 1
                                    Call save(player)
                                    Call meld_from_hand(player, i)
                                    If j <> player Then
                                        Call transfer_card_on_board(player, j, 1)
                                        Call transfer_card_on_board(j, player, 4)
                                    End If
                                    Call restore(player)
                                    If current_score > max_score Then
                                        max_score = current_score
                                        max_index = i
                                        max_index2 = j
                                    End If
                                Next j
                            Next i
                            If max_index > -1 Then
                                Call meld_from_hand(player, max_index)
                                If max_index2 > -1 And max_index2 <> player Then
                                    Call transfer_card_on_board(player, max_index2, 1)
                                    Call transfer_card_on_board(max_index2, player, 4)
                                End If
                            End If

                        End If
                    End If
                End If
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 3
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 3")
#End If
        ElseIf card_name = "Alchemy" Then
            If level = 1 Then
                If icon_total(player, 2) > 2 Then
                    found = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    For i = 1 To Int(icon_total(player, 2) / 3)
                        'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        draw_id = draw_num(player, 4)
                        'UPGRADE_WARNING: Couldn't resolve default property of object found. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If color_lookup(color(draw_id)) = "Red" Then found = 1
                    Next i
                    If found = 1 Then
                        Call append(player, "revealed a Red card, and must return their entire hand")
                        If player = 0 And ai_mode <> 1 Then
                            Call show_prompt("alchemy1", "Return all the cards from your hand", 0)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            For i = size2(hand, player) - 1 To 0 Step -1
                                Call return_from_hand(player, 0)
                            Next i
                        End If
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("alchemy2", dogma(id, level - 1), 0)
                    Else
                        max_score = 0
                        max_index = -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = size2(hand, player) - 1 To 0 Step -1
                            Call save(player)
                            Call meld_from_hand(player, i)
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If size2(hand, player) > 0 Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Call score_from_hand(player, size2(hand, player) - 1)
                            End If
                            Call restore(player)
                            If current_score > max_score Then
                                max_score = current_score
                                max_index = i
                            End If
                        Next i
                        Call meld_from_hand(player, max_index)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If size2(hand, player) > 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Call score_from_hand(player, size2(hand, player) - 1)
                        End If
                    End If
                End If
            End If

        ElseIf card_name = "Compass" Then
            found = 0
            min = 11
            min_index = -1
            For i = 0 To 4
                If board(player, i, 0) > -1 And color_lookup(i) <> "Green" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If card_has_symbol(board(player, i, 0), 1) Then
                        found = 1
                        If age(board(player, i, 0)) < min Then
                            min = age(board(player, i, 0))
                            min_index = i
                        End If
                    End If
                End If
            Next i
            If found = 1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("compass", dogma(id, level - 1), 0)
                Else
                    Call transfer_card_on_board(player, active_player, min_index)
                    'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    max_score = score_game_individual(player)
                    max_index = -1
                    For i = 0 To 4
                        If board(active_player, i, 0) > -1 Then
                            If Not card_has_symbol(board(active_player, i, 0), 1) Then
                                Call save(player)
                                Call transfer_card_on_board(active_player, player, i)
                                Call restore(player)
                                If current_score > max_score Then
                                    max_score = current_score
                                    max_index = i
                                End If
                                Call restore(player)
                            End If
                        End If
                    Next i
                    If max_index > -1 Then Call transfer_card_on_board(active_player, player, max_index)
                End If
            End If


        ElseIf card_name = "Education" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(score_pile, player) > 0 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("education", dogma(id, level - 1), 1)
                Else
                    Call save(player)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call return_from_score_pile(player, size2(score_pile, player) - 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If size2(score_pile, player) > 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object get_highest_card_in_score_pile(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call draw_num(player, age(score_pile(player, get_highest_card_in_score_pile(player))) + 2)
                    Else
                        Call draw_num(player, 2)
                    End If
                    Call restore(player)
                    If current_score > original_score Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call return_from_score_pile(player, size2(score_pile, player) - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If player <> active_player Then dogma_copied = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If size2(score_pile, player) > 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object get_highest_card_in_score_pile(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Call draw_num(player, age(score_pile(player, get_highest_card_in_score_pile(player))) + 2)
                        Else
                            Call draw_num(player, 2)
                        End If
                    End If
                End If
            End If

        ElseIf card_name = "Engineering" Then
            If level = 1 Then
                For i = 0 To 4
                    If board(player, i, 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(board(player, i, 0), 2) Then
                            Call transfer_board_to_score(player, active_player, i)
                        End If
                    End If
                Next i
            Else
                splay_options(1) = 1
                Call perform_best_splay(player, splay_options, "Left")
            End If

        ElseIf card_name = "Feudalism" Then
            If level = 1 Then
                found = 0
                min_index = -1
                min = 99
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = 0 To size2(hand, player) - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, i), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If card_has_symbol(hand(player, i), 2) Then
                        found = 1
                        If age(hand(player, i)) < min Then
                            min_index = i
                            min = age(hand(player, i))
                        End If
                    End If
                Next i
                If found = 1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("feudalism", dogma(id, level - 1), 0)
                    Else
                        Call transfer_card_in_hand(player, active_player, min_index)
                    End If
                End If
            Else
                splay_options(0) = 1
                splay_options(2) = 1
                Call perform_best_splay(player, splay_options, "Left")
            End If

        ElseIf card_name = "Machinery" Then
            If level = 1 Then
                Call exchange_machinery(player)
            Else
                found = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = 0 To size2(hand, player) - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, i), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If card_has_symbol(hand(player, i), 2) Then
                        found = 1
                        If age(hand(player, i)) < min Then
                            min_index = i
                            min = age(hand(player, i))
                        End If
                    End If
                Next i
                If found = 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If player <> active_player Then dogma_copied = 1
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("machinery", dogma(id, level - 1), 0)
                        Exit Sub
                    Else
                        Call score_from_hand(player, min_index)
                    End If
                End If
                splay_options(1) = 1
                Call perform_best_splay(player, splay_options, "Left")
            End If

        ElseIf card_name = "Medicine" Then
            If score_pile(player, 0) = -1 Then
                If score_pile(active_player, 0) > -1 Then
                    If active_player = 0 And ai_mode = 0 Then
                        human_data = player
                        Call show_prompt("medicine1", dogma(id, level - 1) & vbNewLine & "(Choose your lowest card.)", 0)
                    Else
                        Call transfer_card_from_score(active_player, player, 0)
                    End If
                End If
            Else
                human_data = -1
                If score_pile(active_player, 0) > -1 Then human_data = score_pile(active_player, 0)
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("medicine", dogma(id, level - 1) & vbNewLine & "(Choose your highest card.)", 0)
                ElseIf active_player = 0 And ai_mode = 0 And human_data > -1 Then
                    human_data = player
                    Call show_prompt("medicine1", dogma(id, level - 1) & vbNewLine & "(Choose your lowest card.)", 0)
                Else
                    If human_data = -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_from_score(player, active_player, size2(score_pile, player) - 1)
                    Else
                        Call remove_card_from_score_pile(active_player, 0)
                        'MsgBox "step 1, size=" & size2(score_pile, player) & " for player" & player + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_from_score(player, active_player, size2(score_pile, player) - 1)
                        Call Push2(score_pile, player, human_data)
                        Call append(active_player, " transfers a " & age(human_data) & " from his score pile to player " & player + 1 & "'s score pile")
                    End If
                End If
            End If

        ElseIf card_name = "Optics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            draw_id = draw_and_meld(player, 3)
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(draw_id, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(draw_id, 4) Then
                Call draw_and_score(player, 4)
            Else
                If score_pile(player, 0) > -1 Then
                    found = 0
                    min_index = -1
                    min = 99
                    For i = num_players - 1 To 0 Step -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object scores(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If scores(i) < scores(player) Then
                            found = 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If scores(i) < min Then
                                If (min_index = -1 And player = 0) Or player <> 0 Then
                                    'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    min = scores(i)
                                    min_index = i
                                End If
                            End If
                        End If
                    Next i
                    If found = 1 Then
                        If player = 0 And ai_mode <> 1 Then
                            Call show_prompt("optics", dogma(id, level - 1) & " (choose the card in your score pile first)", 0)
                        Else
                            'alert ("still good" & player & min_index)
                            Call transfer_card_from_score(player, min_index, 0)
                        End If
                    End If

                End If
            End If

        ElseIf card_name = "Paper" Then
            If level = 1 Then
                splay_options(3) = 1
                splay_options(4) = 1
                Call perform_best_splay(player, splay_options, "Left")
            Else
                For i = 0 To 4
                    If splayed(player, i) = "Left" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If player <> active_player Then dogma_copied = 1
                        Call draw_num(player, 4)
                    End If
                Next i
            End If

        ElseIf card_name = "Translation" Then
            If level = 1 Then
                If score_pile(player, 0) > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("translation", dogma(id, level - 1), 1)
                    Else
                        ' See if we can meld into all crowns
                        Call save(player)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = size2(score_pile, player) - 1 To 0 Step -1
                            If Not card_has_symbol(score_pile(player, i), 4) Then
                                Call meld(player, score_pile(player, i))
                                Call remove_card_from_score_pile(player, i)
                            End If
                        Next i
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        While (size2(score_pile, player) > 0)
                            Call meld(player, score_pile(player, 0))
                            Call remove_card_from_score_pile(player, 0)
                        End While
                        Call restore(player)

                        found = 0
                        For i = 0 To 4
                            If board(player, i, 0) > -1 Then
                                If Not card_has_symbol(board(player, i, 0), 4) Then
                                    found = 1
                                End If
                            End If
                        Next i

                        If found = 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If player <> active_player Then dogma_copied = 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            For i = size2(score_pile, player) - 1 To 0 Step -1
                                If Not card_has_symbol(score_pile(player, i), 4) Then
                                    Call meld(player, score_pile(player, i))
                                    Call remove_card_from_score_pile(player, i)
                                End If
                            Next i
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            While (size2(score_pile, player) > 0)
                                Call meld(player, score_pile(player, 0))
                                Call remove_card_from_score_pile(player, 0)
                            End While
                            Exit Sub
                        End If

                        ' Meld the pile low to high
                        Call save(player)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        While (size2(score_pile, player) > 0)
                            Call meld(player, score_pile(player, 0))
                            Call remove_card_from_score_pile(player, 0)
                        End While
                        Call restore(player)
                        If current_score > original_score Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If player <> active_player Then dogma_copied = 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            While (size2(score_pile, player) > 0)
                                Call meld(player, score_pile(player, 0))
                                Call remove_card_from_score_pile(player, 0)
                            End While
                        End If
                    End If
                End If
            Else
                found = 0
                For i = 0 To 4
                    If board(player, i, 0) > -1 Then
                        If Not card_has_symbol(board(player, i, 0), 4) Then
                            found = 1
                        End If
                    End If
                Next i
                If found = 0 Then
                    Call achieve_special(player, 12, "World")
                End If
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 4
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 4")
#End If
        ElseIf card_name = "Anatomy" Then
            If score_pile(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("anatomy", dogma(id, level - 1), 0)
                Else
                    max = 0
                    max_index = -1
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = 0 To size2(score_pile, player) - 1
                        Call save(player)
                        min = age(score_pile(player, i))
                        Call return_from_score_pile(player, i)
                        'UPGRADE_WARNING: Couldn't resolve default property of object found. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        found = 0
                        For j = 0 To 4
                            If found = 0 And board(player, j, 0) > -1 Then
                                If age(board(player, j, 0)) = min Then
                                    Call return_from_board(player, j)
                                    found = 1
                                End If
                            End If
                        Next j
                        Call restore(player)
                        If current_score > max Then
                            max = current_score
                            max_index = i
                        End If
                    Next i
                    min = age(score_pile(player, max_index))
                    Call return_from_score_pile(player, max_index)
                    For j = 0 To 4
                        If board(player, j, 0) > -1 Then
                            If age(board(player, j, 0)) = min Then
                                Call return_from_board(player, j)
                                Exit Sub
                            End If
                        End If
                    Next j
                End If
            End If

        ElseIf card_name = "Colonialism" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            found = 0
            While found = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_tuck(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                draw_id = draw_and_tuck(player, 3)
                If Not card_has_symbol(draw_id, 4) Then found = 1
            End While

        ElseIf card_name = "Enterprise" Then
            If level = 1 Then
                found = 0
                min = 11
                min_index = -1
                For i = 0 To 4
                    If board(player, i, 0) > -1 And color_lookup(i) <> "Purple" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(board(player, i, 0), 4) Then
                            found = 1
                            If age(board(player, i, 0)) < min Then
                                min = age(board(player, i, 0))
                                min_index = i
                            End If
                        End If
                    End If
                Next i
                If found = 1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("enterprise", dogma(id, level - 1), 0)
                    Else
                        Call transfer_card_on_board(player, active_player, min_index)
                        Call draw_and_meld(player, 4)
                    End If
                End If
            Else
                splay_options(4) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Experimentation" Then
            Call draw_and_meld(player, 5)
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1

        ElseIf card_name = "Gunpowder" Then
            If level = 1 Then
                found = 0
                For i = 0 To 4
                    min = 99
                    If board(player, i, 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(board(player, i, 0), 2) Then
                            found = 1
                            If age(board(player, i, 0)) < min Then
                                min_index = i
                                min = board(player, i, 0)
                            End If
                            already_returned(i) = 1
                        End If
                    End If
                Next i
                If found > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    demand_met = 1
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("gunpowder", dogma(id, level - 1), 0)
                    Else
                        Call transfer_board_to_score(player, active_player, min_index)
                    End If
                End If

            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If demand_met = 1 Then
                    Call draw_and_score(player, 2)
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If player <> active_player Then dogma_copied = 1
                End If
            End If

        ElseIf card_name = "Invention" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object found. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                found = 0
                For i = 0 To 4
                    splay_options(i) = 0
                    If splayed(player, i) = "Left" And can_splay(player, i) Then
                        found = 1
                        splay_options(i) = 1
                    End If
                Next i
                If found = 1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("invention", dogma(id, level - 1), 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object perform_best_splay(player, splay_options, Right). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If perform_best_splay(player, splay_options, "Right") > -1 Then
                            Call draw_and_score(player, 4)
                        End If
                    End If
                End If
            Else
                count = 0
                For i = 0 To 4
                    If splayed(player, i) <> "" Then count = count + 1
                Next i
                If count = 5 Then
                    Call achieve_special(player, 12, "World")
                End If
            End If

        ElseIf card_name = "Navigation" Then
            found = 0
            min_index = -1
            min = 4
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 0 To size2(score_pile, player) - 1
                If age(score_pile(player, i)) = 2 Or age(score_pile(player, i)) = 3 Then
                    found = 1
                    If age(score_pile(player, i)) < min Then
                        min = age(score_pile(player, i))
                        min_index = i
                    End If
                End If
            Next i
            If found = 1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("navigation", dogma(id, level - 1), 0)
                Else
                    Call transfer_card_from_score(player, active_player, min_index)
                End If
            End If

        ElseIf card_name = "Perspective" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("perspective", dogma(id, level - 1), 1)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If size2(hand, player) > 1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If player <> active_player Then dogma_copied = 1
                        Call return_from_hand(player, 0)
                        For i = 1 To Int(icon_total(player, 3) / 2)
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If size2(hand, player) > 0 Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Call score_from_hand(player, size2(hand, player) - 1)
                            End If
                        Next i
                    End If
                End If
            End If

        ElseIf card_name = "Printing Press" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(score_pile, player) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("printingPress", dogma(id, level - 1), 1)
                    Else
                        count = 0
                        If board(player, 2, 0) > -1 Then count = age(board(player, 2, 0))
                        Call save(player)
                        Call return_from_score_pile(player, 0)
                        Call draw_num(player, count + 2)
                        Call restore(player)
                        If current_score > original_score Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If player <> active_player Then dogma_copied = 1
                            Call return_from_score_pile(player, 0)
                            Call draw_num(player, count + 2)
                        End If
                    End If
                End If
            Else
                splay_options(3) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Reformation" Then
            If level = 1 Then
                If icon_total(player, 1) > 1 And hand(player, 0) > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        human_data = Int(icon_total(player, 1) / 2)
                        Call show_prompt("reformation1", dogma(id, level - 1), 1)
                    Else
                        max = Int(icon_total(player, 1) / 2)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If max > size2(hand, player) Then max = size2(hand, player)
                        Call save(player)
                        For i = 1 To max
                            Call tuck_from_hand(player, 0)
                        Next i
                        'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If can_splay(player, 0) Then splayed(player, 0) = "Right"
                        'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If can_splay(player, 2) Then splayed(player, 2) = "Right"
                        Call restore(player)
                        If current_score > original_score Then
                            For i = 1 To max
                                Call tuck_from_hand(player, 0)
                            Next i
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If player <> active_player Then dogma_copied = 1
                        End If
                    End If
                End If
            Else
                splay_options(0) = 1
                splay_options(2) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 5
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 5")
#End If
        ElseIf card_name = "Astronomy" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
                found = 0
                While found = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    draw_id = draw_num(player, 6)
                    If color(draw_id) = 3 Or color(draw_id) = 4 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(player, find_card(player, draw_id))
                    Else
                        found = 1
                    End If
                End While
            Else
                found = 1
                For i = 0 To 4
                    If board(player, i, 0) > -1 And i <> 2 Then
                        If age(board(player, i, 0)) < 6 Then found = 0
                    End If
                Next i
                If found = 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If player <> active_player Then dogma_copied = 1
                    Call achieve_special(player, 13, "Universe")
                End If
            End If

        ElseIf card_name = "Banking" Then
            If level = 1 Then
                found = 0
                min = 11
                min_index = -1
                For i = 0 To 4
                    If board(player, i, 0) > -1 And color_lookup(i) <> "Green" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(board(player, i, 0), 5) Then
                            found = 1
                            If age(board(player, i, 0)) < min Then
                                min = age(board(player, i, 0))
                                min_index = i
                            End If
                        End If
                    End If
                Next i
                If found = 1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("banking", dogma(id, level - 1), 0)
                    Else
                        Call transfer_card_on_board(player, active_player, min_index)
                        Call draw_and_score(player, 5)
                    End If
                End If
            Else
                splay_options(4) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Chemistry" Then
            If level = 1 Then
                splay_options(3) = 1
                Call perform_best_splay(player, splay_options, "Right")
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object highest_top_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_score(player, highest_top_card(player) + 1)
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("chemistry2", dogma(id, level - 1), 0)
                Else
                    Call return_from_score_pile(player, 0)
                End If
            End If

        ElseIf card_name = "Coal" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
                Call draw_and_tuck(player, 5)
            ElseIf level = 2 Then
                splay_options(1) = 1
                Call perform_best_splay(player, splay_options, "Right")
            Else
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("coal3", dogma(id, level - 1), 1)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    max = score_game_individual(player)
                    max_index = -1
                    For i = 0 To 4
                        If board(player, i, 0) > -1 Then
                            Call save(player)
                            Call score_top_card(player, i)
                            If board(player, i, 0) > -1 Then Call score_top_card(player, i)
                            Call restore(player)
                            If current_score > max Then
                                max = current_score
                                max_index = i
                            End If
                        End If
                    Next i
                    If max_index > -1 Then
                        Call score_top_card(player, max_index)
                        If board(player, max_index, 0) > -1 Then Call score_top_card(player, max_index)
                    End If
                End If
            End If

        ElseIf card_name = "Measurement" Then
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("measurement", dogma(id, level - 1), 1)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If player <> active_player Then dogma_copied = 1
                    max = -1
                    For i = -1 To 4
                        Call save(player)
                        If i = -1 Then
                            ' do nothing
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If can_splay(player, i) Then
                                Call return_from_hand(player, 0)
                                Call splay(player, i, "Right")
                                'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Call draw_num(player, size3(board, player, i))
                            End If
                        End If
                        Call restore(player)
                        If current_score > max Then
                            max = current_score
                            max_index = i
                        End If
                    Next i
                    If max_index > -1 Then
                        Call return_from_hand(player, 0)
                        Call splay(player, max_index, "Right")
                        'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call draw_num(player, size3(board, player, max_index))
                    End If
                End If
            End If

        ElseIf card_name = "Physics" Then
            found = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            For i = 0 To 2
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                draw_id = draw_num(player, 6)
                colors_found(i) = color(draw_id)
            Next i
            If (colors_found(0) = colors_found(1)) Or (colors_found(0) = colors_found(2)) Or (colors_found(2) = colors_found(1)) Then found = 1
            If found = 1 Then
                Call append(player, "revealed matching colors, and must return their entire hand")
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("physics", "Return all the cards from your hand", 0)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        Call return_from_hand(player, 0)
                    Next i
                End If
            End If

        ElseIf card_name = "Societies" Then
            found = 0
            min = 11
            min_index = -1
            For i = 0 To 4
                If board(player, i, 0) > -1 And color_lookup(i) <> "Purple" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If card_has_symbol(board(player, i, 0), 3) Then
                        found = 1
                        If age(board(player, i, 0)) < min Then
                            min = age(board(player, i, 0))
                            min_index = i
                        End If
                    End If
                End If
            Next i
            If found = 1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("societies", dogma(id, level - 1), 0)
                Else
                    Call transfer_card_on_board(player, active_player, min_index)
                    Call draw_num(player, 5)
                End If
            End If

        ElseIf card_name = "Statistics" Then
            If level = 1 Then
                If score_pile(player, 0) > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("statistics", dogma(id, level - 1), 0)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_score_to_hand(player, size2(score_pile, player) - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If score_pile(player, 0) > -1 And size2(hand, player) = 1 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Call transfer_score_to_hand(player, size2(score_pile, player) - 1)
                        End If
                    End If
                End If
            Else
                splay_options(0) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Steam Engine" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            Call draw_and_tuck(player, 4)
            Call draw_and_tuck(player, 4)
            If board(player, 0, 0) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                draw_id = board(player, 0, size3(board, player, 0) - 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                board(player, 0, size3(board, player, 0) - 1) = -1
                Call score(player, draw_id)
            End If

        ElseIf card_name = "The Pirate Code" Then
            If level = 1 Then
                found = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = 0 To size2(score_pile, player) - 1
                    If age(score_pile(player, i)) < 5 Then found = found + 1
                Next i
                If found > 2 Then found = 2
                If found > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    demand_met = 1
                    If player = 0 And ai_mode <> 1 Then
                        human_data = found
                        Call show_prompt("pirateCode", dogma(id, level - 1), 0)
                    Else
                        Call transfer_card_from_score(player, active_player, 0)
                        If found > 1 Then Call transfer_card_from_score(player, active_player, 0)
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If demand_met = 0 Then Exit Sub
                found = 0
                min_index = -1
                min = 9
                For i = 0 To 4
                    If board(player, i, 0) > -1 Then
                        If card_has_symbol(board(player, i, 0), 4) And age(board(player, i, 0)) < min Then
                            min = age(board(player, i, 0))
                            min_index = 1
                            found = 1
                        End If
                    End If
                Next i
                If found > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    If player = 0 And ai_mode <> 1 Then
                        human_data = min
                        Call show_prompt("pirateCode2", dogma(id, level - 1), 0)
                    Else
                        Call score_top_card(player, min_index)
                    End If
                End If
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 6
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 6")
#End If

        ElseIf card_name = "Atomic Theory" Then
            If level = 1 Then
                splay_options(3) = 1
                Call perform_best_splay(player, splay_options, "Right")
            Else
                Call draw_and_meld(player, 7)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
            End If

        ElseIf card_name = "Canning" Then
            If level = 1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("canning", dogma(id, level - 1), 2)
                Else
                    Call save(player)
                    Call draw_and_tuck(player, 6)
                    For i = 0 To 4
                        If board(player, i, 0) > -1 Then
                            If Not card_has_symbol(board(player, i, 0), 5) Then
                                Call score_top_card(player, i)
                            End If
                        End If
                    Next i
                    Call restore(player)
                    If current_score > original_score Then
                        Call draw_and_tuck(player, 6)
                        For i = 0 To 4
                            If board(player, i, 0) > -1 Then
                                If Not card_has_symbol(board(player, i, 0), 5) Then
                                    Call score_top_card(player, i)
                                End If
                            End If
                        Next i
                    End If
                End If
            Else
                splay_options(0) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Classification" Then
            If hand(player, 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("classification", dogma(id, level - 1), 0)
                Else
                    max = 0
                    max_index = -1
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = 0 To size2(hand, player) - 1
                        found = color(hand(player, i))
                        If already_returned(found) = 0 Then
                            already_returned(found) = 1
                            Call save(player)
                            Call collect_all_cards(player, found)
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            For j = size2(hand, player) - 1 To 0 Step -1
                                If color(hand(player, j)) = found Then
                                    Call meld_from_hand(player, j)
                                End If
                            Next j
                            Call restore(player)
                            If current_score > max Then
                                max = current_score
                                max_index = i
                                'alert ("found a max score at " & max_index)
                            End If
                        End If
                    Next i
                    found = color(hand(player, max_index))
                    Call collect_all_cards(player, found)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For j = size2(hand, player) - 1 To 0 Step -1
                        If color(hand(player, j)) = found Then
                            Call meld_from_hand(player, j)
                        End If
                    Next j
                End If
            End If

        ElseIf card_name = "Democracy" Then
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    human_data = 0
                    Call show_prompt("democracy", dogma(id, level - 1) & "  (Must return at least " & democracy_minimum & " to draw and score the 8)", 1)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If size2(hand, player) >= democracy_minimum Then
                        Call save(player)
                        For i = 1 To democracy_minimum
                            Call return_from_hand(player, 0)
                        Next i
                        Call draw_and_score(player, 8)
                        Call restore(player)
                        If current_score > original_score Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If active_player <> player Then dogma_copied = 1
                            For i = 1 To democracy_minimum
                                Call return_from_hand(player, 0)
                            Next i
                            Call draw_and_score(player, 8)
                            democracy_minimum = democracy_minimum + 1
                        End If
                    End If
                End If
            End If

        ElseIf card_name = "Emancipation" Then
            If level = 1 Then
                If hand(player, 0) > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("emancipation", dogma(id, level - 1), 0)
                    Else
                        Call transfer_hand_to_score(player, active_player, 0)
                        Call draw_num(player, 6)
                    End If
                End If
            Else
                splay_options(1) = 1
                splay_options(2) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Encyclopedia" Then
            If score_pile(player, 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                max = age(score_pile(player, size2(score_pile, player) - 1))
                If player = 0 And ai_mode <> 1 Then
                    human_data = max
                    Call show_prompt("encyclopedia", dogma(id, level - 1), 1)
                Else
                    Call save(player)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(score_pile, player) - 1 To 0 Step -1
                        If age(score_pile(player, i)) = max Then
                            Call meld(player, score_pile(player, i))
                            Call remove_card_from_score_pile(player, i)
                        End If
                    Next i
                    Call restore(player)
                    If current_score > original_score Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = size2(score_pile, player) - 1 To 0 Step -1
                            If age(score_pile(player, i)) = max Then
                                Call meld(player, score_pile(player, i))
                                Call remove_card_from_score_pile(player, i)
                            End If
                        Next i
                    End If
                End If

            End If

        ElseIf card_name = "Industrialization" Then
            If level = 1 Then
                If icon_total(player, 5) > 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    count = Int(icon_total(player, 5) / 2)
                    For i = 1 To count
                        Call draw_and_tuck(player, 6)
                    Next i
                End If
            Else
                splay_options(1) = 1
                splay_options(2) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Machine Tools" Then
            If score_pile(player, 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_score(player, age(score_pile(player, size2(score_pile, player) - 1)))
            End If

        ElseIf card_name = "Metric System" Then
            If level = 1 Then
                If splayed(player, 4) = "Right" Then
                    For i = 0 To 4
                        splay_options(i) = 1
                    Next i
                    Call perform_best_splay(player, splay_options, "Right")
                End If
            Else
                splay_options(4) = 1
                Call perform_best_splay(player, splay_options, "Right")
            End If

        ElseIf card_name = "Vaccination" Then
            'Call update_display
            'Exit Sub
            If level = 1 Then
                If score_pile(player, 0) > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    demand_met = 1
                    min = age(score_pile(player, 0))
                    human_data = min
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("vaccination", dogma(id, level - 1), 0)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = size2(score_pile, player) - 1 To 0 Step -1
                            If age(score_pile(player, i)) = min Then
                                'Call alert("returning " & i & " " & title(score_pile(player, i)))
                                Call return_from_score_pile(player, i)
                            End If
                        Next i
                        Call draw_and_meld(player, 6)
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If demand_met = 1 Then Call draw_and_meld(player, 7)
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 7
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 7")
#End If
        ElseIf card_name = "Bicycle" Then
            If player = 0 And ai_mode <> 1 Then
                Call show_prompt("bicycle", dogma(id, level - 1), 2)
            Else
                Call save(player)
                Call exchange_bicycle(player)
                Call restore(player)
                If current_score > original_score Then
                    Call exchange_bicycle(player)
                End If

            End If

        ElseIf card_name = "Combustion" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            found = size2(score_pile, player)
            If found > 2 Then found = 2
            If found > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                demand_met = 1
                If player = 0 And ai_mode <> 1 Then
                    human_data = found
                    Call show_prompt("combustion", dogma(id, level - 1), 0)
                Else
                    Call transfer_card_from_score(player, active_player, 0)
                    If found > 1 Then Call transfer_card_from_score(player, active_player, 0)
                End If
            End If

        ElseIf card_name = "Electricity" Then
            found = 0
            For i = 0 To 4
                If board(player, i, 0) > -1 Then
                    If Not card_has_symbol(board(player, i, 0), 5) Then found = found + 1
                End If
            Next i
            If found > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                If player = 0 And ai_mode <> 1 Then
                    For i = 0 To 4
                        already_returned(i) = 0
                    Next i
                    human_data = found
                    Call show_prompt("electricity", dogma(id, level - 1), 0)
                Else
                    For i = 0 To 4
                        If board(player, i, 0) > -1 Then
                            If Not card_has_symbol(board(player, i, 0), 5) Then
                                Call return_from_board(player, i)
                            End If
                        End If
                    Next i
                    For i = 1 To found
                        Call draw_num(player, 8)
                    Next i
                End If
            End If

        ElseIf card_name = "Evolution" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            If player = 0 And ai_mode <> 1 Then
                cmdYes.Text = "Draw and Score an 8"
                Call show_prompt("evolution", dogma(id, level - 1), 2)
            Else
                If score_pile(player, 0) = -1 Then
                    Call draw_and_score(player, 8)
                    Call return_from_score_pile(player, 0)
                    Exit Sub
                End If

                Call save(player)
                Call draw_and_score(player, 8)
                Call return_from_score_pile(player, 0)
                Call restore(player)
                max = current_score
                Call save(player)
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, age(score_pile(player, size2(score_pile, player) - 1)) + 1)
                Call restore(player)
                If current_score > max Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_num(player, age(score_pile(player, size2(score_pile, player) - 1)) + 1)
                Else
                    Call draw_and_score(player, 8)
                    Call return_from_score_pile(player, 0)
                End If
            End If

        ElseIf card_name = "Explosives" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            found = size2(hand, player)
            If found > 3 Then found = 3
            If found > 0 Then
                If player = 0 And ai_mode <> 1 Then
                    human_data = found
                    Call show_prompt("explosives", dogma(id, level - 1), 0)
                Else
                    For i = 1 To found
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_in_hand(player, active_player, size2(hand, player) - 1)
                    Next i
                    If hand(player, 0) = -1 Then Call draw_num(player, 7)
                End If
            End If

        ElseIf card_name = "Lighting" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                For i = 0 To 9
                    already_returned(i) = 0
                Next i
                If player = 0 And ai_mode <> 1 Then
                    human_data = 3
                    Call show_prompt("lighting", dogma(id, level - 1), 1)
                Else
                    ' Return a card from each age
                    count = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        If already_returned(age(hand(player, i))) = 0 And count < 3 And age(hand(player, i)) < 9 Then
                            already_returned(age(hand(player, i))) = 1
                            Call tuck_from_hand(player, i)
                            count = count + 1
                            If hand(player, 0) > -1 Then i = -1
                        End If
                    Next i
                    For i = 0 To count - 1
                        Call draw_and_score(player, 7)
                    Next i
                End If
            End If

        ElseIf card_name = "Publications" Then
            If level = 1 Then
                found = 0
                max_index = -1
                max_index2 = -1
                max = 0
                For i = 0 To 4
                    'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If size3(board, player, i) > 1 Then
                        found = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For j = 1 To size3(board, player, i) - 1
                            count = age(board(player, i, j)) - age(board(player, i, 0))
                            If count > max Then
                                max = count
                                max_index = i
                                max_index2 = j
                                'alert ("Found a max at " & i & j)
                            End If
                        Next j
                    End If
                Next i
                If found > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("publications", dogma(id, level - 1), 1)
                    ElseIf max_index > -1 Then
                        ' Put the largest card in the pile on top
                        count = board(player, max_index, 0)
                        'alert ("Setting top of board to " & max_index & max_index2 & title(board(player, max_index, max_index2)))
                        board(player, max_index, 0) = board(player, max_index, max_index2)
                        board(player, max_index, max_index2) = count
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        Call update_display()
                    End If
                End If
            Else
                splay_options(0) = 1
                splay_options(3) = 1
                Call perform_best_splay(player, splay_options, "Up")
            End If

        ElseIf card_name = "Railroad" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                If player = 0 And ai_mode <> 1 Then
                    If hand(player, 0) = -1 Then
                        For i = 1 To 3
                            Call draw_num(player, 6)
                        Next i
                    Else
                        Call show_prompt("railroad", dogma(id, level - 1), 0)
                    End If
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        Call return_from_hand(player, 0)
                    Next i
                    For i = 1 To 3
                        Call draw_num(player, 6)
                    Next i
                End If
            Else
                found = 0
                For i = 0 To 4
                    splay_options(i) = 0
                    If splayed(player, i) = "Right" And can_splay(player, i) Then
                        found = 1
                        splay_options(i) = 1
                    End If
                Next i
                If found = 1 Then
                    Call perform_best_splay(player, splay_options, "Up")
                End If
            End If

        ElseIf card_name = "Refrigeration" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                count = Int(size2(hand, player) / 2)
                If count > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        human_data = count
                        Call show_prompt("refrigeration", dogma(id, level - 1), 0)
                    Else
                        For i = 1 To count
                            Call return_from_hand(player, 0)
                        Next i
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("refrigeration2", dogma(id, level - 1), 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call score_from_hand(player, size2(hand, player) - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                    End If
                End If
            End If

        ElseIf card_name = "Sanitation" Then
            ' Store  the active player's lowest cards
            already_returned(0) = -1
            If hand(active_player, 0) > -1 Then
                already_returned(0) = hand(active_player, 0)
            End If

            If player = 0 And ai_mode <> 1 Then
                human_data = 0
                Call show_prompt("sanitation", dogma(id, level - 1), 0)
            Else
                'MsgBox "player = 0"
                If active_player = 0 And hand(active_player, 0) > -1 Then
                    human_data = player
                    Call show_prompt("sanitation2", dogma(id, level - 1), 0)
                Else
                    'MsgBox ("here we go " & already_returned(0))
                    If hand(active_player, 0) > -1 Then Call remove_card_from_hand(active_player, 0)
                    If hand(player, 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_in_hand(player, active_player, size2(hand, player) - 1)
                        If hand(player, 0) > -1 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Call transfer_card_in_hand(player, active_player, size2(hand, player) - 1)
                        End If
                    End If
                    If already_returned(0) > -1 Then
                        Call Push2(hand, player, already_returned(0))
                        Call append(active_player, "transfers a " & age(already_returned(0)) & " from his hand to " & player + 1 & "'s hand")
                    End If
                End If
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 8
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 8")
#End If
        ElseIf card_name = "Antibiotics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                For i = 0 To 9
                    already_returned(i) = 0
                Next i
                If player = 0 And ai_mode <> 1 Then
                    human_data = 3
                    Call show_prompt("antibiotics", dogma(id, level - 1), 1)
                Else
                    ' Return a card from each age
                    count = 0
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        If already_returned(age(hand(player, i))) = 0 And count < 3 Then
                            already_returned(age(hand(player, i))) = 1
                            Call return_from_hand(player, i)
                            count = count + 1
                            If hand(player, 0) > -1 Then i = -1
                        End If
                    Next i
                    For i = 0 To count - 1
                        Call draw_num(player, 8)
                        Call draw_num(player, 8)
                    Next i
                End If
            End If

        ElseIf card_name = "Corporations" Then
            If level = 1 Then
                found = 0
                min = 11
                min_index = -1
                For i = 0 To 4
                    If board(player, i, 0) > -1 And color_lookup(i) <> "Green" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(board(player, i, 0), 5) Then
                            found = 1
                            If age(board(player, i, 0)) < min Then
                                min = age(board(player, i, 0))
                                min_index = i
                            End If
                        End If
                    End If
                Next i
                If found = 1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("corporations", dogma(id, level - 1), 0)
                    Else
                        Call transfer_board_to_score(player, active_player, min_index)
                        Call draw_and_meld(player, 8)
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call draw_and_meld(player, 8)
            End If

        ElseIf card_name = "Empiricism" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                If player = 0 And ai_mode <> 1 Then
                    human_data = 2
                    For i = 0 To 4
                        cmdOddball(i).Visible = True
                        cmdOddball(i).Text = color_lookup(i)
                    Next i
                    Call show_prompt("empiricism", dogma(id, level - 1), 0)
                Else
                    ' AI always guesses right.  How lucky.
                    'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    draw_id = draw_and_meld(player, 9)
                    Call splay(player, color(draw_id), "Up")
                End If
            Else
                If icon_total(player, 3) >= 20 Then Call end_game_generic(player, "Player " & player + 1 & " has 20 or more Lightbulbs, and wons the game via Empiricism!")
            End If

        ElseIf card_name = "Flight" Then
            If level = 1 Then
                If splayed(player, 1) = "Up" Then
                    For i = 0 To 4
                        splay_options(i) = 1
                    Next i
                    Call perform_best_splay(player, splay_options, "Up")
                End If
            Else
                splay_options(1) = 1
                Call perform_best_splay(player, splay_options, "Up")
            End If

        ElseIf card_name = "Mass Media" Then
            If level = 1 Then
                If hand(player, 0) > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("massMedia", dogma(id, level - 1), 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object score_game(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        max_score = score_game(player)
                        max = max_score
                        max_index = -1
                        phase = "massMedia"
                        For i = 1 To 10
                            Call save(player)
                            Call return_from_hand(player, 0)
                            'MsgBox "returning from score " & i
                            Call return_from_all_score_piles(i)
                            'UPGRADE_WARNING: Couldn't resolve default property of object score_game(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            max_score = score_game(player)
                            If max_score > max Then
                                max = max_score
                                max_index = i
                            End If
                            Call restore(player)
                        Next i
                        If max_index > -1 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If active_player <> player Then dogma_copied = 1
                            Call return_from_hand(player, 0)
                            Call return_from_all_score_piles(max_index)
                        End If

                    End If
                End If
            Else
                splay_options(2) = 1
                Call perform_best_splay(player, splay_options, "Up")
            End If

        ElseIf card_name = "Mobility" Then
            If level = 1 Then
                found = 0
                max = -1
                max_index = -1
                max_index2 = -1
                For i = 0 To 4
                    If board(player, i, 0) > -1 And color_lookup(i) <> "Red" Then
                        If Not card_has_symbol(board(player, i, 0), 5) Then
                            'If player = 1 Then MsgBox "inspecting " & i
                            If age(board(player, i, 0)) > max Then
                                'If player = 1 Then MsgBox "max " & i
                                found = found + 1
                                max = age(board(player, i, 0))
                                max_index2 = max_index
                                max_index = i
                            ElseIf max_index2 > -1 Then
                                'If player = 1 Then MsgBox "max2 upgrade" & i
                                If age(board(player, i, 0)) > age(board(player, max_index2, 0)) Then
                                    found = found + 1
                                    max_index2 = i
                                End If
                            ElseIf max_index2 = -1 Then
                                'If player = 1 Then MsgBox "max2 " & i
                                found = found + 1
                                max_index2 = i
                            Else
                                'If player = 1 Then MsgBox "nothing " & i
                            End If
                        End If
                    End If
                Next i
                If found > 0 Then
                    If found > 2 Then found = 2
                    If player = 0 And ai_mode <> 1 Then
                        For i = 0 To 4
                            already_returned(i) = 0
                        Next i
                        human_data = found
                        Call show_prompt("mobility", dogma(id, level - 1), 0)
                    Else
                        Call transfer_board_to_score(player, active_player, max_index)
                        If max_index2 > -1 Then Call transfer_board_to_score(player, active_player, max_index2)
                        Call draw_num(player, 8)
                    End If
                End If
            End If

        ElseIf card_name = "Quantum Theory" Then
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    human_data = 0
                    Call show_prompt("quantumTheory", dogma(id, level - 1), 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf size2(hand, player) > 1 Then
                    ' Return worst 2 cards
                    For i = 1 To 2
                        Call return_from_hand(player, 0)
                    Next i
                    Call draw_num(player, 10)
                    Call draw_and_score(player, 10)
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                End If
            End If

        ElseIf card_name = "Rocketry" Then
            If icon_total(player, 6) > 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("rocketry", dogma(id, level - 1) & vbNewLine & "Click on a player's number near their score to select that player.", 0)
                Else
                    ' Always kill the human!!
                    For i = 1 To Int(icon_total(player, 6) / 2)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If score_pile(0, 0) > -1 Then Call return_from_score_pile(0, size2(score_pile, 0) - 1)
                    Next i
                End If
            End If

        ElseIf card_name = "Skyscrapers" Then
            found = 0
            min = 11
            min_index = -1
            For i = 0 To 4
                If board(player, i, 0) > -1 And color_lookup(i) <> "Yellow" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, i, 0), 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If card_has_symbol(board(player, i, 0), 6) Then
                        found = 1
                        If age(board(player, i, 0)) < min Then
                            min = age(board(player, i, 0))
                            min_index = i
                        End If
                    End If
                End If
            Next i
            If found = 1 Then
                If player = 0 And ai_mode <> 1 Then
                    human_data = 0
                    Call show_prompt("skyscrapers", dogma(id, level - 1), 0)
                Else
                    Call transfer_card_on_board(player, active_player, min_index)
                    If board(player, min_index, 0) > -1 Then Call score_top_card(player, min_index)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, min_index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size3(board, player, min_index) - 1 To 0 Step -1
                        Call return_from_board(player, min_index)
                    Next i
                End If
            End If

        ElseIf card_name = "Socialism" Then
            If player = 0 And ai_mode <> 1 Then
                human_data = 0
                Call show_prompt("socialism", dogma(id, level - 1), 1)
            Else
                Call save(player)
                found = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = size2(hand, player) - 1 To 0 Step -1
                    If color_lookup(color(hand(player, 0))) = "Purple" Then found = 1
                    Call tuck_from_hand(player, 0)
                Next i
                If found = 1 Then Call steal_lowest_cards(player)
                Call restore(player)
                If current_score > original_score Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        Call tuck_from_hand(player, 0)
                    Next i
                    If found = 1 Then Call steal_lowest_cards(player)
                End If
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 9
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 9")
#End If
        ElseIf card_name = "Collaboration" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                already_returned(0) = draw_num(player, 9)
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                already_returned(1) = draw_num(player, 9)
                If active_player = 0 And ai_mode <> 1 Then
                    human_data = player
                    Call show_prompt("collaboration", "Player  " & player + 1 & " drew " & title(already_returned(0)) & " and " & title(already_returned(1)) & ".  Would you like to take " & title(already_returned(0)) & "?", 2)
                Else
                    ' The active player will steal the one that benefits them most
                    'UPGRADE_WARNING: Couldn't resolve default property of object age1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    age1 = age(already_returned(0))
                    age2 = age(already_returned(1))
                    If board(active_player, color(already_returned(0)), 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object age1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        age1 = age(already_returned(0)) - age(board(active_player, color(already_returned(0)), 0))
                    End If
                    If board(active_player, color(already_returned(1)), 0) > -1 Then
                        age2 = age(already_returned(1)) - age(board(active_player, color(already_returned(1)), 0))
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object age1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If age1 > age2 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(player, find_card(player, already_returned(1)))
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_in_hand(player, active_player, find_card(player, already_returned(0)))
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(active_player, find_card(active_player, already_returned(0)))
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(player, find_card(player, already_returned(0)))
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_in_hand(player, active_player, find_card(player, already_returned(1)))
                        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(active_player, find_card(active_player, already_returned(1)))
                    End If
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size3(board, player, 4) > 9 Then
                    Call end_game_generic(player, "Player " & player + 1 & " has 10 or mGreen cards howing, and wins the game via Collaboration!")
                End If
            End If

        ElseIf card_name = "Composites" Then
            If player = 0 And ai_mode <> 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 1 Then
                    Call show_prompt("composites", dogma(id, level - 1), 0)
                ElseIf score_pile(player, 0) > -1 Then
                    Call show_prompt("composites2", "Transfer the highest card in your score pile.", 0)
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(hand, player) > 1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 1 Step -1
                        Call transfer_card_in_hand(player, active_player, 0)
                    Next i
                End If
                If score_pile(player, 0) > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_from_score(player, active_player, size2(score_pile, player) - 1)
                End If
            End If

        ElseIf card_name = "Computers" Then
            If level = 1 Then
                splay_options(1) = 1
                splay_options(4) = 1
                Call perform_best_splay(player, splay_options, "Up")
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                draw_id = draw_and_meld(player, 10)
                Call perform_solo_dogma_effects(player, draw_id)
            End If

        ElseIf card_name = "Ecology" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("ecology", dogma(id, level - 1), 1)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If size2(hand, player) > 1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If player <> active_player Then dogma_copied = 1
                        Call return_from_hand(player, 0)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call score_from_hand(player, size2(hand, player) - 1)
                        Call draw_num(player, 10)
                        Call draw_num(player, 10)
                    End If
                End If
            End If

        ElseIf card_name = "Fission" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If demand_met = 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    draw_id = draw_num(player, 10)
                    If color_lookup(color(draw_id)) = "Red" Then
                        ' Kaboom!
                        Call append(player, "revealed a Red card.  Kaboom!")
                        'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        demand_met = 1
                        For i = 0 To num_players - 1
                            For j = 0 To 4
                                board(i, j, 0) = -1
                            Next j
                            score_pile(i, 0) = -1
                            hand(i, 0) = -1
                        Next i
                        Call update_display()
                    End If
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf demand_met = 0 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("fission", dogma(id, level - 1), 0)
                Else
                    ' Find the max top card to return
                    max_index = -1
                    max = -1
                    For i = 0 To num_players - 1
                        For j = 0 To 4
                            If i <> player And board(i, j, 0) <> -1 Then
                                If age(board(i, j, 0)) > max Then
                                    max = age(board(i, j, 0))
                                    max_index = i
                                    max_index2 = j
                                End If
                            End If
                        Next j
                    Next i
                    If max_index = -1 Then
                        For j = 0 To 4
                            If board(player, j, 0) <> -1 And j <> 1 Then
                                If age(board(player, j, 0)) > max Then
                                    max = age(board(player, j, 0))
                                    max_index = player
                                    max_index2 = j
                                End If
                            End If
                        Next j
                    End If

                    If max_index > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        Call return_from_board(max_index, max_index2)
                    End If
                End If
            End If

        ElseIf card_name = "Genetics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            draw_id = draw_and_meld(player, 10)
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, color(draw_id)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = size3(board, player, color(draw_id)) - 1 To 1 Step -1
                Call score(player, board(player, color(draw_id), 1))
                Call DeleteArrayItem3(board, player, color(draw_id), 1)
            Next i
            Call update_display()

        ElseIf card_name = "Satellites" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                If player = 0 And ai_mode <> 1 Then
                    If hand(player, 0) = -1 Then
                        For i = 1 To 3
                            Call draw_num(player, 8)
                        Next i
                    Else
                        Call show_prompt("satellites", dogma(id, level - 1), 0)
                    End If
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        Call return_from_hand(player, 0)
                    Next i
                    For i = 1 To 3
                        Call draw_num(player, 8)
                    Next i
                End If
            ElseIf level = 2 Then
                splay_options(2) = 1
                Call perform_best_splay(player, splay_options, "Up")
            Else
                If hand(player, 0) > 0 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("satellites3", dogma(id, level - 1), 0)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        draw_id = hand(player, size2(hand, player) - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call meld_from_hand(player, size2(hand, player) - 1)
                        Call perform_solo_dogma_effects(player, draw_id)
                    End If
                End If
            End If

        ElseIf card_name = "Services" Then
            If score_pile(player, 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                max = age(score_pile(player, size2(score_pile, player) - 1))
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = size2(score_pile, player) - 1 To 0 Step -1
                    If age(score_pile(player, i)) = max Then
                        Call transfer_score_to_other_hand(player, active_player, i)
                    End If
                Next i
                max_index = -1
                max = -1
                For i = 0 To 4
                    If board(active_player, i, 0) > -1 Then
                        If Not card_has_symbol(board(active_player, i, 0), 1) Then
                            If age(board(active_player, i, 0)) > max Then
                                max_index = i
                                max = age(board(active_player, i, 0))
                            End If
                        End If
                    End If
                Next i
                If max_index > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("services", dogma(id, level - 1), 0)
                    Else
                        'MsgBox active_player + 1 & " to " & player + 1 & " index=" & max_index
                        Call transfer_board_to_hand(active_player, player, max_index)
                    End If
                End If
            End If

        ElseIf card_name = "Specialization" Then
            If level = 1 Then
                If hand(player, 0) > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("specialization", dogma(id, level - 1), 0)
                    Else
                        For i = 0 To 4
                            already_returned(i) = 0
                        Next i
                        max = 0
                        max_index = -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = 0 To size2(hand, player) - 1
                            found = color(hand(player, i))
                            If already_returned(found) = 0 Then
                                already_returned(found) = 1
                                Call save(player)
                                For j = 0 To num_players - 1
                                    If j <> player And board(j, found, 0) > -1 Then
                                        Call transfer_board_to_hand(j, player, found)
                                    End If
                                Next j
                                Call restore(player)
                                'MsgBox "Considering " & color_lookup(found) & " for player " & player + 1 & " score=" & current_score

                                If current_score > max Then
                                    'MsgBox "Choosing " & color_lookup(found) & " for player " & player + 1
                                    max = current_score
                                    max_index = i
                                End If
                            End If
                        Next i
                        found = color(hand(player, max_index))
                        Call append(player, " reveals a " & color_lookup(found) & " card")
                        For j = 0 To num_players - 1
                            If j <> player And board(j, found, 0) > -1 Then
                                Call transfer_board_to_hand(j, player, found)
                            End If
                        Next j
                    End If
                End If
            Else
                splay_options(0) = 1
                splay_options(3) = 1
                Call perform_best_splay(player, splay_options, "Up")
            End If

        ElseIf card_name = "Suburbia" Then
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    human_data = 0
                    Call show_prompt("suburbia", dogma(id, level - 1), 1)
                Else
                    ' Try to tuck everything
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    count = size2(hand, player)
                    Call save(player)
                    For i = count - 1 To 0 Step -1
                        Call tuck_from_hand(player, i)
                    Next i
                    For i = 1 To count
                        Call draw_and_score(player, 1)
                    Next i
                    Call restore(player)
                    If current_score > max_score Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        For i = count - 1 To 0 Step -1
                            Call tuck_from_hand(player, i)
                        Next i
                        For i = 1 To count
                            Call draw_and_score(player, 1)
                        Next i
                    End If

                End If
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' Age 10
            ''''''''''''''''''''''''''''''''''''''''''''''''
#If VERBOSE Then
            Call append(player, "In activate_card() AGE 10")
#End If
        ElseIf card_name = "A.I." Then
            If level = 1 Then
                Call draw_and_score(player, 10)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
            Else
                found = 0
                For i = 0 To num_players - 1
                    For j = 0 To 4
                        If board(i, j, 0) > -1 Then
                            If title(board(i, j, 0)) = "Robotics" Then found = 1
                        End If
                    Next j
                Next i
                For i = 0 To num_players - 1
                    For j = 0 To 4
                        If board(i, j, 0) > -1 Then
                            If title(board(i, j, 0)) = "Software" And found = 1 Then
                                Call end_game_ai()
                            End If
                        End If
                    Next j
                Next i
            End If

        ElseIf card_name = "Bioengineering" Then
            If level = 1 Then
                max = -1
                max_index = -1
                For i = 0 To num_players - 1
                    For j = 0 To 4
                        If board(i, j, 0) > -1 And i <> player Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(i, j, 0), 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If card_has_symbol(board(i, j, 0), 1) Then
                                If age(board(i, j, 0)) > max Then
                                    max_index = i
                                    max_index2 = j
                                    max = age(board(i, j, 0))
                                End If
                            End If
                        End If
                    Next j
                Next i
                If max_index > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("bioengineering", dogma(id, level - 1), 0)
                    Else
                        'MsgBox active_player + 1 & " to " & player + 1 & " index=" & max_index
                        Call transfer_board_to_score(max_index, player, max_index2)
                    End If
                End If
            Else
                max_count = 0
                max = 0
                max_index = -1
                found = 0
                For i = 0 To num_players - 1
                    If icon_total(i, 1) < 3 Then found = 1
                    If icon_total(i, 1) > max Then
                        max = icon_total(i, 1)
                        max_index = i
                        max_count = 1
                    ElseIf icon_total(i, 1) = max Then
                        max_count = max_count + 1
                    End If
                Next i
                If found = 1 And max_count = 1 Then
                    Call end_game_generic(max_index, "Player " & max_index + 1 & " has the most leaf symbols and wins the game via Bioengineering.")
                End If
            End If

        ElseIf card_name = "Databases" Then
            If score_pile(player, 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                human_data = Int((1 + size2(score_pile, player)) / 2)
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("databases", dogma(id, level - 1), 0)
                Else
                    For i = 1 To human_data
                        Call return_from_score_pile(player, 0)
                    Next i
                End If
            End If

        ElseIf card_name = "Globalization" Then
            If level = 1 Then
                min_index = -1
                min = 11
                For j = 0 To 4
                    If board(player, j, 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, j, 0), 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If card_has_symbol(board(player, j, 0), 1) Then
                            If age(board(player, j, 0)) < min Then
                                min = age(board(player, j, 0))
                                min_index = j
                            End If
                        End If
                    End If
                Next j
                If min_index > -1 Then
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("globalization", dogma(id, level - 1), 0)
                    Else
                        Call return_from_board(player, min_index)
                    End If
                End If
            Else
                Call draw_and_score(player, 6)
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1

                max_count = 0
                max = 0
                max_index = -1
                found = 0
                For i = 0 To num_players - 1
                    If icon_total(i, 1) > icon_total(i, 5) Then
                        found = 1
                        'MsgBox "found it for " & i + 1 & " " & icon_total(i, 1) & " >" & icon_total(i, 5)

                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If scores(i) > max Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        max = scores(i)
                        max_index = i
                        max_count = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf scores(i) = max Then
                        max_count = max_count + 1
                    End If
                Next i
                If found = 0 And max_count = 1 Then
                    Call end_game_generic(max_index, "Player " & max_index + 1 & " has the most points and wins the game via Globalization.")
                End If
            End If

        ElseIf card_name = "Miniaturization" Then
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("miniaturization", dogma(id, level - 1), 1)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If age(hand(player, size2(hand, player) - 1)) = 10 And score_pile(player, 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        Call return_from_hand(player, min_index)
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = 0 To size2(score_pile, player) - 1
                            If already_returned(age(score_pile(player, i))) = 0 Then
                                already_returned(age(score_pile(player, i))) = 1
                                Call draw_num(player, 10)
                            End If
                        Next i
                    End If
                End If
            End If

        ElseIf card_name = "Robotics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            If board(player, 4, 0) > -1 Then Call score_top_card(player, 4)
            'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            draw_id = draw_and_meld(player, 10)
            'MsgBox "about to perform " & title(draw_id) & " for " & player + 1
            Call perform_solo_dogma_effects(player, draw_id)

        ElseIf card_name = "Self Service" Then
            If level = 1 Then
                ' Find the max top card to activate
                max_index = -1
                max = -1
                For i = 0 To 3
                    If board(player, i, 0) <> -1 Then
                        If age(board(player, i, 0)) > max Then
                            max = age(board(player, i, 0))
                            max_index = i
                        End If
                    End If
                Next i
                If max_index > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If active_player <> player Then dogma_copied = 1
                    If player = 0 And ai_mode <> 1 Then
                        Call show_prompt("selfService", dogma(id, level - 1), 0)
                    Else
                        Call perform_solo_dogma_effects(player, board(player, max_index, 0))
                    End If
                End If
            Else
                max_count = 0
                max = 0
                max_index = -1
                For i = 0 To num_players - 1
                    If vps(i) > max Then
                        max = vps(i)
                        max_index = i
                        max_count = 1
                    ElseIf vps(i) = max Then
                        max_count = max_count + 1
                    End If
                Next i
                If max_count = 1 Then
                    Call end_game_generic(max_index, "Player " & max_index + 1 & " has the most achievements and wins the game via Self Service.")
                End If
            End If

        ElseIf card_name = "Software" Then
            If level = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call draw_and_score(player, 10)
            Else
                Call draw_and_meld(player, 10)
                'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                draw_id = draw_and_meld(player, 10)
                Call perform_solo_dogma_effects(player, draw_id)
            End If

        ElseIf card_name = "Stem Cells" Then
            If hand(player, 0) > -1 Then
                If player = 0 And ai_mode <> 1 Then
                    Call show_prompt("stemCells", dogma(id, level - 1), 2)
                Else
                    Call save(player)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For i = size2(hand, player) - 1 To 0 Step -1
                        Call score_from_hand(player, 0)
                    Next i
                    Call restore(player)
                    If current_score > original_score Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If active_player <> player Then dogma_copied = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For i = size2(hand, player) - 1 To 0 Step -1
                            Call score_from_hand(player, 0)
                        Next i
                    End If
                End If
            End If

        ElseIf card_name = "The Internet" Then
            If level = 1 Then
                splay_options(4) = 1
                Call perform_best_splay(player, splay_options, "Up")
            ElseIf level = 2 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call draw_and_score(player, 10)
            Else
                found = Int(icon_total(player, 6) / 2)
                For i = 1 To found
                    Call draw_and_meld(player, 10)
                Next i
            End If
        End If

    End Sub

    Private Sub show_prompt(ByRef which_phase As String, ByRef str_Renamed As String, ByRef allow_cancel As Short)
        If ai_mode = 1 Then MsgBox("Shouldn't be showing this!")
        'MsgBox "Showing " & which_phase
        phase = which_phase
        cmdNext.Visible = False
        If allow_cancel = 0 Then
            cmdCancel.Visible = False
            cmdYes.Visible = False
        ElseIf allow_cancel = 1 Then
            cmdCancel.Visible = True
            cmdYes.Visible = False
        Else
            cmdCancel.Visible = True
            cmdYes.Visible = True
        End If

        lblPrompt.Text = str_Renamed
        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        break_dogma_loop = 1
        Call disable_actions()

    End Sub

    Public Sub draw(ByVal player As Short)
        ' A player can only draw a card of their highest value
        'If ai_mode = 1 Then MsgBox "finding next draw"
        'If ai_mode = 1 Then MsgBox "finding next draw = " & find_next_draw(player)
        'UPGRADE_WARNING: Couldn't resolve default property of object find_next_draw(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call draw_num(player, find_next_draw(player))
        'If ai_mode = 1 Then MsgBox "card drawn"
    End Sub

    Private Function find_next_draw(ByVal player As Short) As Object
        Dim num As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object highest_top_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        num = highest_top_card(player)
        If num = 0 Then num = 1
        num = num - 1
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(deck, num). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        While (size2(deck, num) = 0)
            num = num + 1
            If num > 9 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object find_next_draw. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                find_next_draw = num + 1
                Exit Function
            End If
        End While
        'UPGRADE_WARNING: Couldn't resolve default property of object find_next_draw. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        find_next_draw = num + 1
    End Function

    Private Function find_card(ByVal player As Short, ByVal id As Short) As Object
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            If hand(player, i) = id Then
                'UPGRADE_WARNING: Couldn't resolve default property of object find_card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                find_card = i
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object find_card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        find_card = -1
    End Function

    Private Sub draw_and_score(ByVal player As Short, ByVal num As Short)
        Dim id As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        silent = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        id = draw_num(player, num)
        'MsgBox "Found " & id & " at " & find_card(player, id)

        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call remove_card_from_hand(player, find_card(player, id))
        Call score(player, id)
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        silent = 0
        Call append(player, "drew and scored a " & age(id))
    End Sub

    Private Function draw_and_meld(ByVal player As Short, ByVal num As Short) As Object
        Dim id As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        silent = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        id = draw_num(player, num)
        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call remove_card_from_hand(player, find_card(player, id))
        Call meld(player, id)
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        silent = 0
        Call append(player, "drew and melded " & title(id))
        'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_meld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        draw_and_meld = id
    End Function

    Private Function draw_and_tuck(ByVal player As Short, ByVal num As Short) As Object
        Dim id As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        silent = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object draw_num(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        id = draw_num(player, num)
        'UPGRADE_WARNING: Couldn't resolve default property of object find_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call remove_card_from_hand(player, find_card(player, id))
        Call tuck(player, id)
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        silent = 0
        Call append(player, "drew and tucked " & title(id))
        'UPGRADE_WARNING: Couldn't resolve default property of object draw_and_tuck. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        draw_and_tuck = id
    End Function

    Private Sub score_from_hand(ByVal player As Short, ByVal index As Short)
        Dim id As Short
        id = hand(player, index)
        Call remove_card_from_hand(player, index)
        Call score(player, id)
        Call update_display()
    End Sub

    Private Sub score(ByVal player As Short, ByVal id As Short)
#If VERBOSE Then ' too much info
        Call append_simple("score() -----------------------------")
#End If
        Call append(player, "scored a " & age(id))
        Call Push2(score_pile, player, id)
        Call update_display()
        'UPGRADE_WARNING: Couldn't resolve default property of object scored_this_turn(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object scored_this_turn(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        scored_this_turn(player) = scored_this_turn(player) + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object scored_this_turn(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If scored_this_turn(player) > 5 Then
            Call achieve_special(player, 9, "Monument")
        End If
    End Sub

    Private Function draw_num(ByVal player As Short, ByVal num As Short) As Object
#If VERBOSE Then ' too much info
        Call append_simple("draw_num() -----------------------------")
#End If
        Dim picked As Object
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object picked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        picked = 11

        ' What's the next available pile with cards?
        num = num - 1
        For i = num To 9
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(deck, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(deck, i) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object picked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                picked = i
                'MsgBox "value found"
                i = 10
            End If
        Next i

        'UPGRADE_WARNING: Couldn't resolve default property of object picked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If picked = 11 Then
            ' End the game!!

            'MsgBox "what happened?"
            If phase <> "new_game" Then
                'MsgBox ("pile is empty: " & num + 1)
                Call end_game_10(player)
            End If
            ' this is a hack to stop the game from crashing
            'UPGRADE_WARNING: Couldn't resolve default property of object draw_num. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            draw_num = 0
            Call Push2(hand, player, 0)
            Exit Function
        End If

        'MsgBox draw_num & " came from " & num & " WHEN PICKING " & picked
        'MsgBox deck(2, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object picked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        num = picked
        draw_num = deck(num, 0)


        Call Push2(hand, player, deck(num, 0))
        Call Shift2(deck, num)
        'UPGRADE_WARNING: Couldn't resolve default property of object draw_num. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call append(player, "drew a " & age(draw_num))
        Call update_display()
    End Function

    Public Sub set_icon_image(ByRef image As Object, ByRef id As Object, ByRef icon_id As Object)
#If VERBOSE2 Then ' too much info
        Call append_simple("In set_icon_image() -----------------------------")
#End If
        If image.Tag = Nothing Then
            image.Tag = -1
        End If
        If ai_mode <> 0 Then Exit Sub
        'UPGRADE_WARNING: Couldn't resolve default property of object image.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If image.Tag <> id Then
            'UPGRADE_WARNING: Couldn't resolve default property of object icon_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If icon_id = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object image.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                image.Visible = False
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object image.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                image.Visible = True
                'UPGRADE_WARNING: Couldn't resolve default property of object image.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object icon_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                image.Image = icon_images(icon_id)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object image.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            image.Tag = id
        End If
    End Sub

    Public Sub set_color_image(ByRef image As Object, ByRef id As Object, ByRef color_id As Object)
#If VERBOSE2 Then ' too much info?
        Call append_simple("In set_color_image() -----------------------------")
#End If
        If ai_mode <> 0 Then Exit Sub
        'UPGRADE_WARNING: Couldn't resolve default property of object image.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If image.Tag <> id Then
            'UPGRADE_WARNING: Couldn't resolve default property of object image.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object color_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            image.Image = color_images(color_id)
            'UPGRADE_WARNING: Couldn't resolve default property of object image.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            image.Tag = id
        End If
    End Sub

    Public Function highest_top_card(ByVal player As Short) As Object
        Dim i, max As Short
        max = 0
        For i = 0 To 4
            If board(player, i, 0) > -1 Then
                If age(board(player, i, 0)) > max Then max = age(board(player, i, 0))
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object highest_top_card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        highest_top_card = max
    End Function

    Public Sub update_scores()
#If VERBOSE2 Then ' too much info
        Call append_simple("In update_scores() -----------------------------")
#End If
        Dim i, j As Short
        For i = 0 To 3 ' player
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            scores(i) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For j = 0 To size2(score_pile, i) - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                scores(i) = scores(i) + age(score_pile(i, j))
            Next j
        Next i
        'MsgBox "score for player 3: " & scores(3)
    End Sub

    Public Sub update_icon_total()
#If VERBOSE2 Then ' too much info
        Call append_simple("In update_icon_total() -----------------------------")
#End If

        Dim l, j, i, k, m, board_id As Short

        ' Update the icons based on the top cards
        For i = 0 To 3 ' player
            For j = 1 To 6 ' icon type
                icon_total(i, j) = 0
                For k = 0 To COLORCOUNT ' top cards
                    board_id = board(i, k, 0)
                    If board_id > -1 Then
                        For l = 0 To 3 ' icons on a card
                            If icons(board_id, l) = j Then
                                icon_total(i, j) = icon_total(i, j) + 1
                            End If
                        Next l

                        ' Now iterate through splays
                        'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, i, k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        For m = 1 To size3(board, i, k) - 1
                            If splayed(i, k) = "Up" And icons(board(i, k, m), 1) = j Then icon_total(i, j) = icon_total(i, j) + 1
                            If splayed(i, k) = "Up" And icons(board(i, k, m), 2) = j Then icon_total(i, j) = icon_total(i, j) + 1
                            If splayed(i, k) = "Up" And icons(board(i, k, m), 3) = j Then icon_total(i, j) = icon_total(i, j) + 1
                            If splayed(i, k) = "Down" And icons(board(i, k, m), 0) = j Then icon_total(i, j) = icon_total(i, j) + 1
                            If splayed(i, k) = "Right" And icons(board(i, k, m), 0) = j Then icon_total(i, j) = icon_total(i, j) + 1
                            If splayed(i, k) = "Right" And icons(board(i, k, m), 1) = j Then icon_total(i, j) = icon_total(i, j) + 1
                            If splayed(i, k) = "Left" And icons(board(i, k, m), 3) = j Then icon_total(i, j) = icon_total(i, j) + 1
                        Next m
                    End If
                Next k
            Next j
        Next i

        ' If a player has 12 more clocks, they claim 'World'
        Dim scored_it As Object
        Dim player As Short
        If achievements_scored(12) = 0 Then
            For i = active_player To (active_player + num_players) - 1 ' player
                If icon_total(player, 6) >= 12 Then
                    achievements_scored(12) = 1
                    vps(player) = vps(player) + 1
                    alert(("Player " & player + 1 & " claimed the 'World' achievement."))
                    Call append(player, "claimed the ##### 'World' achievement #####")
                    Call update_icon_total()
                    Call end_game_points()
                    Exit Sub
                End If
            Next i
        End If

        ' If a player has 3 more of all 6 types, they claim 'Empire'
        If achievements_scored(10) = 0 Then
            For i = active_player To (active_player + num_players) - 1 ' player
                'UPGRADE_WARNING: Couldn't resolve default property of object scored_it. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                scored_it = 1
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                player = i Mod num_players
                For j = 1 To 6 ' icon type
                    'UPGRADE_WARNING: Couldn't resolve default property of object scored_it. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If icon_total(player, j) < 3 Then scored_it = 0
                Next j
                'UPGRADE_WARNING: Couldn't resolve default property of object scored_it. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If scored_it = 1 Then
                    achievements_scored(10) = 1
                    vps(player) = vps(player) + 1
                    alert(("Player " & player + 1 & " claimed the 'Empire' achievement."))
                    Call append(player, "claimed the ##### 'Empire' achievement ####")
                    Call update_icon_total()
                    Call end_game_points()
                    Exit Sub
                End If
            Next i
        End If

    End Sub

    Private Sub check_for_achievements()
#If VERBOSE Then ' too much info
        Call append_simple("In check_for_achievements() -----------------------------")
#End If
        ' Dim player, wonder_ok As Object
        Dim i, j, player, universe_ok, wonder_ok As Short

        ' If a player has 5 colors, each splayed up or right, they claim Wonder
        ' If a player has 5 colors with values >= 8, the score Universe
        If achievements_scored(11) = 0 Or achievements_scored(13) = 0 Then
            For i = active_player To (active_player + num_players) - 1 ' player
                wonder_ok = 1 - achievements_scored(11)
                universe_ok = 1 - achievements_scored(13)
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                player = i Mod num_players
                For j = 0 To COLORCOUNT
                    'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If size3(board, player, j) = 0 Then
                        wonder_ok = 0
                        universe_ok = 0
                    Else
                        If age(board(player, j, 0)) < 8 Then universe_ok = 0
                        If splayed(player, j) <> "Up" And splayed(player, j) <> "Right" Then
                            wonder_ok = 0
                            'MsgBox "wonder failed"
                        End If
                    End If
                Next j
                If universe_ok = 1 Then
                    achievements_scored(13) = 1
                    vps(player) = vps(player) + 1
                    alert(("Player " & player + 1 & " claimed the 'Universe' achievement."))
                    Call append(player, "claimed the #### 'Universe' achievement ####")
                    Call end_game_points()
                    Call check_for_achievements()
                    Exit Sub
                End If

                If wonder_ok = 1 Then
                    achievements_scored(11) = 1
                    vps(player) = vps(player) + 1
                    alert(("Player " & player + 1 & " claimed the 'Wonder' achievement."))
                    Call append(player, "claimed the #### 'Wonder' achievement ####")
                    Call end_game_points()
                    Call check_for_achievements()
                    Exit Sub
                End If
            Next i
        End If

    End Sub

    Private Sub do_alert(ByVal id As Short)
        'UPGRADE_WARNING: Couldn't resolve default property of object found(id). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If found(id) = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object found(id). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            found(id) = 1
        ElseIf id = 0 Then
            alert_found = 1
            alert("duplicate card found: " & title(id))
        End If
    End Sub

    Public Sub update_display()
#If VERBOSE2 Then ' too much info
        Call append_simple("In update_display() -----------------------------")
#End If
        Call update_scores()
        Call update_icon_total()

        If ai_mode = 1 Then Exit Sub

        ' Update the number of cards in each pile
        'Dim id As Object
        Dim which, id, k, i, j As Short
        Dim max As Short

        ' Check to make sure that the decks have the right cards in them
        If alert_found = 0 And ai_mode = 0 Then
            For i = 0 To AGECOUNT
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(deck, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = 0 To size2(deck, i) - 1
                    If age(deck(i, j)) <> i + 1 Then
                        alert("bad card found in deck " & i + 1)
                        alert_found = 1
                    End If
                Next j
            Next i


            For i = 0 To num_cards - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object found(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                found(i) = 0
            Next i
            ' Deck
            For i = 0 To AGECOUNT
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(deck, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = 0 To size2(deck, i) - 1
                    Call do_alert(deck(i, j))
                Next j
            Next i
            ' Hand
            For i = 0 To num_players - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = 0 To size2(hand, i) - 1
                    Call do_alert(hand(i, j))
                Next j
            Next i
            ' Score Pile
            For i = 0 To num_players - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = 0 To size2(score_pile, i) - 1
                    Call do_alert(score_pile(i, j))
                Next j
            Next i
            ' Board
            For i = 0 To num_players - 1
                For k = 0 To 4
                    'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, i, k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For j = 0 To size3(board, i, k) - 1
                        Call do_alert(board(i, k, j))
                    Next j
                Next k
            Next i
        End If


        For i = 0 To AGECOUNT
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(deck, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lblAgeCards(i).Text = i + 1 & ") " & size2(deck, i)
        Next i

        ' Unsplay any piles that have 0 or 1 top cards
        For i = 0 To COLORCOUNT
            For j = 0 To num_players - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, j, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size3(board, j, i) < 2 Then splayed(j, i) = ""
            Next j
        Next i

        ' Update achievement list
        For i = 0 To TOTACHIEVE
            lblAchievement(i).Visible = False
            If achievements_scored(i) = 0 Then lblAchievement(i).Visible = True
        Next i

        ' Update your hand
        Call sort2(hand, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For j = 0 To size2(hand, 0) - 1
            lblHand(j).Text = age(hand(0, j)) & "-" & title(hand(0, j))
            lblHand(j).Tag = hand(0, j)
            lblHand(j).Visible = True

            Call set_icon_image(imgHandIcon(j), hand(0, j), dogma_icon(hand(0, j)))
            Call set_color_image(imgHandColor(j), hand(0, j), color(hand(0, j)))
            lblHand(j).BackColor = background_colors(color_lookup(color(hand(0, j))))
            imgHandIcon(j).Visible = True
            imgHandColor(j).Visible = True
        Next j
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For j = size2(hand, 0) To 39
            lblHand(j).Visible = False
            imgHandIcon(j).Visible = False
            imgHandColor(j).Visible = False
        Next j

        ' Update your score pile
        Call sort2(score_pile, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For j = 0 To size2(score_pile, 0) - 1
            lblScoreTitle(j).Text = age(score_pile(0, j)) & "-" & title(score_pile(0, j))
            lblScoreTitle(j).Tag = score_pile(0, j)
            lblScoreTitle(j).Visible = True

            Call set_icon_image(imgScoreIcon(j), score_pile(0, j), dogma_icon(score_pile(0, j)))
            Call set_color_image(imgScoreColor(j), score_pile(0, j), color(score_pile(0, j)))
            lblScoreTitle(j).BackColor = background_colors(color_lookup(color(score_pile(0, j))))
            imgScoreIcon(j).Visible = True
            imgScoreColor(j).Visible = True
        Next j
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For j = size2(score_pile, 0) To 39
            lblScoreTitle(j).Visible = False
            imgScoreIcon(j).Visible = False
            imgScoreColor(j).Visible = False
        Next j
        'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lblScorePile.Text = "Your Score Pile (" & scores(0) & " points)"

        ' Update each opponent's hand
        For i = 0 To num_players - 2
            '    If current_turn > 1 Then MsgBox "checking hand for " & i
            Call sort2(hand, i + 1)
            lblOppHand(i).Text = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For j = 0 To size2(hand, i + 1) - 1
                lblOppHand(i).Text = lblOppHand(i).Text & age(hand(i + 1, j)) & " "
            Next j

            ' Update Score pile
            Call sort2(score_pile, i + 1)
            lblOppScore(i).Text = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, i + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For j = 0 To size2(score_pile, i + 1) - 1
                lblOppScore(i).Text = lblOppScore(i).Text & age(score_pile(i + 1, j)) & " "
            Next j
        Next i
        ' If current_turn > 1 Then MsgBox "check 1"

        ' Update the icons based on the top cards
        For j = 1 To ICONCOUNT ' icon type 
            'max = 0
            'For i = 0 To num_players - 1
            '    If icon_total(i, j) > max Then max = icon_total(i, j)
            'Next i
            For i = 0 To num_players - 1
                id = i + (j - 1) * 4
                lblIcon(id).ForeColor = System.Drawing.Color.Black
                lblIcon(id).Font = VB6.FontChangeBold(lblIcon(id).Font, False)
                lblIcon(id).Text = CStr(icon_total(i, j))
                If icon_total(i, j) > icon_total(0, j) Then
                    lblIcon(id).ForeColor = System.Drawing.Color.Red
                    lblIcon(id).Font = VB6.FontChangeBold(lblIcon(id).Font, True)
                ElseIf icon_total(i, j) < icon_total(0, j) Then
                    lblIcon(id).ForeColor = System.Drawing.ColorTranslator.FromOle(&H8000)
                    lblIcon(id).Font = VB6.FontChangeBold(lblIcon(id).Font, True)
                End If
            Next i
        Next j

        ' Update Scores and VPs
        max = 0
        For i = 0 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If scores(i) > max Then max = scores(i)
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lblScore(i).Text = scores(i)
        Next i
        'For i = 0 To num_players - 1
        '    lblScore(i).ForeColor = vbBlack
        '    lblScore(i).FontBold = False
        '    If scores(i) = max And max > 0 Then
        '        lblScore(i).ForeColor = vbRed
        '        lblScore(i).FontBold = True
        '    End If
        'Next i

        max = 0
        For i = 0 To num_players - 1
            If vps(i) > max Then max = vps(i)
            lblVP(i).Text = CStr(vps(i))
        Next i
        For i = 0 To num_players - 1
            lblVP(i).ForeColor = System.Drawing.Color.Black
            lblVP(i).Font = VB6.FontChangeBold(lblVP(i).Font, False)
            If vps(i) = max And max > 0 Then
                lblVP(i).ForeColor = System.Drawing.Color.Red
                lblVP(i).Font = VB6.FontChangeBold(lblVP(i).Font, True)
            End If
        Next i

        ' Update your board
        'UPGRADE_WARNING: Couldn't resolve default property of object highest_top_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lblHighestTop.Text = Str(highest_top_card(0))
        For i = 0 To COLORCOUNT
            id = board(0, i, 0)
            If id > -1 Then
                lblBoardTitle(i).Visible = True
                lblBoardTitle(i).Text = title(id)
                lblBoardTitle(i).Tag = id

                lblBoardAge(i).Visible = True
                lblBoardAge(i).Text = CStr(age(id))
                lblBoardAge(i).Tag = id

                lblDogma(i).Visible = True
                lblDogma(i).Tag = id

                lblBoardDetails(i).Visible = True
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, 0, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblBoardDetails(i).Text = "Size: " & size3(board, 0, i)
                If splayed(0, i) <> "" Then
                    lblBoardDetails(i).Text = lblBoardDetails(i).Text & "  Splay: " & splayed(0, i)
                End If
                'cmdDogma(i).Visible = True
                cmdStack(i).Visible = True

                imgBoard(i).Visible = True
                Call set_color_image(imgBoard(i), id, color(id))
                imgBoard(i).Tag = id

                Call set_icon_image(imgBoardDogma(i), id, dogma_icon(id))
                For j = 0 To 3
                    If icons(id, j) > 0 Then
                        imgIconSmall(4 * i + j).Visible = True
                        Call set_icon_image(imgIconSmall(4 * i + j), id, icons(id, j))
                    Else
                        imgIconSmall(4 * i + j).Visible = False
                    End If
                Next j
            Else
                lblBoardTitle(i).Visible = False
                lblBoardAge(i).Visible = False
                lblBoardDetails(i).Visible = False
                lblDogma(i).Visible = False
                cmdDogma(i).Visible = False
                cmdStack(i).Visible = False
                imgBoardDogma(i).Visible = False
                imgBoardDogma(i).Tag = -1
                imgBoard(i).Visible = False
                For j = 0 To 3
                    ' MsgBox ("Hiding " & i & " + " & j)
                    imgIconSmall(4 * i + j).Visible = False
                Next j
            End If
        Next i

        ' Update Opponent's Board
        For i = 0 To num_players - 2
            For j = 0 To COLORCOUNT
                id = board(i + 1, j, 0)
                which = 5 * i + j
                If id > -1 Then
                    lblOppBoard(which).Text = age(id) & "-" & title(id)
                    lblOppBoard(which).Tag = id
                    lblOppBoard(which).Visible = True
                    lblOppDetail(which).Visible = True
                    lblOppDetail(which).Text = "Size: " & size3(board, i + 1, j)
                    If splayed(i + 1, j) <> "" Then
                        lblOppDetail(which).Text = lblOppDetail(which).Text & " Splay: " & splayed(i + 1, j)
                    End If

                    Call set_icon_image(imgOppBoard(which), id, dogma_icon(id))
                    Call set_color_image(imgOppBoardColor(which), id, color(id))
                    imgOppBoard(which).Visible = True
                    imgOppBoardColor(which).Visible = True
                Else
                    lblOppBoard(which).Visible = False
                    lblOppDetail(which).Visible = False
                    imgOppBoard(which).Visible = False
                    imgOppBoardColor(which).Visible = False
                End If
            Next j
        Next i

    End Sub


    Private Sub initialize_card_data()

#If VERBOSE Then
        Call append_simple("In initialize_card_data() -----------------------------") 'This isnt echoping... ???
#End If
        color_lookup(0) = "Yellow"
        color_lookup(1) = "Red"
        color_lookup(2) = "Purple"
        color_lookup(3) = "Blue"
        color_lookup(4) = "Green"

        icon_lookup(0) = "x"
        icon_lookup(1) = "Leaf"
        icon_lookup(2) = "Castle"
        icon_lookup(3) = "Lightbulb"
        icon_lookup(4) = "Crown"
        icon_lookup(5) = "Factory"
        icon_lookup(6) = "Clock"

        Dim fileData, lineData As String
        Dim word As String
        Dim i, k, line_index, word_index, j As Short

        Dim filename As String = ".\Innovation.txt" ' Holds all the card data

        If System.IO.File.Exists(filename) = False Then
            'MessageBox.Show("Exiting Fatal Error. Missing data file: " & filename)
            MsgBox("Exiting Fatal Error. Missing data file: " & filename, vbCritical)
            End
        End If
#If VERBOSE Then
        Call append_simple("In initialize_card_data() found file .\Innovation.txt -------") 'This isnt echoing... ???
#End If
        FileOpen(1, My.Application.Info.DirectoryPath & filename, OpenMode.Binary)
        fileData = InputString(1, LOF(1))
        FileClose(1)
        line_index = InStr(fileData, Chr(13))
        fileData = Mid(fileData, line_index + 1)
        line_index = InStr(fileData, Chr(13))
        fileData = Mid(fileData, line_index + 1)
        line_index = InStr(fileData, Chr(13))
        i = 0

        While (line_index > 0)
            lineData = Mid(fileData, 1, line_index)

            ' Skip the 1st column
            word_index = InStr(lineData, Chr(9))
            lineData = Mid(lineData, word_index + 1)

            ' Load age
            word_index = InStr(lineData, Chr(9))
            word = Mid(lineData, 1, word_index)
            lineData = Mid(lineData, word_index + 1)
            age(i) = Int(CDbl(word))

            ' Load color
            word_index = InStr(lineData, Chr(9))
            word = Mid(lineData, 1, word_index - 1)
            lineData = Mid(lineData, word_index + 1)
            For j = 0 To COLORCOUNT
                If color_lookup(j) = word Then
                    color(i) = j
                End If
            Next
            'MsgBox(color_lookup(color(i)))
            ' FK Not worth adding as a debug option

            ' Load title
            word_index = InStr(lineData, Chr(9))
            word = Mid(lineData, 1, word_index - 1)
            lineData = Mid(lineData, word_index + 1)
            title(i) = RTrim(word)

            ' Load icons
            For k = 0 To 3
                word_index = InStr(lineData, Chr(9))
                word = Mid(lineData, 1, word_index - 1)
                lineData = Mid(lineData, word_index + 1)
                For j = 0 To 6
                    If icon_lookup(j) = word Then
                        icons(i, k) = j
                    End If
                Next
            Next

            ' Load dogma icon
            word_index = InStr(lineData, Chr(9))
            lineData = Mid(lineData, word_index + 1)
            word_index = InStr(lineData, Chr(9))
            word = Mid(lineData, 1, word_index - 1)
            lineData = Mid(lineData, word_index + 1)
            For j = 0 To 6
                If icon_lookup(j) = word Then
                    dogma_icon(i) = j
                End If
            Next

            ' Load dogma conditions
            For k = 0 To 2
                word_index = InStr(lineData, Chr(9))
                If word_index < 1 Then word_index = InStr(lineData, Chr(13))
                word = Mid(lineData, 1, word_index - 1)
                lineData = Mid(lineData, word_index + 1)
                If Mid(word, 1, 1) = Chr(34) Then word = Mid(word, 2)
                If VB.Right(word, 1) = Chr(34) Then word = Mid(word, 1, Len(word) - 1)
                dogma(i, k) = word
                If Mid(word, 1, 8) = "I demand" Then
                    is_demand(i, k) = 1
                End If
                'MsgBox(word)
                ' FK Not worth adding as a debug option
            Next

            ' Go to the next line
            fileData = Mid(fileData, line_index + 1)
            line_index = InStr(fileData, Chr(13))
            i = i + 1
        End While

        num_cards = i
        'MsgBox(Str(i) + " cards loaded")
#If VERBOSE Then
        Call append_simple("In initial_card_data() 105 cards loaded -------") 'FK This isnt echoing... ???
#End If
        'When functioning 105 are loaded

        cmbCards.Items.Clear()
        For i = 0 To num_cards - 1
            cmbCards.Items.Add((title(i)))
        Next i

        cmbCheatLevel.Items.Clear()
        For i = 0 To 5
            cmbCheatLevel.Items.Add((20 * i & "%"))
        Next i
        cmbCheatLevel.SelectedIndex = 5

    End Sub

    Private Sub initialize_images()
#If VERBOSE Then
        Call append_simple("In initialize_images() -----------------------------") 'FK This isnt echoing... ???
#End If
        title_image = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/title.jpg")
        achievement_images(0) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve1.jpg")
        achievement_images(1) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve2.jpg")
        achievement_images(2) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve3.jpg")
        achievement_images(3) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve4.jpg")
        achievement_images(4) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve5.jpg")
        achievement_images(5) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve6.jpg")
        achievement_images(6) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve7.jpg")
        achievement_images(7) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve8.jpg")
        achievement_images(8) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve9.jpg")
        'although there is an age 10 there isn't an achievement
        'achievement_images(9) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/achieve10.jpg")
        achievement_images(9) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/monumentachievement.jpg")
        achievement_images(10) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/empireachievement.jpg")
        achievement_images(11) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/wonderachievement.jpg")
        achievement_images(12) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/worldachievement.jpg")
        achievement_images(13) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/universeachievement.jpg")

        ' load all the image resources
        Dim column_offset, i, j As Short
        For i = 0 To COLORCOUNT
            ' MsgBox App.Path & "/images/" & color_lookup(i) & ".jpg"
            color_images(i) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/" & color_lookup(i) & ".jpg")
            color_images(i).Tag = -1
        Next i
        For i = 0 To ICONCOUNT
            icon_images(i) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "/images/" & icon_lookup(i) & ".jpg")
            icon_images(i).Tag = -1
        Next i

        ' initialize the hand images
        For i = 0 To 40
            If i > imgHandIcon.UBound Then
                lblHand.Load(i)
                imgHandIcon.Load(i)
                imgHandColor.Load(i)
                lblScoreTitle.Load(i)
                imgScoreIcon.Load(i)
                imgScoreColor.Load(i)
            End If
            imgHandIcon(i).Visible = False
            imgHandColor(i).Visible = False
            lblHand(i).Visible = False
            imgScoreIcon(i).Visible = False
            imgScoreColor(i).Visible = False
            lblScoreTitle(i).Visible = False

            imgHandIcon(i).SetBounds(imgHandIcon(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgHandIcon(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y) 'you must determine where you want it
            imgHandColor(i).SetBounds(imgHandColor(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgHandColor(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            lblHand(i).SetBounds(lblHand(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblHand(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

            imgScoreIcon(i).SetBounds(imgScoreIcon(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgScoreIcon(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y) 'you must determine where you want it
            imgScoreColor(i).SetBounds(imgScoreColor(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgScoreColor(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            lblScoreTitle(i).SetBounds(lblScoreTitle(0).Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblScoreTitle(0).Top) + 360 * i), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            ' 360 - 252  =
            For j = 23 To 5 Step -6
                column_offset = j / 6
                If i > j And i <= j + 6 Then
                    imgHandColor(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgHandColor(0).Left) + column_offset * (VB6.PixelsToTwipsX(imgHandColor(0).Width) + 100)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgHandColor(0).Top) + (VB6.PixelsToTwipsY(imgHandColor(0).Height) + 108) * (i - j - 1)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                    imgHandIcon(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgHandIcon(0).Left) + column_offset * (VB6.PixelsToTwipsX(imgHandColor(0).Width) + 100)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgHandIcon(0).Top) + (VB6.PixelsToTwipsY(imgHandColor(0).Height) + 108) * (i - j - 1)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                    lblHand(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lblHand(0).Left) + column_offset * (VB6.PixelsToTwipsX(imgHandColor(0).Width) + 100)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblHand(0).Top) + (VB6.PixelsToTwipsY(imgHandColor(0).Height) + 108) * (i - j - 1)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                    imgScoreColor(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgScoreColor(0).Left) + column_offset * (VB6.PixelsToTwipsX(imgScoreColor(0).Width) + 100)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgScoreColor(0).Top) + (VB6.PixelsToTwipsY(imgScoreColor(0).Height) + 108) * (i - j - 1)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                    imgScoreIcon(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgScoreIcon(0).Left) + column_offset * (VB6.PixelsToTwipsX(imgScoreColor(0).Width) + 100)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(imgScoreIcon(0).Top) + (VB6.PixelsToTwipsY(imgScoreColor(0).Height) + 108) * (i - j - 1)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
                    lblScoreTitle(i).SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lblScoreTitle(0).Left) + column_offset * (VB6.PixelsToTwipsX(imgScoreColor(0).Width) + 100)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblScoreTitle(0).Top) + (VB6.PixelsToTwipsY(imgScoreColor(0).Height) + 108) * (i - j - 1)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

                End If
            Next j
            imgHandColor(i).SendToBack()
            imgHandIcon(i).BringToFront()
            imgHandIcon(i).Tag = -1
            imgHandColor(i).Tag = -1
            imgScoreColor(i).SendToBack()
            imgScoreIcon(i).BringToFront()
            imgScoreIcon(i).Tag = -1
            imgScoreColor(i).Tag = -1
        Next i
    End Sub
    '

    ' FK Thexe 2 routines append_simple/append are the main output routines echoing text messages into text window it has
    ' been overloaded to handle logfile and debug messages as well
    '
    Public Sub append_simple(ByVal str_Renamed As String)
        'UPGRADE_WARNING: Couldn't resolve default property of object silent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If ai_mode <> 0 Or silent = 1 Then Exit Sub
        txtLog.Text = txtLog.Text & str_Renamed & Chr(13) & Chr(10)
        txtLog.SelectionStart = Len(txtLog.Text)
        txtLog.Select(txtLog.TextLength - 1, 1) 'select the last chr in the textbox
        txtLog.ScrollToCaret() 'scroll to the selected position
        'Stupidly inefficient for now
        My.Computer.FileSystem.WriteAllText("log.txt", txtLog.Text, False)
    End Sub

    Public Sub append(ByVal player As Short, ByVal str_Renamed As String)
        Call append_simple("Player " & player + 1 & " " & str_Renamed & ".")
    End Sub

    Public Sub alert(ByRef str_Renamed As String)

        If ai_mode = 0 Then MsgBox(str_Renamed)
#If VERBOSE Then ' Too verbose
        Call append(0, "In load_picture()" & str_Renamed)
#End If
    End Sub


    Private Sub load_picture(ByRef id As Object)

#If VERBOSE2 Then ' Too verbose
        Call append_simple("In load_picture()")
#End If
        Dim i As Short
        ' MsgBox "Loading " & id
        If lblRules.Tag <> id Then
            lblRules.Tag = id
            lblRules.Text = dogma(id, 0) & vbNewLine
            lblRules.BackColor = background_colors(color_lookup(color(id)))
            lblLarge.BackColor = background_colors(color_lookup(color(id)))
            lblDogmaSymbol.BackColor = background_colors(color_lookup(color(id)))
            If Len(dogma(id, 1)) > 1 Then lblRules.Text = "1) " & lblRules.Text & "2) " & dogma(id, 1)
            If Len(dogma(id, 2)) > 1 Then lblRules.Text = lblRules.Text & vbNewLine & "3) " & dogma(id, 2)
            lblLarge.Text = title(id) & " - Age " & age(id)
            Call set_icon_image(imgDogmaSymbol, id, dogma_icon(id))
            Call set_color_image(imgLarge, id, color(id))
            For i = 0 To 3
                Call set_icon_image(imgLargeIcon(i), id, icons(id, i))
            Next i
            imgLarge.SendToBack()

        End If
    End Sub

    Public Sub meld_from_hand(ByVal player As Short, ByVal index As Short)
        'If ai_mode = 1 Then MsgBox "Trying to meld " & player & " card " & Index
#If VERBOSE Then
        Call append(0, "In meld_from_hand() ----Trying to meld " & player & " card " & index)
#End If
        Dim id As Short
        id = hand(player, index)
        Call remove_card_from_hand(player, index)
        Call meld(player, id)
    End Sub

    Private Sub meld(ByVal player As Short, ByVal id As Short)
        'MsgBox "melding"
#If VERBOSE Then
        Call append_simple("In meld() melding------------------")
#End If
        Call Unshift3(board, player, color(id), id)
        Call append(player, "melds " & title(id))
        Call update_icon_total()
        Call check_for_achievements()
        Call update_display()
    End Sub

    Private Sub remove_card_from_hand(ByVal player As Short, ByVal index As Short)
#If VERBOSE Then
        Call append_simple("In remove_card_from_hand------------------")
#End If
        Call DeleteArrayItem2(hand, player, index)
        Call update_display()
    End Sub

    Private Sub remove_card_from_board(ByVal player As Short, ByVal index As Short)
#If VERBOSE Then
        Call append_simple("In remove_card_from_board------------------")
#End If
        Call DeleteArrayItem3(board, player, index, 0)
        Call update_display()
    End Sub


    Private Sub remove_card_from_score_pile(ByVal player As Short, ByVal index As Short)
#If VERBOSE Then
        Call append_simple("In remove_card_from_score_pile------------------")
#End If
        Call DeleteArrayItem2(score_pile, player, index)
        Call update_display()
    End Sub

    Private Function winner_found() As Object
        Dim i As Short
        Call update_display()
        For i = 0 To num_players - 1
            If winners(i) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object winner_found. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                winner_found = TriState.True
                Exit Function
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object winner_found. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        winner_found = TriState.False
    End Function


    Private Sub end_game_10(ByRef player As Object)
        If winner_found() Then Exit Sub

        Dim min As Object
        Dim i, winner_count As Short
        Dim winner_str As String
        winner_count = 0
        winner_str = ""

        'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        min = 10 * scores(0) + vps(0)
        For i = 1 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If 10 * scores(i) + vps(i) > min Then min = 10 * scores(i) + vps(i)
        Next i
        For i = 0 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If 10 * scores(i) + vps(i) = min Then
                winners(i) = 1
                If winner_count > 0 Then
                    winner_str = winner_str & " and Player " & i + 1
                Else
                    winner_str = "Player " & i + 1
                End If
                winner_count = winner_count + 1
            End If
        Next i
        If winner_count > 1 Then
            winner_str = winner_str & " win!"
        Else
            winner_str = winner_str & " wins!"
        End If

        If ai_mode = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call MsgBox("Player " & player + 1 & " has tried to draw a 10 when the 10 pile was empty.  " & winner_str)
            Call start_new_game()
        End If
    End Sub

    Private Sub end_game_points()
        If winner_found() Then Exit Sub

        Dim i As Short
        For i = 0 To num_players - 1
            If ((vps(i) >= 4 And num_players = 4) Or (vps(i) >= 5 And num_players = 3) Or (vps(i) >= 6 And num_players = 2)) Then
                winners(i) = 1
                If ai_mode = 0 Then
                    Call MsgBox("Player " & i + 1 & " has earned " & vps(i) & " achievements and won the game!")
                    Call start_new_game()
                End If
            End If
        Next i
    End Sub

    'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub end_game_generic(ByVal player As Short, ByVal str_Renamed As String)
        If winner_found() Then Exit Sub

        winners(player) = 1
        If ai_mode = 0 Then
            Call MsgBox(str_Renamed)
            Call start_new_game()
        End If
    End Sub

    Private Sub end_game_ai()
        If winner_found() Then Exit Sub

        Dim min As Object
        Dim i, winner_count As Short
        Dim winner_str As String = ""
        winner_count = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object scores(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        min = scores(0)
        For i = 1 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If scores(i) < min Then min = scores(i)
        Next i
        For i = 0 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If scores(i) = min Then
                winners(i) = 1
                If winner_count > 0 Then
                    winner_str = winner_str & " and Player " & i + 1
                Else
                    winner_str = "Player " & i + 1
                End If
                winner_count = winner_count + 1
            End If
        Next i
        If winner_count > 1 Then
            winner_str = winner_str & " win!"
        Else
            winner_str = winner_str & " wins"
        End If

        If ai_mode = 0 Then
            Call MsgBox("Robitics and Software were top cards. " & winner_str)
            Call start_new_game()
        End If
    End Sub



    ' Click on a player number
    Private Sub process_player_click(ByRef index As Short)
#If VERBOSE Then
        Call append_simple("In process_player_click()-----Click on a player number only 3 cards Roadbuilding2/optics1/rocketry")
#End If
        Dim player As Object
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        player = 0

        If index = 0 Then Exit Sub

        If phase = "roadBuilding2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_on_board(player, index, 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_on_board(index, player, 4)
            Call resume_dogma()

        ElseIf phase = "optics1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If scores(player) > scores(index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_from_score(player, index, human_data)
                Call resume_dogma()
            End If

        ElseIf phase = "rocketry" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 1 To Int(icon_total(player, 6) / 2)
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If score_pile(index, 0) > -1 Then Call return_from_score_pile(index, size2(score_pile, index) - 1)
            Next i
            Call resume_dogma()
        End If
    End Sub


    ' Click on a card on opponent's board
    Private Sub process_opp_board_click(ByRef index As Short)
#If VERBOSE Then
        Call append_simple("In process_opp_board_click()------------------")
#End If
        Dim card, clicked_player, player As Object
        ' FK not used Dim valid As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        clicked_player = Int(index / 5) + 1
        index = index Mod 5
        'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        card = board(clicked_player, index, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        player = 0

        If phase = "compass2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If clicked_player = active_player Then
                'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not card_has_symbol(card, 1) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_on_board(active_player, player, index)
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "fission" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_board(clicked_player, index)
            Call resume_dogma()

        ElseIf phase = "services" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player = clicked_player And Not card_has_symbol(board(clicked_player, index, 0), 1) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_board_to_hand(active_player, player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "bioengineering" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(clicked_player, index, 0), 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(clicked_player, index, 0), 1) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object clicked_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_board_to_score(clicked_player, player, index)
                Call resume_dogma()
            End If
        End If

    End Sub

    ' Click on a card on board
    Private Sub process_board_click(ByRef index As Short)
#If VERBOSE Then
        Call append_simple("In process_board_click()------------------")
#End If
        Dim card, player As Object
        ' FK not used Dim valid As Short
        Dim found, i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        card = board(0, index, 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        player = 0
        If phase = "splay" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, index) And splay_options(index) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call splay(player, index, splay_direction)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "cityStates" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(card, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(card, 2) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_on_board(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 1)
                Call resume_dogma()
            End If

        ElseIf phase = "philosophy1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If can_splay(player, index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call splay(player, index, "Left")
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "compass" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 1) And color_lookup(color(board(player, index, 0))) <> "Green" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_on_board(player, active_player, index)
                For i = 0 To 4
                    If board(active_player, i, 0) > -1 Then
                        If Not card_has_symbol(board(active_player, i, 0), 1) Then
                            phase = "compass2"
                        End If
                    End If
                Next i
                If phase <> "compass2" Then Call resume_dogma()
            End If

        ElseIf phase = "monotheism" Then
            If board(active_player, index, 0) = -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_board_to_score(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_tuck(player, 1)
                Call resume_dogma()
            End If

        ElseIf phase = "anatomy2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(board(player, index, 0)) = human_data Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call return_from_board(player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "enterprise" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 4) And color_lookup(color(board(player, index, 0))) <> "Purple" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_on_board(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_meld(player, 4)
                Call resume_dogma()
            End If

        ElseIf phase = "gunpowder" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, index, 0), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 2) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_board_to_score(player, active_player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "invention" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If splayed(player, index) = "Left" And can_splay(player, index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call splay(player, index, "Right")
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_score(player, 4)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "banking" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 5) And color_lookup(color(board(player, index, 0))) <> "Green" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_on_board(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_score(player, 5)
                Call resume_dogma()
            End If

        ElseIf phase = "coal3" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call score_top_card(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If board(player, index, 0) > -1 Then Call score_top_card(player, index)
            Call resume_dogma()

        ElseIf phase = "measurement2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call splay(player, index, "Right")
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call draw_num(player, size3(board, player, index))
            Call resume_dogma()

        ElseIf phase = "societies" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 3) And color_lookup(color(board(player, index, 0))) <> "Purple" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_on_board(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 5)
                Call resume_dogma()
            End If

        ElseIf phase = "pirateCode2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(board(player, index, 0)) = human_data And card_has_symbol(card, 4) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call score_top_card(player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "metricSystem" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object splay(player, index, Right). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If splay(player, index, "Right") Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
            End If
            Call resume_dogma()

        ElseIf phase = "electricity" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Not card_has_symbol(board(player, index, 0), 5) And already_returned(index) = 0 Then
                already_returned(index) = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call return_from_board(player, index)
                human_data = human_data - 1
                If human_data = 0 Then
                    For i = 0 To 4
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If already_returned(i) = 1 Then Call draw_num(player, 8)
                    Next i
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "publications" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size3(board, player, index) > 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                lblPrompt.Text = "Choose the order you want this pile in, bottom to top."
                cmdCancel.Visible = False
                Call showArray.load_pictures(0, index, "board")
            End If

        ElseIf phase = "railroad2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If splayed(player, index) = "Right" And can_splay(player, index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call splay(player, index, "Up")
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "corporations" Then
            'MsgBox title(board(player, index, 0))
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 5) And color_lookup(color(board(player, index, 0))) <> "Green" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_board_to_score(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_meld(player, 8)
                Call resume_dogma()
            End If

        ElseIf phase = "flight" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object splay(player, index, Up). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If splay(player, index, "Up") Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
            End If
            Call resume_dogma()

        ElseIf phase = "mobility" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Not card_has_symbol(board(player, index, 0), 5)) And color_lookup(color(board(player, index, 0))) <> "Red" And already_returned(index) <> 1 Then
                found = 0
                For i = 0 To 4
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If board(player, i, 0) > -1 And already_returned(i) = 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If (Not card_has_symbol(board(player, i, 0), 5)) And color_lookup(color(board(player, i, 0))) <> "Red" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If age(board(player, i, 0)) > age(board(player, index, 0)) Then found = 1
                        End If
                    End If
                Next i
                If found = 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_board_to_score(player, active_player, index)
                    already_returned(index) = 1
                    human_data = human_data - 1
                    If human_data = 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call draw_num(player, 8)
                        Call resume_dogma()
                    End If
                End If
            End If

        ElseIf phase = "skyscrapers" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 6) And color_lookup(color(board(player, index, 0))) <> "Yellow" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_on_board(player, active_player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If board(player, index, 0) > -1 Then Call score_top_card(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = size3(board, player, index) - 1 To 0 Step -1
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call return_from_board(player, index)
                Next i
                Call resume_dogma()
            End If

        ElseIf phase = "fission" Then
            If index <> 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call return_from_board(player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "globalization" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(board(player, index, 0), 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(board(player, index, 0), 1) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call return_from_board(player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "selfService" Then
            If index <> 4 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call perform_solo_dogma_effects(player, board(player, index, 0))
            End If

        End If
    End Sub


    ' Click on a card in hand
    Private Sub process_hand_click(ByRef index As Short)
#If VERBOSE Then
        Call append_simple("In process_hand_click() ++++++")
#End If
        Dim player, card, valid As Object
        Dim found As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        card = hand(0, index)
        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        player = 0
        'MsgBox phase & current_turn
        Dim min As String
        Dim min_player As Object
        Dim i As Short
        If phase = "initial_meld" Then
            ' The player with the lowest card alphabetically goes first
            min = title(hand(0, index))
            'UPGRADE_WARNING: Couldn't resolve default property of object min_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            min_player = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            break_dogma_loop = 0
#If VERBOSE Then
            Call append_simple("In process_hand_click() Initial Meld ++++++")
#End If
            Call meld_from_hand(0, index)

            ' All AI players will meld a card also
            For i = 1 To num_players - 1
                If title(hand(i, 0)) < min Then
                    min = title(hand(i, 0))
                    'UPGRADE_WARNING: Couldn't resolve default property of object min_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    min_player = i
                End If
                Call meld_from_hand(i, 0)
            Next i

            'UPGRADE_WARNING: Couldn't resolve default property of object min_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call append(min_player, "is the starting player")
            current_turn = 1
            actions_remaining = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object min_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            active_player = min_player
#If VERBOSE Then
            Call append_simple("In process_handclick() ----------")
#End If
            phase = "done with initial meld"
            Call play_game()
        ElseIf phase = "waiting for action" And current_turn <> 0 And cmdNext.Visible = False Then
            Call meld_from_hand(0, index)
            actions_remaining = actions_remaining - 1
            Call play_game()

#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 1 ++++++")
#End If


        ElseIf phase = "agriculture" Then

            Call return_from_hand(player, index)
            Call draw_and_score(player, age(card) + 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            Call resume_dogma()
        ElseIf phase = "archery" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_hand(0, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If is_highest_card_in_hand(0, index) Then
                Call transfer_card_in_hand(player, active_player, index)
                Call resume_dogma()
            End If
        ElseIf phase = "codeOfLaws" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object valid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            valid = 0
            For i = 0 To 4
                If board(player, i, 0) > -1 Then
                    If color(board(player, i, 0)) = color(card) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object valid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        valid = 1
                    End If
                End If
            Next i
            'UPGRADE_WARNING: Couldn't resolve default property of object valid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If valid = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call tuck_from_hand(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, color(card)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If can_splay(player, color(card)) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    splay_options(color(card)) = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call perform_best_splay(player, splay_options, "Left")
                    'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    lblPrompt.Text = "You may splay your " & color_lookup(color(card)) & " cards left." & vbNewLine & "(click a top card to splay its pile)"
                Else
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "clothing1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If board(0, color(card), 0) = -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call meld_from_hand(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "domestication" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object is_lowest_card_in_hand(player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If is_lowest_card_in_hand(player, index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call meld_from_hand(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "masonry" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(card, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(card, 2) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call meld_from_hand(player, index)
                human_data = human_data + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If human_data = 4 Then Call achieve_special(player, 9, "Monument")
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If hand(player, 0) = -1 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())
            End If

        ElseIf phase = "oars" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(card, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(card, 4) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call append(active_player, "steals " & title(card) & " from player " & player + 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call remove_card_from_hand(player, index)
                Call Push2(score_pile, active_player, card)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                demand_met = 1
                Call resume_dogma()
            End If

        ElseIf phase = "pottery" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            human_data = human_data + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If human_data = 3 Or hand(player, 0) = -1 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())

        ElseIf phase = "tools1" Then
            cmdCancel.Visible = False
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            human_data = human_data + 1
            If human_data = 3 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_meld(player, 3)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If

        ElseIf phase = "tools2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(card) = 3 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call return_from_hand(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                Call resume_dogma()
            End If
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 2 ++++++")
#End If
            ' Age 2
        ElseIf phase = "currency" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            already_returned(age(hand(player, index)) - 1) = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) = -1 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())

        ElseIf phase = "construction" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_in_hand(player, active_player, index)
            human_data = human_data + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If human_data = 2 Or size2(hand, player) = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 2)
                Call resume_dogma()
            End If

        ElseIf phase = "mathematics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call draw_and_meld(player, age(card) + 1)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            Call resume_dogma()

        ElseIf phase = "philosophy2" Or phase = "refrigeration2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call score_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            Call resume_dogma()

        ElseIf phase = "roadBuilding1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(player, index)
            human_data = human_data + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            If human_data = 2 Then
                phase = "roadBuilding2"
                lblPrompt.Text = "Choose a player to swap cards with (give your top Red card and get their top Green card).  Click on a player's number near their score to select that player."
            End If

            ' Age 3
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 3 ++++++")
#End If
        ElseIf phase = "alchemy1" Or phase = "physics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) = 0 Then Call resume_dogma()

        ElseIf phase = "alchemy2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) > 0 Then
                lblPrompt.Text = "Score a card from your hand"
                phase = "alchemy3"
            Else
                Call resume_dogma()
            End If

        ElseIf phase = "alchemy3" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call score_from_hand(player, index)
            Call resume_dogma()

        ElseIf phase = "feudalism" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, index), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(hand(player, index), 2) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_in_hand(player, active_player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "machinery" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object card_has_symbol(hand(player, index), 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If card_has_symbol(hand(player, index), 2) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call score_from_hand(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If active_player <> player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object can_splay(player, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If can_splay(player, 1) Then
                    splay_options(1) = 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call perform_best_splay(player, splay_options, "Left")
                    lblPrompt.Text = "You may splay your Red cards left"
                Else
                    Call resume_dogma()
                End If
            End If

            ' Age 4
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 4 ++++++")
#End If
        ElseIf phase = "perspective" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) > -1 Then
                cmdCancel.Visible = False
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                human_data = Int(icon_total(player, 3) / 2)
                phase = "perspective1"
                lblPrompt.Text = "Score a card from hand"
            Else
                Call resume_dogma()
            End If

        ElseIf phase = "perspective1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call score_from_hand(player, index)
            human_data = human_data - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) = -1 Or human_data = 0 Then
                Call resume_dogma()
            End If

        ElseIf phase = "reformation1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call tuck_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            human_data = human_data - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) = -1 Or human_data = 0 Then
                Call resume_dogma()
            End If

            ' Age 5
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 5 ++++++")
#End If

        ElseIf phase = "measurement" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            phase = "measurement2"
            lblPrompt.Text = "Choose a color to splay right"
            cmdCancel.Visible = False

            ' Age 6
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 6 ++++++")
#End If
        ElseIf phase = "classification" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            human_data = color(card)
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call collect_all_cards(player, color(card))
            lblPrompt.Text = "Meld all " & color_lookup(human_data) & " cards."
            phase = "classification2"

        ElseIf phase = "classification2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If color(card) = human_data Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call meld_from_hand(player, index)
                found = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = 0 To size2(hand, player) - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If color(hand(player, i)) = human_data Then found = 1
                Next i
                If found = 0 Then Call resume_dogma()
            End If

        ElseIf phase = "democracy" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            human_data = human_data + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) = -1 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())

        ElseIf phase = "emancipation" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_hand_to_score(player, active_player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call draw_num(player, 6)
            Call resume_dogma()

            ' Age 7
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 7 ++++++")
#End If
        ElseIf phase = "explosives" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(hand(player, index)) = age(hand(player, size2(hand, player) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_in_hand(player, active_player, index)
                human_data = human_data - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If hand(player, 0) = -1 Then Call draw_num(player, 7)
                If human_data = 0 Then Call resume_dogma()
            End If

        ElseIf phase = "lighting" Then
            human_data = human_data - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            already_returned(age(hand(player, index)) - 1) = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call tuck_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If human_data = 0 Or hand(player, 0) = -1 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())

        ElseIf phase = "railroad" Then
            'MsgBox ("returning " & title(index))
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) = 0 Then
                For i = 1 To 3
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_num(player, 6)
                Next i
                Call resume_dogma()
            End If

        ElseIf phase = "refrigeration" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            human_data = human_data - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) = 0 Or human_data = 0 Then Call resume_dogma()

        ElseIf phase = "sanitation" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(hand(player, index)) = age(hand(player, size2(hand, player) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_in_hand(player, active_player, index)
                human_data = human_data + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If human_data = 2 Or hand(player, 0) = -1 Then
                    If already_returned(0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_in_hand(active_player, player, already_returned(0))
                    End If
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "sanitation2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(hand(player, index)) = age(hand(player, 0)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_in_hand(player, human_data, index)
                If hand(human_data, 0) > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_in_hand(human_data, player, size2(hand, human_data) - 1)
                    If hand(human_data, 0) > -1 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call transfer_card_in_hand(human_data, player, size2(hand, human_data) - 1)
                    End If
                End If
                Call resume_dogma()
            End If

            ' Age 8
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 8 ++++++")
#End If
        ElseIf phase = "antibiotics" Then
            human_data = human_data - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            already_returned(age(hand(player, index)) - 1) = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            If human_data = 0 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())


        ElseIf phase = "massMedia" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            For i = 0 To 9
                lblPrompt.Text = "Choose a number to return from each score pile."
                cmdCancel.Visible = False
                cmdOddball(i).Visible = True
                cmdOddball(i).Text = CStr(i + 1)
            Next i

        ElseIf phase = "quantumTheory" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            human_data = human_data + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            If human_data = 2 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 10)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_and_score(player, 10)
                Call resume_dogma()
            End If

        ElseIf phase = "socialism" Then
            cmdCancel.Visible = False
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If color_lookup(color(hand(player, index))) = "Purple" Then human_data = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call tuck_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) = -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call steal_lowest_cards(player)
                Call resume_dogma()
            End If

            ' Age 9
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 9 ++++++")
#End If
        ElseIf phase = "composites" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_in_hand(player, active_player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If score_pile(player, 0) > -1 Then
                    Call show_prompt("composites2", "Transfer the highest card in your score pile.", 0)
                Else
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "ecology" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) > -1 Then
                cmdCancel.Visible = False
                phase = "ecology1"
                lblPrompt.Text = "Score a card from hand"
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 10)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call draw_num(player, 10)
                Call resume_dogma()
            End If

        ElseIf phase = "ecology1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call score_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call draw_num(player, 10)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call draw_num(player, 10)
            Call resume_dogma()

        ElseIf phase = "satellites" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(hand, player) = 0 Then
                For i = 1 To 3
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_num(player, 8)
                Next i
                Call resume_dogma()
            End If

        ElseIf phase = "satellites3" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call perform_solo_dogma_effects(player, card)


        ElseIf phase = "specialization" Then
            For i = 0 To num_players - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If i <> player And board(i, color(card), 0) > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_board_to_hand(i, player, color(card))
                End If
            Next i
            Call resume_dogma()

        ElseIf phase = "suburbia" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call tuck_from_hand(player, index)
            human_data = human_data + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If hand(player, 0) = -1 Then Call cmdCancel_Click(cmdCancel, New System.EventArgs())

            ' Age 10
#If VERBOSE Then
            Call append_simple("In process_hand_click() AGE 10 ++++++")
#End If
        ElseIf phase = "miniaturization" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If active_player <> player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object valid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            valid = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object valid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(hand(player, index)) = 10 Then valid = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_hand(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object valid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If valid = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For i = 0 To size2(score_pile, player) - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If already_returned(age(score_pile(player, i))) = 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        already_returned(age(score_pile(player, i))) = 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Call draw_num(player, 10)
                    End If
                Next i
            End If
            Call resume_dogma()


        End If

    End Sub

    Private Sub steal_lowest_cards(ByVal player As Short)
#If VERBOSE Then
        Call append_simple("In steal_lowest_cards() ++++++")
#End If
        Dim found, i, j As Short
        For i = 0 To num_players - 1
            If hand(i, 0) > -1 And i <> player Then
                found = age(hand(i, 0))
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For j = size2(hand, i) - 1 To 0 Step -1
                    If age(hand(i, j)) = found Then
                        Call transfer_card_in_hand(i, player, j)
                    End If
                Next j
            End If
        Next i
    End Sub


    Public Sub process_score_pile_click(ByVal index As Short)
#If VERBOSE Then
        Call append_simple("In process_score_pile_click() ++++++")
#End If
        Dim player, card, id As Object
        ' FK not used Dim found,valid As Object
        Dim count, j As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        card = score_pile(0, index)
        'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        player = 0
        If phase = "mapmaking" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_from_score(player, active_player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "education" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_score_pile(player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If is_highest_card_in_score_pile(player, index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call return_from_score_pile(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If size2(score_pile, player) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object get_highest_card_in_score_pile(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_num(player, age(score_pile(player, get_highest_card_in_score_pile(player))) + 2)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_num(player, 2)
                End If
                Call resume_dogma()
            End If

        ElseIf phase = "medicine" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = age(score_pile(player, size2(score_pile, player) - 1)) Then
                If human_data = -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_from_score(player, active_player, index)
                Else
                    Call remove_card_from_score_pile(active_player, 0)
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_from_score(player, active_player, index)
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call Push2(score_pile, player, human_data)
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call append(active_player, " transfers a " & age(human_data) & " from his score pile to player " & player + 1 & "'s score pile")
                End If
                Call resume_dogma()
            End If

        ElseIf phase = "medicine1" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = age(score_pile(player, 0)) Then
                If score_pile(human_data, 0) = -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_from_score(player, human_data, index)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call remove_card_from_score_pile(player, index)
                    'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call transfer_card_from_score(human_data, player, size2(score_pile, human_data) - 1)
                    Call Push2(score_pile, human_data, card)
                    'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call append(player, " transfers a " & age(card) & " from his score pile to player " & human_data + 1 & "'s score pile")
                End If
                Call resume_dogma()
            End If

        ElseIf phase = "optics" Then
            phase = "optics1"
            human_data = index
            lblPrompt.Text = "Choose a player with a lower score to give this card to."

        ElseIf phase = "translation" Then
            cmdCancel.Visible = False
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call meld(player, score_pile(player, index))
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call remove_card_from_score_pile(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If size2(score_pile, player) = 0 Then Call resume_dogma()

        ElseIf phase = "anatomy" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            id = score_pile(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_score_pile(player, index)
            For j = 0 To 4
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If board(player, j, 0) > -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If age(board(player, j, 0)) = age(id) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        human_data = age(id)
                        'UPGRADE_WARNING: Couldn't resolve default property of object id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        lblPrompt.Text = "Return a top card of value " & age(id)
                        phase = "anatomy2"
                        Exit Sub
                    End If
                End If
            Next j
            Call resume_dogma()

        ElseIf phase = "navigation" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = 2 Or age(score_pile(player, index)) = 3 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_from_score(player, active_player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "printingPress" Then
            count = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If board(player, 2, 0) > -1 Then count = age(board(player, 2, 0))
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If player <> active_player Then dogma_copied = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_score_pile(player, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call draw_num(player, count + 2)
            Call resume_dogma()

        ElseIf phase = "chemistry2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_score_pile(player, index)
            Call resume_dogma()

        ElseIf phase = "statistics" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object is_highest_card_in_score_pile(player, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If is_highest_card_in_score_pile(player, index) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_score_to_hand(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not (size2(hand, player) = 1 And score_pile(player, 0) > -1) Then
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "pirateCode" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object card. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(card) < 5 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_from_score(player, active_player, index)
                human_data = human_data - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If score_pile(player, 0) = -1 Then
                    Call resume_dogma()
                    Exit Sub
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If age(score_pile(player, 0)) > 4 Or human_data = 0 Then Call resume_dogma()
            End If

        ElseIf phase = "encyclopedia" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = human_data Then
                cmdCancel.Visible = False
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If player <> active_player Then dogma_copied = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call meld(player, score_pile(player, index))
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call remove_card_from_score_pile(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If age(score_pile(player, 0)) = -1 Then
                    Call resume_dogma()
                    Exit Sub
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If age(score_pile(player, size2(score_pile, player) - 1)) <> human_data Then Call resume_dogma()
            End If

        ElseIf phase = "vaccination" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = human_data Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call remove_card_from_score_pile(player, index)
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If score_pile(player, 0) = -1 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_and_meld(player, 6)
                    Call resume_dogma()
                    Exit Sub
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If age(score_pile(player, 0)) <> human_data Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Call draw_and_meld(player, 6)
                    Call resume_dogma()
                End If
            End If

        ElseIf phase = "combustion" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call transfer_card_from_score(player, active_player, index)
            human_data = human_data - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If score_pile(player, 0) = -1 Or human_data = 0 Then Call resume_dogma()

        ElseIf phase = "evolution2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_score_pile(player, index)
            Call resume_dogma()

        ElseIf phase = "composites2" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If age(score_pile(player, index)) = age(score_pile(player, size2(score_pile, player) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Call transfer_card_from_score(player, active_player, index)
                Call resume_dogma()
            End If

        ElseIf phase = "databases" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call return_from_score_pile(player, index)
            human_data = human_data - 1
            If human_data = 0 Then Call resume_dogma()

        End If
    End Sub


    Private Sub lblAchievement_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblAchievement.Click
        Dim index As Short = lblAchievement.GetIndex(eventSender)
        If phase = "waiting for action" And index < 9 And cmdNext.Visible = False Then
            'UPGRADE_WARNING: Couldn't resolve default property of object achieve(0, index + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If achieve(0, index + 1) Then
                Call play_game()
            End If
        End If
    End Sub

    Public Function achieve(ByVal player As Short, ByVal num As Short) As Object

        ' MsgBox "trying to click " & num
#If VERBOSE Then
        Call append(0, "In achieve() +++trying to click " & num)
#End If
        If num = 0 Then MsgBox("trying to score 0 ")
        'UPGRADE_WARNING: Couldn't resolve default property of object can_achieve(player, num). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If can_achieve(player, num) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object achieve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            achieve = TriState.True
            achievements_scored(num - 1) = 1
            vps(player) = vps(player) + 1
            Call append(player, "has claimed the -" & num & "- achievement")
            If player = 0 Then actions_remaining = actions_remaining - 1
            Call update_display()
            Call end_game_points()
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object achieve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            achieve = TriState.False
        End If
    End Function

    Public Function can_achieve(ByVal player As Short, ByVal num As Short) As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object highest_top_card(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object scores(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If achievements_scored(num - 1) = 0 And scores(player) >= 5 * num And highest_top_card(player) >= num Then
            'UPGRADE_WARNING: Couldn't resolve default property of object can_achieve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            can_achieve = TriState.True ' FK? 3 state checkboxes: Visble Locked, Visible Unlocked, Invisible
#If VERBOSE Then
            Call append(player, "In can_achieve() +++ HEY U CAN ACHIEVE!")
#End If
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object can_achieve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            can_achieve = TriState.False
#If VERBOSE Then
            Call append(player, "In can_achieve() +++ HEY u cannot ACHIEVE")
#End If
        End If
    End Function


    ''''''''''''''''''''''''
    '
    ' GUI mouseover Handler Functions
    '
    ''''''''''''''''''''''''

    Public Sub imgMouseMove(ByRef img As Object, ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef Y As Single)
        With img
            'This range check fails on VB.net
            'If Not ((X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height)) Then
            load_picture(img.Tag)
            'End If
        End With
    End Sub

    Private Sub imgHandColor_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgHandColor.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgHandColor.GetIndex(eventSender)
        Call imgMouseMove(imgHandColor(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgHandIcon_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgHandIcon.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgHandIcon.GetIndex(eventSender)
        Call imgMouseMove(imgHandIcon(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblHand_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblHand.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblHand.GetIndex(eventSender)
        Call imgMouseMove(lblHand(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgScoreColor_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgScoreColor.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgScoreColor.GetIndex(eventSender)
        Call imgMouseMove(imgScoreColor(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgScoreIcon_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgScoreIcon.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgScoreIcon.GetIndex(eventSender)
        Call imgMouseMove(imgScoreIcon(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblOtherApps_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblOtherApps.Click
        ShellExecute(Me.Handle.ToInt32, "open", "http://www.slightlymagic.net/wiki/Other_Apps_by_jatill", vbNullString, "", 0)
    End Sub

    Private Sub lblScoreTitle_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblScoreTitle.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblScoreTitle.GetIndex(eventSender)
        Call imgMouseMove(lblScoreTitle(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgBoard_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgBoard.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgBoard.GetIndex(eventSender)
        ' MsgBox ("called " & imgBoard(Index).Tag)

        Call imgMouseMove(imgBoard(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgIconSmall_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgIconSmall.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgIconSmall.GetIndex(eventSender)
        Call imgMouseMove(imgIconSmall(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgBoardDogma_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgBoardDogma.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgBoardDogma.GetIndex(eventSender)
        Call imgMouseMove(imgBoardDogma(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblBoardTitle_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblBoardTitle.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblBoardTitle.GetIndex(eventSender)
        Call imgMouseMove(lblBoardTitle(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblBoardAge_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblBoardAge.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblBoardAge.GetIndex(eventSender)
        Call imgMouseMove(lblBoardAge(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblDogma_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblDogma.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblDogma.GetIndex(eventSender)
        Call imgMouseMove(lblDogma(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblOppBoard_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblOppBoard.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblOppBoard.GetIndex(eventSender)
        Call imgMouseMove(lblOppBoard(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgOppBoardColor_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgOppBoardColor.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgOppBoardColor.GetIndex(eventSender)
        Call imgMouseMove(imgOppBoardColor(index), Button, Shift, X, Y)
    End Sub

    Private Sub imgOppBoard_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgOppBoard.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = imgOppBoard.GetIndex(eventSender)
        Call imgMouseMove(imgOppBoard(index), Button, Shift, X, Y)
    End Sub

    Private Sub lblAchievement_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblAchievement.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim index As Short = lblAchievement.GetIndex(eventSender)
        ' FK mod to add achievements 0-9
        ' If index > 8 Then
        'MsgBox "loading " & Index
        'imgLarge.Image = achievement_images(index - 9)
        imgLarge.Image = achievement_images(index)
        imgLarge.BringToFront()
        'End If
    End Sub

    ''''''''''''''''''''''''
    '
    ' GUI Click Handlers
    '
    ''''''''''''''''''''''''
    Private Sub lblOppBoard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblOppBoard.Click
        Dim index As Short = lblOppBoard.GetIndex(eventSender)
        Call process_opp_board_click(index)
    End Sub
    Private Sub imgOppBoard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgOppBoard.Click
        Dim index As Short = imgOppBoard.GetIndex(eventSender)
        Call process_opp_board_click(index)
    End Sub
    Private Sub imgOppBoardColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgOppBoardColor.Click
        Dim index As Short = imgOppBoardColor.GetIndex(eventSender)
        Call process_opp_board_click(index)
    End Sub

    Private Sub lblHand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblHand.Click
        Dim index As Short = lblHand.GetIndex(eventSender)
        Call process_hand_click(index)
    End Sub
    Private Sub imgHandIcon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgHandIcon.Click
        Dim index As Short = imgHandIcon.GetIndex(eventSender)
        Call process_hand_click(index)
    End Sub
    Private Sub imgHandColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgHandColor.Click
        Dim index As Short = imgHandColor.GetIndex(eventSender)
        Call process_hand_click(index)
    End Sub

    Private Sub lblOppDetail_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblOppDetail.Click
        Dim index As Short = lblOppDetail.GetIndex(eventSender)
        Call showArray.load_pictures(Int(index / 5) + 1, index Mod 5, "board")
    End Sub

    Private Sub lblPlayer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblPlayer.Click
        Dim index As Short = lblPlayer.GetIndex(eventSender)
        Call process_player_click(index)
    End Sub

    Private Sub lblPlayerDetail_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblPlayerDetail.Click
        Dim index As Short = lblPlayerDetail.GetIndex(eventSender)
        Call process_player_click(index + 1)
    End Sub

    Private Sub lblScoreTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblScoreTitle.Click
        Dim index As Short = lblScoreTitle.GetIndex(eventSender)
        Call process_score_pile_click(index)
    End Sub
    Private Sub imgScoreIcon_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgScoreIcon.Click
        Dim index As Short = imgScoreIcon.GetIndex(eventSender)
        Call process_score_pile_click(index)
    End Sub
    Private Sub imgScoreColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgScoreColor.Click
        Dim index As Short = imgScoreColor.GetIndex(eventSender)
        Call process_score_pile_click(index)
    End Sub



    Private Sub imgBoard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBoard.Click
        Dim index As Short = imgBoard.GetIndex(eventSender)
        Call process_board_click(index)
    End Sub

    Private Sub imgIconSmall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgIconSmall.Click
        Dim index As Short = imgIconSmall.GetIndex(eventSender)
        Call process_board_click(Int(index / 4))
    End Sub

    Private Sub imgBoardDogma_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBoardDogma.Click
        Dim index As Short = imgBoardDogma.GetIndex(eventSender)
        Call process_board_click(index)
    End Sub

    Private Sub lblBoardTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblBoardTitle.Click
        Dim index As Short = lblBoardTitle.GetIndex(eventSender)
        Call process_board_click(index)
    End Sub

    Private Sub lblBoardAge_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblBoardAge.Click
        Dim index As Short = lblBoardAge.GetIndex(eventSender)
        Call process_board_click(index)
    End Sub

    Private Sub lblDogma_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblDogma.Click
        Dim index As Short = lblDogma.GetIndex(eventSender)
        Call process_board_click(index)
    End Sub

    Private Sub lblGoals_Click(sender As Object, e As EventArgs) Handles lblGoals.Click

    End Sub

    Private Sub _lblAchievement_0_Click(sender As Object, e As EventArgs) Handles _lblAchievement_0.Click

    End Sub

    Private Sub _lblBoardTitle_1_Click(sender As Object, e As EventArgs) Handles _lblBoardTitle_1.Click

    End Sub

    Private Sub _cmdDogma_3_Click(sender As Object, e As EventArgs) Handles _cmdDogma_3.Click

    End Sub

    Private Sub _imgLargeIcon_0_Click(sender As Object, e As EventArgs) Handles _imgLargeIcon_0.Click

    End Sub

    Private Sub lblPrompt_Click(sender As Object, e As EventArgs) Handles lblPrompt.Click

    End Sub

    Private Sub txtLog_TextChanged(sender As Object, e As EventArgs) Handles txtLog.TextChanged

    End Sub

    Private Sub _lblAchievement_13_Click(sender As Object, e As EventArgs) Handles _lblAchievement_13.Click

    End Sub


End Class
