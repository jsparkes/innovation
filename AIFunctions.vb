'
' 
' AIFunctions.vb This code is licenced under "Creative Commons Attribution Non Commercial 4.0 International"
' See: https://creativecommons.org/licenses/by-nc/4.0/legalcode
'
'
' This file implements the AI functions for WinInnovation

' Conditional compile directives

#Const VERBOSE = True ' FK adding debugging Frame work via c compiler like if defs should be DEBUG but not sure about interference


Option Strict Off
Option Explicit On


' Compiler/Build directives

#Disable Warning IDE1006 'Inherited code with variable names this suppresses: These words must begin with upper case characters
#Disable Warning IDE0054 'Inherited code with assignments (60) this suppresses: Use compound assignment x += 5 vs x = x + 5
#Disable Warning BC40000 'VB compatibility warning messages are suppressed

Module AIFunctions

    Public game_state(100000, 2000) As Short
    Public game_phase(100000) As String
    Public game_state_strings(100000) As String ' gss = game state strings?
    Public gss_count, gs_count As Object ' FK why not Integer?
    Public human_data As Integer

    ' FK Consts added for code readbility
    Const AGECOUNT As Short = 9 ' FK Ages are 1-10 IRL, code uses 0-9
    Const ICONCOUNT As Short = 6 ' FK Icons are 1-6 IRL Leaf. Castle, Lightbulb, Crown, Factory ,Clock, Code has an extra placeholder X
    Const COLORCOUNT As Short = 4 ' FK Colors are 5 Yellow, Red, Purple, Blue, Green
    Const MAXPLAYERS As Short = 3 ' FK 4 0-3 players max
    Const MAXPLAYER As Short = 4 ' FK 4 0-3 players max
    Const PLAYERMODE As Short = 0 'FK checking as player
    Const AIMODE As Short = 1 'FK checking as AI

    ' Game Data
    Public num_players As Short
    Public active_player As Short
    Public num_cards As Short
    Public deck(10, 25) As Short
    Public ai_mode As Short
    Public achievements_scored(20) As Short
    Public phase As String
    Public current_turn As Short
    Public actions_remaining As Short
    Public dogma_copied, dogma_level, dogma_player, break_dogma_loop As Object
    Public affected_by_dogma(MAXPLAYER, 3) As Object
    Public dogma_id As Short
    Public demand_met As Object
    Public democracy_minimum As Short
    Public solo_dogma_player, solo_dogma_id As Object
    Public solo_dogma_level As Short
    Public silent As Object
    Public gss_depth As Short
    Public winners(MAXPLAYER) As Short

    'Player Data
    Public hand(MAXPLAYER, 100) As Short
    Public board(MAXPLAYER, 5, 100) As Short
    Public score_pile(MAXPLAYER, 100) As Short
    Public splayed(MAXPLAYER, 4) As String
    Public icon_total(MAXPLAYER, 7) As Short
    Public scores(MAXPLAYER) As Object
    Public scored_this_turn(MAXPLAYER) As Object
    Public tucked_this_turn(MAXPLAYER) As Short
    Public vps(MAXPLAYER) As Short ' Victory Points ?

    ' Card Data
    Public age(500) As Short
    Public title(500) As String
    Public color(500) As Short
    Public icons(500, 4) As Short
    Public dogma_icon(500) As Short
    Public dogma(500, 3) As String
    Public color_lookup(5) As String
    Public icon_lookup(ICONCOUNT) As String
    Public is_demand(500, 3) As Short

    Public test_card As Short

    Public Function make_ai_move(ByVal player As Short, ByVal choices As Object, ByVal depth As Short) As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In make_ai_move() ++++++")
#End If
        ' Actions:
        ' 0 - achieve
        ' 1 - draw a card
        ' 2 - activate dogma
        ' 3 - meld card from hand
        Dim action_choice As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        action_choice = choices(0)
        If ai_mode = PLAYERMODE And depth > 0 Then MsgBox("depth got too big")
        'MsgBox "called with depth = " & depth
#If VERBOSE Then
        Call Main_Renamed.append(player, "In make_ai_move() called with depth = " & depth)
        Call Main_Renamed.append(player, "In make_ai_move() Trying move " & choices(0) & "," & choices(1) & " at level " & depth & " (ai_mode = " & ai_mode & ")  gss_count = " & gss_count)
#End If
        'Open App.Path & "\log.txt" For Append As #1
        'Print #1, "Trying move " & choices(0) & "," & choices(1) & " at level " & depth & " (ai_mode = " & ai_mode & ")  " & gss_count
        'If choices(0) = 2 Then Print #1, title(choices(1))
        'Close #1

        'MsgBox "Trying move " & choices(0) & "," & choices(1) & " at level " & depth & " (ai_mode = " & ai_mode & ")  " & gss_count

        'If action_choice = 3 Then MsgBox "Choosing " & names(constructs(active_player, choices(1))) & " used: " & construct_used(active_player, choices(1))

        'If ai_mode = 0 And depth = 0 And (action_choice = 0 Or action_choice = 3) And choices(1) = -1 Then MsgBox "the weirdness"
        If ai_mode = PLAYERMODE Then
            actions_remaining = actions_remaining - 1
            'MsgBox "Made a real move " & choices(0) & "," & choices(1) & " actions_remaining = " & actions_remaining
#If VERBOSE Then
            Call Main_Renamed.append(player, "In make_ai_move() Made a real move " & choices(0) & "," & choices(1) & " actions_remaining = " & actions_remaining)
#End If
        End If

        If action_choice = 3 Then
            ' choices(1) is which card to play from hand
            'MsgBox "option 0, depth=" & depth & ", ai_mode=" & ai_mode
            'MsgBox "trying a meld from ai - " & title(hand(player, choices(1))) & " depth = " & depth
            'UPGRADE_WARNING: Couldn't resolve default property of object choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
#If VERBOSE Then
            Call Main_Renamed.append(player, "In make_ai_move() " & choices(1) & "  is the card to play from hand ----")
            Call Main_Renamed.append(player, "In make_ai_move() option 0 - Achieve, depth= " & depth & ", ai_mode= " & ai_mode)
            Call Main_Renamed.append(player, "In make_ai_move() trying a meld from ai - " & title(hand(player, choices(1))) & " depth = " & depth)
#End If
            Call Main_Renamed.meld_from_hand(player, choices(1))
            'MsgBox "melded"
#If VERBOSE Then
            Call Main_Renamed.append_simple("In make_ai_move() AI Doing meld ----")
#End If
        ElseIf action_choice = 2 Then
            ' choices(1) is the id of the dogma to activate

            'MsgBox "Performing dogma " & title(choices(1))
            'UPGRADE_WARNING: Couldn't resolve default property of object choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
#If VERBOSE Then
            Call Main_Renamed.append(player, "In make_ai_move() Performing dogma " & title(choices(1)))
#End If
            Call Main_Renamed.perform_dogma_effect(choices(1))
        ElseIf action_choice = 1 Then
            Call Main_Renamed.draw(player)
#If VERBOSE Then
            Call Main_Renamed.append_simple("In make_ai_move() AI doing DRAW ----")
#End If
        ElseIf action_choice = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Call Main_Renamed.achieve(player, choices(1))
        End If

        If ai_mode = AIMODE Then
            ' If we are in AI mode, and have already tried this move, then back out
            gss_depth = depth
            'UPGRADE_WARNING: Couldn't resolve default property of object is_repeat_gss(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If is_repeat_gss() Then
                'MsgBox "We have already been at this game state... skipping"
                'UPGRADE_WARNING: Couldn't resolve default property of object make_ai_move. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                make_ai_move = -100000
#If VERBOSE Then
                Call Main_Renamed.append_simple("In make_ai_move() AI doing exit did I die? ----We have already been at this game state... skipping")
#End If
                Exit Function
            End If

            ' Stop once we reach a certain depth
            If depth = 0 Then
                'MsgBox "Going deeper, into depth " & depth + 1
#If VERBOSE Then
                Call Main_Renamed.append(player, "In make_ai_move() Going deeper, into depth = " & depth + 1)
#End If

                'UPGRADE_WARNING: Couldn't resolve default property of object pick_ai_move(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                make_ai_move = pick_ai_move(player, choices, depth + 1)
            Else
                'MsgBox "done, time to score"

                'If demand_met = 1 Then MsgBox "Got there 1!"
                'UPGRADE_WARNING: Couldn't resolve default property of object score_game(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                make_ai_move = score_game(player)
                'MsgBox "Got to bottom of a chain, and score = " & make_ai_move
#If VERBOSE Then
                Call Main_Renamed.append_simple("In make_ai_move() done, time to score")
                Call Main_Renamed.append(player, "In make_ai_move() done, Got to bottom of a chain, and score = " & make_ai_move)
#End If
            End If
            '    MsgBox "Going to look for another play after doing " & choices(0) & "," & choices(1)
#If VERBOSE Then
            Call Main_Renamed.append(player, "In make_ai_move() Going to look for another play after doing choices0/1... " & choices(0) & " , " & choices(1))
#End If
        End If

    End Function


    Public Function pick_ai_move(ByVal player As Short, ByRef choices As Object, ByVal depth As Short) As Object


        ' Dim max_score, score, max_index As Object
        Dim max_score, score As Integer
        Dim count, i, max_index, State As Short
        Dim all_choices(1000, 3) As Short

        ' If there is a card we're testing, always activate it
#If VERBOSE Then
        Call Main_Renamed.append_simple("In pick_ai_move() If there is a card we're testing, always activate it")
#End If
        If test_card > 0 Then
            If board(player, color(test_card), 0) = test_card Then
                'MsgBox ("picking bogus move")
#If VERBOSE Then
                Call Main_Renamed.append_simple("In pick_ai_move() picking bogus move")
#End If
                'UPGRADE_WARNING: Couldn't resolve default property of object choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                choices(0) = 2
                'UPGRADE_WARNING: Couldn't resolve default property of object choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                choices(1) = test_card
                Exit Function
            End If
        End If

        'MsgBox "calling pick with depth = " & depth
#If VERBOSE Then
        Call Main_Renamed.append(player, "In pick_ai_move() calling pick with depth = " & depth)
#End If
        count = 0
        If count = 0 Then
#If VERBOSE Then
            Call Main_Renamed.append_simple("In pick_ai_move() Add the choice to ACHIEVE")
#End If
            'Add the choice to achieve
            For i = 1 To 9
                'UPGRADE_WARNING: Couldn't resolve default property of object Main_Renamed.can_achieve(player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Main_Renamed.can_achieve(player, i) Then
                    all_choices(count, 0) = 0
                    all_choices(count, 1) = i
                    count = count + 1
                    i = 9
                End If
            Next i

            'Add the choice to draw
#If VERBOSE Then
            Call Main_Renamed.append_simple("In pick_ai_move() Add the choice to DRAW")
#End If
            all_choices(count, 0) = 1
            count = count + 1

            ' Get all buy choices
#If VERBOSE Then
            Call Main_Renamed.append_simple("In pick_ai_move() Get all buy choices")
#End If
            'UPGRADE_WARNING: Couldn't resolve default property of object get_all_dogma_choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            count = get_all_dogma_choices(player, all_choices, count)
            'MsgBox " choice1 = " & count
#If VERBOSE Then
            Call Main_Renamed.append(0, "In pick_ai_move()  choice1 = " & count)
#End If
            ' Get all cards in hand we might want to play
            'UPGRADE_WARNING: Couldn't resolve default property of object get_all_hand_choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            count = get_all_hand_choices(player, all_choices, count)
            'MsgBox " choice2 = " & count
#If VERBOSE Then
            Call Main_Renamed.append(0, "In pick_ai_move()  Get all cards in hand we might want to play choice2 = " & count)
#End If

        End If
        ' Make each move and see which scores best
#If VERBOSE Then
        Call Main_Renamed.append_simple("In pick_ai_move() Make each move and see which scores best ++++++")
#End If

        'ai_mode = 1
        max_score = -100000
        ' MsgBox count & " choices are available at depth " & depth & " Game States: " & gs_count
#If VERBOSE Then
        Call Main_Renamed.append(player, "In pick_ai_move() Make each move and see which scores best ++++++" & count & " choices are available at depth " & depth & " Game States: " & gs_count)
#End If
        State = gs_count
        copy_game_state(State)
        gs_count = gs_count + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To count - 1
            Call flatten2(all_choices, i, choices)
            If ai_mode = 0 And depth > 0 Then MsgBox("bad depth in pick_action")
            'UPGRADE_WARNING: Couldn't resolve default property of object make_ai_move(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            score = make_ai_move(player, choices, depth)
            If depth = 0 Then
                '    MsgBox "Making move " & i + 1 & " of " & count & " -> " & choices(0) & " , " & choices(1) & " , score=" & score & " at depth " & depth
#If VERBOSE Then
                Call Main_Renamed.append(player, "In pick_ai_move() Making move " & i + 1 & " of " & count & " -> " & choices(0) & " , " & choices(1) & " , score=" & score & " at depth " & depth)
#End If
            End If
            'MsgBox "Score at depth " & depth & ": " & score
#If VERBOSE Then
            Call Main_Renamed.append(player, "In pick_ai_move() Making move  ++++++ Score at depth " & depth & ": " & score)
#End If
            If score > max_score Then
                max_score = score
                max_index = i
                'MsgBox "new max score: " & max_score
#If VERBOSE Then
                Call Main_Renamed.append(player, "In pick_ai_move() new max score = " & max_score)
#End If
            End If
            restore_game_state(State)
        Next i
        gs_count = State
        Call flatten2(all_choices, max_index, choices)
        'MsgBox "Choosing choice " & max_index & "(" & choices(0) & "," & choices(1) & ") at depth " & depth & " with score " & max_score
#If VERBOSE Then
        Call Main_Renamed.append(player, "In pick_ai_move() Making move  ++++++ Choosing choice max_index = " & max_index & "( choices(0) = " & choices(0) & ", choices(1) =" & choices(1) & ") at depth = " & depth & " with max_score = " & max_score)
#End If
        'UPGRADE_WARNING: Couldn't resolve default property of object pick_ai_move. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        pick_ai_move = max_score


    End Function


    Private Function get_all_hand_choices(ByVal player As Short, ByRef all_choices As Object, ByVal choice_count As Short) As Object
        ' If we can play any card from hand, then do it

#If VERBOSE Then
        Call Main_Renamed.append_simple("In get_all_hand_choices() If we can play any card from hand, then do it ++++++")
#End If
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object all_choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            all_choices(choice_count, 0) = 3
            'UPGRADE_WARNING: Couldn't resolve default property of object all_choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            all_choices(choice_count, 1) = i
            choice_count = choice_count + 1
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_all_hand_choices. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_all_hand_choices = choice_count
#If VERBOSE Then
        Call Main_Renamed.append(player, "In get_all_hand_choices() The number of HAND choices = " & choice_count)
#End If
    End Function

    Private Function get_all_dogma_choices(ByVal player As Short, ByRef all_choices As Object, ByVal choice_count As Short) As Object
        ' If we can activate any dogma, then do it
#If VERBOSE Then
        Call Main_Renamed.append_simple("In get_all_dogma_choice() If we can activate any dogma, then do it --------")
#End If

        Dim i, id As Short
        For i = 0 To COLORCOUNT
            id = board(player, i, 0)
            If id > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object all_choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                all_choices(choice_count, 0) = 2
                'UPGRADE_WARNING: Couldn't resolve default property of object all_choices(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                all_choices(choice_count, 1) = board(player, i, 0)
                choice_count = choice_count + 1
            End If
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_all_dogma_choices. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_all_dogma_choices = choice_count
#If VERBOSE Then
        Call Main_Renamed.append(player, "In get_all_dogma_choice() The number of DOGMA choices = " & choice_count)
#End If
    End Function

    Public Function is_repeat_gss() As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In is_repeat_gss() --------")
#End If
        Dim str_Renamed As String
        'UPGRADE_WARNING: Couldn't resolve default property of object get_game_state_string(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        str_Renamed = get_game_state_string()

        Dim i As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object gss_count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 1 To gss_count - 1

            'If gss_count = 3 Then MsgBox "Comparing " & i & Chr(13) & game_state_strings(i) & Chr(13) & str
#If VERBOSE Then
            If gss_count = 3 Then
                Call Main_Renamed.append(0, "In is_repeat_gss() Comparing " & i & Chr(13) & game_state_strings(i) & Chr(13) & str_Renamed)
            End If
#End If
            If str_Renamed = game_state_strings(i) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object is_repeat_gss. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                is_repeat_gss = 1
                'MsgBox "Match found"
#If VERBOSE Then
                Call Main_Renamed.append_simple("In is_repeat_gss() Match found --------")
#End If
                Exit Function
            End If
        Next i

        ' Not a repeat, add on this value
        'UPGRADE_WARNING: Couldn't resolve default property of object gss_count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state_strings(gss_count) = str_Renamed
        'UPGRADE_WARNING: Couldn't resolve default property of object gss_count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gss_count = gss_count + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object is_repeat_gss. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        is_repeat_gss = 0
    End Function

    Public Function get_game_state_string() As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In get_game_state_string() ++++++")
#End If

        Dim j, i, k As Short
        Dim str_Renamed As String

        str_Renamed = gss_depth & "|"
        For i = 0 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            str_Renamed = str_Renamed & scores(i) & "|"
            str_Renamed = str_Renamed & vps(i) & "|"
            str_Renamed = str_Renamed & winners(i) & "|"
#If VERBOSE Then
            Call Main_Renamed.append(0, "In get_game_state_string() ++++++ str_Renamed = " & str_Renamed)
#End If
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For j = 0 To size2(hand, i)
                str_Renamed = str_Renamed & hand(i, j) & "|"
            Next j
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For j = 0 To size2(score_pile, i)
                str_Renamed = str_Renamed & score_pile(i, j) & "|"
            Next j
            For j = 0 To COLORCOUNT
                If board(i, j, k) > -1 Then
                    str_Renamed = str_Renamed & splayed(i, j) & "|"
                    'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, i, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    For k = 0 To size3(board, i, j) - 1
                        str_Renamed = str_Renamed & board(i, j, k) & "|"
                    Next k
                End If
            Next j
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object get_game_state_string. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        get_game_state_string = str_Renamed
#If VERBOSE Then
        Call Main_Renamed.append(0, "In get_game_state_string() ++++++ FINISHED str_Renamed = " & str_Renamed)
#End If
    End Function

    Public Sub save_game()
#If VERBOSE Then
        Call Main_Renamed.append_simple("In save_game() writing to \save.txt ++++++")
#End If

        Dim i, n As Short
        FileOpen(1, My.Application.Info.DirectoryPath & "\save.txt", OpenMode.Output)
        'UPGRADE_WARNING: Couldn't resolve default property of object copy_game_state(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        n = copy_game_state(0)
        For i = 0 To n - 1
            PrintLine(1, game_state(0, i))
        Next i

        FileClose(1)
    End Sub

    Public Function load_game() As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In load_game() reading from \save.txt ++++++")
#End If

        Dim i As Short
        Main_Renamed.txtLog.Text = ""
        FileOpen(1, My.Application.Info.DirectoryPath & "\save.txt", OpenMode.Input)
        On Error GoTo jeff
        i = 0
        While (1)
            Input(1, game_state(0, i))
            i = i + 1
        End While

jeff:
        FileClose(1)
        Call restore_game_state(0)
        Main_Renamed.cmdCancel.Visible = False
        Call Main_Renamed.update_display()
    End Function

    Public Function copy_game_state(ByVal Index As Short) As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In copy_game_state() Backup hands, boards etc ----")
#End If
        ' Dim player As Object
        Dim k, i, j, player, n As Short
        n = 0

        game_state(Index, n) = num_players
        n = n + 1

        ' Back up hands, boards, etc
        For player = 0 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            game_state(Index, n) = scores(player)
            n = n + 1

            game_state(Index, n) = vps(player)
            n = n + 1

            game_state(Index, n) = winners(player)
            n = n + 1

            'UPGRADE_WARNING: Couldn't resolve default property of object scored_this_turn(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            game_state(Index, n) = scored_this_turn(player)
            n = n + 1

            game_state(Index, n) = tucked_this_turn(player)
            n = n + 1

            ' Hand
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            game_state(Index, n) = size2(hand, player)
            n = n + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 0 To size2(hand, player)
                game_state(Index, n) = hand(player, i)
                n = n + 1
            Next i

            ' Score Pile
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            game_state(Index, n) = size2(score_pile, player)
            n = n + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object size2(score_pile, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 0 To size2(score_pile, player)
                game_state(Index, n) = score_pile(player, i)
                n = n + 1
            Next i

            ' Boards
            For j = 0 To COLORCOUNT
                If splayed(player, j) = "" Then game_state(Index, n) = 0
                If splayed(player, j) = "Left" Then game_state(Index, n) = 1
                If splayed(player, j) = "Right" Then game_state(Index, n) = 2
                If splayed(player, j) = "Up" Then game_state(Index, n) = 3
                If splayed(player, j) = "Down" Then game_state(Index, n) = 4
                n = n + 1

                'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                game_state(Index, n) = size3(board, player, j)
                n = n + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(board, player, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For k = 0 To size3(board, player, j)
                    game_state(Index, n) = board(player, j, k)
                    n = n + 1
                Next k
            Next j

            ' Icon Total
            For i = 1 To ICONCOUNT
                game_state(Index, n) = icon_total(player, i)
                n = n + 1
            Next i

            ' Affected by dogma
            For i = 0 To 2
                'UPGRADE_WARNING: Couldn't resolve default property of object affected_by_dogma(player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                game_state(Index, n) = affected_by_dogma(player, i)
                n = n + 1
            Next i
        Next player

#If VERBOSE Then
        Call Main_Renamed.append(0, "In copy_game_state() ++++++ n = " & n)
        ' shouldn't n be set back to zero here? n=0
#End If

        ''''''''''''''''''''''''''
        ' Game Variables
        '''''''''''''''''''''''''''
        game_state(Index, n) = active_player
        n = n + 1

        game_state(Index, n) = current_turn
        n = n + 1

        game_state(Index, n) = actions_remaining
        n = n + 1

        game_state(Index, n) = democracy_minimum
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = dogma_level
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = dogma_player
        n = n + 1

        game_state(Index, n) = solo_dogma_level
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = solo_dogma_player
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = solo_dogma_id
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = dogma_copied
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = break_dogma_loop
        n = n + 1

        game_state(Index, n) = dogma_id
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        game_state(Index, n) = demand_met
        n = n + 1

        For i = 0 To 14
            game_state(Index, n) = achievements_scored(i)
            n = n + 1
        Next i

        For i = 0 To AGECOUNT
            For j = 0 To 14
                game_state(Index, n) = deck(i, j)
                n = n + 1
            Next j
        Next i

        'UPGRADE_WARNING: Couldn't resolve default property of object copy_game_state. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
#If VERBOSE Then
        Call Main_Renamed.append(0, "Exiting copy_game_state() ++++++ n = " & n)
        'FK shouldn't n be set back to zero here? n=0
#End If
        copy_game_state = n
    End Function

    Public Function restore_game_state(ByVal Index As Short) As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In restore_game_state() ++++++")
#End If
        ' Dim player As Object
        Dim j, i, k, n, n0, player As Short
        n = 0

        num_players = game_state(Index, n)
        n = n + 1

        ' Restore hands, boards, etc
        For player = 0 To num_players - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            scores(player) = game_state(Index, n)
            n = n + 1

            vps(player) = game_state(Index, n)
            n = n + 1

            winners(player) = game_state(Index, n)
            n = n + 1

            'UPGRADE_WARNING: Couldn't resolve default property of object scored_this_turn(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            scored_this_turn(player) = game_state(Index, n)
            n = n + 1

            tucked_this_turn(player) = game_state(Index, n)
            n = n + 1

            ' Hand
            n0 = game_state(Index, n)

            n = n + 1
            For i = 0 To n0
                hand(player, i) = game_state(Index, n)
                n = n + 1
            Next i

            ' Score Pile
            n0 = game_state(Index, n)
            n = n + 1
            For i = 0 To n0
                score_pile(player, i) = game_state(Index, n)
                n = n + 1
            Next i

            ' Boards
            For j = 0 To COLORCOUNT
                If game_state(Index, n) = 0 Then splayed(player, j) = ""
                If game_state(Index, n) = 1 Then splayed(player, j) = "Left"
                If game_state(Index, n) = 2 Then splayed(player, j) = "Right"
                If game_state(Index, n) = 3 Then splayed(player, j) = "Up"
                If game_state(Index, n) = 4 Then splayed(player, j) = "Down"

                n = n + 1

                n0 = game_state(Index, n)
                n = n + 1
                For k = 0 To n0
                    board(player, j, k) = game_state(Index, n)
                    n = n + 1
                Next k
            Next j

            ' Icon Total
            For i = 1 To ICONCOUNT
                icon_total(player, i) = game_state(Index, n)
                n = n + 1
            Next i

            ' Affected by dogma
            For i = 0 To 2
                'UPGRADE_WARNING: Couldn't resolve default property of object affected_by_dogma(player, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                affected_by_dogma(player, i) = game_state(Index, n)
                n = n + 1
            Next i
        Next player

#If VERBOSE Then
        Call Main_Renamed.append(0, "In restore_game_state() ++++++ n = " & n)
#End If
        ''''''''''''''''''''''''''
        ' Game Variables
        '''''''''''''''''''''''''''
        ' 0 Active Player
        ' 1 Current Turn
        ' 2 Actions Remaining
        ' 3 democracy minimum
        ' 4 dogma_level
        ' 5 dogma_player
        ' 6 solo_dogma_level
        ' 7 solo_dogma_player
        ' 8 solo_dogma_id
        ' 9 dogma_copied
        ' 10 break_dogma_loop
        ' 11 dogma_id
        ' 12 demand_met

        active_player = game_state(Index, n)
        n = n + 1

        current_turn = game_state(Index, n)
        n = n + 1

        actions_remaining = game_state(Index, n)
        n = n + 1

        democracy_minimum = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_level = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_player = game_state(Index, n)
        n = n + 1

        solo_dogma_level = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_player. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        solo_dogma_player = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object solo_dogma_id. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        solo_dogma_id = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object dogma_copied. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dogma_copied = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object break_dogma_loop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        break_dogma_loop = game_state(Index, n)
        n = n + 1

        dogma_id = game_state(Index, n)
        n = n + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object demand_met. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        demand_met = game_state(Index, n)
        n = n + 1

        For i = 0 To 14
            achievements_scored(i) = game_state(Index, n)
            n = n + 1
        Next i

        For i = 0 To AGECOUNT
            For j = 0 To 14
                deck(i, j) = game_state(Index, n)
                n = n + 1
            Next j
        Next i
        'MsgBox "Game restored"
        'UPGRADE_WARNING: Couldn't resolve default property of object restore_game_state. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
#If VERBOSE Then
        Call Main_Renamed.append_simple("In restore_game_state() ++++++ Game restored")
#End If
        restore_game_state = n
    End Function

    Public Function score_game(ByVal player As Short) As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In score_game() ++++++")
#End If
        'Dim total As Object
        Dim total As Integer
        Dim i As Short
        For i = 0 To num_players - 1
            If i <> player Then
                'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                total = total - score_game_individual(i)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                total = total + 2 * (num_players - 1) * score_game_individual(i)
            End If
        Next i
        'total = score_game_individual(player)
#If VERBOSE Then
        Call Main_Renamed.append(0, "In score_game() ++++++ total= " & score_game_individual(player))
#End If
        'UPGRADE_WARNING: Couldn't resolve default property of object score_game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        score_game = total
    End Function

    Public Function score_game_individual(ByVal player As Short) As Object
#If VERBOSE Then
        Call Main_Renamed.append_simple("In score_game_individual() ++++++")
#End If
        ' If somehow a course of action has led this player to win the game, give them a jillion points
        'Dim num_at_max, max, top_card As Object
        'Dim j, i, top_score As Short
        Dim j, i, top_card As Short
        Dim num_at_max, max, top_score, total As Integer 'FK short +- 32K integers bigger for score

        Call Main_Renamed.update_scores()
        Call Main_Renamed.update_icon_total()

        total = 0

        If winners(player) = 1 Then
            total = total + 1000000 - 100 * gss_depth
            'MsgBox "winning player found: " & player + 1 & " with " & scores(1) & " vs " & scores(2)
#If VERBOSE Then
            Call Main_Renamed.append(0, "In score_game_individual() winning player found:" & player + 1 & " with " & scores(1) & " vs " & scores(2))
#End If
        End If

        ' Grant massive points for achievements
        total = total + 10000 * vps(player)
#If VERBOSE Then
        Call Main_Renamed.append(player, "In score_game_individual() Grant massive points for achievements total = " & total)
#End If

        ' Award points for value of top card ( 5 - 500 pts )
        'UPGRADE_WARNING: Couldn't resolve default property of object Main_Renamed.highest_top_card(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        top_card = Main_Renamed.highest_top_card(player)
        total = total + 5 * top_card ^ 2
#If VERBOSE Then
        Call Main_Renamed.append(0, "In score_game_individual() Award points for value of top card ( 5 - 500 pts ) total = " & total)
#End If
        ' Award points for values of other top cards
        ' Award points for splays ( Up = 40, Right = 25, Left = 10)
        ' Award 20 points for each color in play + 1 point for the depth of the pile
#If VERBOSE Then
        Call Main_Renamed.append_simple("Award points for values of other top cards")
        Call Main_Renamed.append_simple("Award points for splays ( Up = 40, Right = 25, Left = 10)")
        Call Main_Renamed.append_simple("Award 20 points for each color in play + 1 point for the depth of the pile")
#End If
        For i = 0 To COLORCOUNT
            If board(player, i, 0) > -1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                total = total + 20 + size3(board, player, i) + age(board(player, i, 0))
                If splayed(player, i) = "Up" Then total = total + 40
                If splayed(player, i) = "Right" Then total = total + 25
                If splayed(player, i) = "Left" Then total = total + 10
#If VERBOSE Then

                Call Main_Renamed.append(player, "Award points for splays ( Up = 40, Right = 25, Left = 10) now total = " & total)
#End If
            End If
        Next i

        ' Award 75 points for each icon you're winning (50 for ties), plus 2 for each extra icon
        For j = 1 To ICONCOUNT ' icon type
#If VERBOSE Then
            Call Main_Renamed.append(0, "In score_game_individual() Award 75 points for each icon you're winning (50 for ties), plus 2 for each extra icon = " & j)
#End If
            max = 0
            num_at_max = 1
            For i = 0 To num_players - 1
                If icon_total(i, j) = max Then num_at_max = num_at_max + 1
                If icon_total(i, j) > max Then
                    max = icon_total(i, j)
                    num_at_max = 1
                End If
            Next i
            total = total + 2 * icon_total(player, j)
            If icon_total(i, j) = max And num_at_max = 1 Then total = total + 75
            If icon_total(i, j) = max And num_at_max > 1 Then total = total + 50
        Next j
#If VERBOSE Then
        Call Main_Renamed.append(0, "In score_game_individual() after icon scoring total = " & total)
#End If
        ' Award points equal to the total value of hand
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(hand, player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(hand, player) - 1
            total = total + 2 * age(hand(player, i))
        Next i

        ' Award points for your score pile (only do exponents for scores that are close to your top card)
#If VERBOSE Then
        Call Main_Renamed.append(0, "In score_game_individual()  total value after hand scoring = " & total)
        Call Main_Renamed.append_simple("In score_game_individual()  Award points for your score pile (only do exponents for scores that are close to your top card)")
#End If
        If top_card > 8 Then top_card = 8
        top_score = 5 * (top_card + 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object scores(player). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If scores(player) > top_score Then
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            total = total + 3 * top_score ^ 1.5 + (scores(player) - top_score) / 2
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object scores(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            total = total + 3 * scores(player) ^ 1.5
        End If
        'If player = 3 Then MsgBox "adding " & scores(3) & " making total " & total

        'UPGRADE_WARNING: Couldn't resolve default property of object score_game_individual. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        score_game_individual = total

    End Function
End Module
