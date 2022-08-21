'
' 
' Functions.vb This code is licenced under "Creative Commons Attribution Non Commercial 4.0 International"
' See: https://creativecommons.org/licenses/by-nc/4.0/legalcode
'
'

' Conditional compile directives

#Const VERBOSE = True ' FK adding debugging Frame work via c compiler like if defs should be DEBUG but not sure about interference

Option Strict Off
Option Explicit On

#Disable Warning IDE0054 'Inherited code with assignments (60) this suppresses: Use compound assignment x += 5 vs x = x + 5
#Disable Warning IDE1006 'Disable Naming rule violation: These words must begin with upper case characters:

Module ArrayFunctions
    Public arr(4, 5, 100) As Short
    Public Sub InsertArrayItem(ByRef arr As Object, ByVal index As Integer, ByVal newValue As Object)
        Dim i As Integer
        Dim j As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = size(arr)
        For i = j To index Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(i + 1) = arr(i)
        Next
        'UPGRADE_WARNING: Couldn't resolve default property of object newValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index) = newValue
    End Sub

    Public Sub DeleteArrayItem(ByRef arr As Object, ByVal index As Integer)
        Dim i As Integer
        Dim j As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = size(arr)
        For i = index To j - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(i) = arr(i + 1)
        Next
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(j) = -1
    End Sub

    Public Sub DeleteArrayItem2(ByRef arr As Object, ByVal index As Short, ByVal item As Short)
        Dim temp_arr(500) As Object
        Call flatten(arr, index, temp_arr)
        Call DeleteArrayItem(temp_arr, item)
        Call rebuild(arr, index, temp_arr)
    End Sub

    Public Sub DeleteArrayItem3(ByRef arr As Object, ByVal index1 As Short, ByVal index2 As Short, ByVal index As Integer)
        Dim i As Integer
        Dim j As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = size3(arr, index1, index2)
        For i = index To j - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(index1, index2, i) = arr(index1, index2, i + 1)
        Next
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index1, index2, j) = -1
    End Sub

    Public Sub Shift(ByRef arr As Object)
        Call DeleteArrayItem(arr, 0)
    End Sub

    Public Sub Shift2(ByRef arr As Object, ByVal index As Short)
        Dim temp_arr(500) As Object
        Call flatten(arr, index, temp_arr)
        Call Shift(temp_arr)
        Call rebuild(arr, index, temp_arr)
    End Sub

    Public Sub shift3(ByRef arr As Object, ByVal index1 As Short, ByVal index2 As Short)
        Call DeleteArrayItem3(arr, index1, index2, 0)
    End Sub

    Public Sub Unshift(ByRef arr As Object, ByVal newValue As Object)
        Dim i As Integer
        Dim s As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        s = size(arr)
        For i = s To 1 Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(i) = arr(i - 1)
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object newValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(0) = newValue
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(s + 1) = -1
    End Sub

    Public Sub Unshift2(ByRef arr As Object, ByVal index As Short, ByVal newValue As Object)
        Dim temp_arr(500) As Object
        Call flatten(arr, index, temp_arr)
        Call Unshift(temp_arr, newValue)
        Call rebuild(arr, index, temp_arr)
    End Sub

    Public Sub Unshift3(ByRef arr As Object, ByVal index1 As Short, ByVal index2 As Short, ByVal newValue As Object)
        Dim i As Integer
        Dim s As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        s = size3(arr, index1, index2)
        For i = s To 1 Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(index1, index2, i) = arr(index1, index2, i - 1)
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object newValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index1, index2, 0) = newValue
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index1, index2, s + 1) = -1
    End Sub

    Public Sub Push(ByRef arr As Object, ByVal newValue As Object)
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = size(arr)
        'UPGRADE_WARNING: Couldn't resolve default property of object newValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(i) = newValue
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(i + 1) = -1
    End Sub

    Public Sub Push2(ByRef arr As Object, ByVal index As Short, ByVal newValue As Object)
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = size2(arr, index)
        'UPGRADE_WARNING: Couldn't resolve default property of object newValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index, i) = newValue
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index, i + 1) = -1
    End Sub

    Public Sub Push3(ByRef arr As Object, ByVal index1 As Short, ByVal index2 As Short, ByVal newValue As Object)
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size3(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = size3(arr, index1, index2)
        'UPGRADE_WARNING: Couldn't resolve default property of object newValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index1, index2, i) = newValue
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr(index1, index2, i + 1) = -1
    End Sub

    Public Function flatten(ByRef arr As Object, ByVal index As Short, ByRef temp_arr As Object) As Object
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(arr, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(arr, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object temp_arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp_arr(i) = arr(index, i)
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object temp_arr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object flatten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        flatten = temp_arr
    End Function

    Public Function flatten2(ByRef arr As Object, ByVal index As Short, ByRef temp_arr As Object) As Object
        Dim i As Short
        For i = 0 To 2
            'MsgBox "setting " & i & " to " & arr(index, i)
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object temp_arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp_arr(i) = arr(index, i)
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object temp_arr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object flatten2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        flatten2 = temp_arr
    End Function

    Public Sub rebuild(ByRef arr As Object, ByVal index As Short, ByRef temp_arr As Object)
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(arr, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(arr, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object temp_arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(index, i) = temp_arr(i)
        Next i
    End Sub

    Public Sub copy_array(ByRef arr_from As Object, ByRef arr_to As Object)
        Dim i As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(arr_from). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size(arr_from)
            'UPGRADE_WARNING: Couldn't resolve default property of object arr_from(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arr_to(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr_to(i) = arr_from(i)
        Next i
    End Sub

    Public Sub copy_array2(ByRef arr_from As Object, ByRef arr_to As Object, ByVal index As Short)
        Dim i As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object size2(arr_from, index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size2(arr_from, index)
            'UPGRADE_WARNING: Couldn't resolve default property of object arr_from(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arr_to(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr_to(index, i) = arr_from(index, i)
        Next i
    End Sub

    Public Sub copy_array3(ByRef arr_from As Object, ByRef arr_to As Object, ByVal index1 As Short, ByVal index2 As Short)
        Dim i As Integer
        ' Not used Dim j As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size3(arr_from, index1, index2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 0 To size3(arr_from, index1, index2)
            'UPGRADE_WARNING: Couldn't resolve default property of object arr_from(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arr_to(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr_to(index1, index2, i) = arr_from(index1, index2, i)
        Next i
    End Sub

    Public Function size(ByRef arr As Object) As Object
        Dim i As Short
        i = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        While arr(i) <> -1
            i = i + 1
        End While
        'UPGRADE_WARNING: Couldn't resolve default property of object size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        size = i
    End Function

    Public Function size2(ByRef arr As Object, ByVal index As Short) As Object
        Dim i As Short
        i = 0
        'If Index > 5 Then MsgBox "Index: " & Index
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(index, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        While arr(index, i) <> -1
            i = i + 1
            If i > 400 Then
                'MsgBox "Blown up... looking for index " & Index
                'MsgBox "Hand size: " & size2(hands, Index)
                'MsgBox "Discard size: " & size2(discards, Index)
                'MsgBox "Deck size: " & size2(decks, Index)
            End If
        End While
        'UPGRADE_WARNING: Couldn't resolve default property of object size2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        size2 = i
    End Function

    Public Function size3(ByRef arr As Object, ByVal index1 As Short, ByVal index2 As Short) As Object
        Dim i As Short
        i = 0
        'If Index > 5 Then MsgBox "Index: " & Index
        'UPGRADE_WARNING: Couldn't resolve default property of object arr(index1, index2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        While arr(index1, index2, i) <> -1
            i = i + 1
        End While
        'UPGRADE_WARNING: Couldn't resolve default property of object size3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        size3 = i
    End Function

    Public Sub sort(ByRef arr As Object)
        ' Dim k, i, j, temp As Object
        Dim k, i, j, temp, v As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(arr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 1 To size(arr) - 1
            For j = 0 To i - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                v = arr(j)
                'UPGRADE_WARNING: Couldn't resolve default property of object arr(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If arr(i) < v Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    temp = arr(i)
                    For k = i - 1 To j Step -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        arr(k + 1) = arr(k)
                    Next k
                    'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    arr(j) = temp
                End If
            Next j
        Next i
    End Sub

    Public Sub rsort(ByRef arr As Object)
        ' Dim k, i, j, temp As Object
        Dim k, i, j, temp, value As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(arr). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For i = 1 To size(arr) - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            value = arr(i)
            For j = 0 To i - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object arr(j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If value > arr(j) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    temp = arr(i)
                    For k = i - 1 To j Step -1
                        'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        arr(k + 1) = arr(k)
                    Next k
                    'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    arr(j) = temp
                End If
            Next j
        Next i
    End Sub

    Public Sub sort2(ByRef arr As Object, ByVal index As Short)
        Dim temp_arr(500) As Object
        Call flatten(arr, index, temp_arr)
        Call sort(temp_arr)
        Call rebuild(arr, index, temp_arr)
    End Sub

    Public Sub rsort2(ByRef arr As Object, ByVal index As Short)
        Dim temp_arr(500) As Object
        Call flatten(arr, index, temp_arr)
        Call rsort(temp_arr)
        Call rebuild(arr, index, temp_arr)
    End Sub

    Public Sub randomize_array(ByRef arr As Object)
        Dim new_array(500) As Object
        Dim picked(500) As Object
        Randomize()
        'Dim arr_size, i As Object
        Dim arr_size, i, r As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object size(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        arr_size = size(arr)
        For i = 0 To arr_size - 1
            r = Int(arr_size * Rnd())
            'UPGRADE_WARNING: Couldn't resolve default property of object picked(r). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            While (picked(r) = 1)
                r = Int(arr_size * Rnd())
            End While
            'UPGRADE_WARNING: Couldn't resolve default property of object picked(r). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            picked(r) = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object new_array(r). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            new_array(r) = arr(i)
        Next i

        For i = 0 To arr_size - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object new_array(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            arr(i) = new_array(i)
        Next i
        'UPGRADE_WARNING: Couldn't resolve default property of object new_array(arr_size). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        new_array(arr_size) = -1
#If VERBOSE2 Then
        Call Main_Renamed.append_simple("In randomize_array() ++++++")
        Call print_array(new_array, "Shuffled")
#End If
        'Call print_array(new_array, "Shuffled")
    End Sub

    Public Sub randomize_array2(ByRef arr As Object, ByVal index As Short)
        Dim temp_arr(500) As Object
        Call flatten(arr, index, temp_arr)
        Call randomize_array(temp_arr)
        Call rebuild(arr, index, temp_arr)
    End Sub

    Public Sub print_array(ByRef arr As Object, ByVal msg As String)
        'Dim i As Object
        Dim i, j As Short
        'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim str_Renamed As String
        str_Renamed = msg & Chr(13)
        'UPGRADE_WARNING: Couldn't resolve default property of object size(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = size(arr)
        For j = 0 To i - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object arr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            str_Renamed = str_Renamed & j & ") " & arr(j) & Chr(13)
        Next j
        MsgBox(str_Renamed)
    End Sub

    ' FK not being called by anybody
    ' Public Sub print_array2(ByRef arr As Object, ByVal index As Short, ByVal msg As String)
    'Dim temp_arr(500) As Object
    'Call flatten(arr, index, temp_arr)
    'Call print_array(temp_arr, msg)
    'End Sub
End Module
