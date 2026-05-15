Imports Microsoft.ClearScript

Public Module Collections
    Public Function ArrayOf(ParamArray items As Object()) As Object()
        Return items
    End Function

    Public Function CreateArrayList(Optional capacity As Integer = 0) As ArrayList
        Return If(capacity > 0, New ArrayList(capacity), New ArrayList)
    End Function

    Public Function CreateHashSet() As HashSet(Of Object)
        Return New HashSet(Of Object)
    End Function

    Public Function CreateHashtable() As Hashtable
        Return New Hashtable
    End Function

    Public Function Where(source As Object, predicate As Object) As ArrayList
        Dim result As New ArrayList()
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return result

        For Each item In enumerable
            Dim predicateResult = InvokeFunction(predicate, {item})
            If CBool(predicateResult) Then result.Add(item)
        Next

        Return result
    End Function

    Public Function [Select](source As Object, selector As Object) As ArrayList
        Dim result As New ArrayList
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return result
        For Each item In enumerable
            Dim selectedValue = InvokeFunction(selector, {item})
            result.Add(selectedValue)
        Next

        Return result
    End Function

    Public Function First(source As Object, Optional predicate As Object = Nothing) As Object
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return Nothing

        If predicate Is Nothing Then
            For Each item In enumerable : Return item : Next
        Else
            For Each item In enumerable
                Dim predicateResult = InvokeFunction(predicate, {item})
                If CBool(predicateResult) Then Return item
            Next
        End If

        Return Nothing
    End Function

    Public Function Last(source As Object, Optional predicate As Object = Nothing) As Object
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return Nothing
        Dim lastItem As Object = Nothing

        If predicate Is Nothing Then
            For Each item In enumerable : lastItem = item : Next item
        Else
            For Each item In enumerable
                Dim predicateResult = InvokeFunction(predicate, {item})
                If CBool(predicateResult) Then lastItem = item
            Next item
        End If

        Return lastItem
    End Function

    Public Function Count(source As Object, Optional predicate As Object = Nothing) As Integer
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return 0

        Dim countResult As Integer = 0
        If predicate Is Nothing Then
            For Each item In enumerable : countResult += 1 : Next
        Else
            For Each item In enumerable
                Dim predicateResult = InvokeFunction(predicate, {item})
                If CBool(predicateResult) Then countResult += 1
            Next
        End If

        Return countResult
    End Function

    Public Function Any(source As Object, Optional predicate As Object = Nothing) As Boolean
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return False
        If predicate Is Nothing Then Return Count(enumerable) > 0

        For Each item In enumerable
            Dim predicateResult = InvokeFunction(predicate, {item})
            If CBool(predicateResult) Then Return True
        Next item

        Return False
    End Function

    Public Function All(source As Object, predicate As Object) As Boolean
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)

        If enumerable Is Nothing Then Return False
        For Each item In enumerable
            Dim predicateResult = InvokeFunction(predicate, {item})
            If Not CBool(predicateResult) Then Return False
        Next item

        Return True
    End Function

    Public Function Contains(source As Object, value As Object) As Boolean
        Dim areEqual = Function(obj1 As Object, obj2 As Object)
                           If obj1 Is Nothing AndAlso obj2 Is Nothing Then Return True
                           If obj1 Is Nothing OrElse obj2 Is Nothing Then Return False
                           Return obj1.Equals(obj2)
                       End Function

        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return False
        For Each item In enumerable
            If areEqual(item, value) Then Return True
        Next item

        Return False
    End Function

    Public Function Sort(source As ArrayList) As ArrayList
        source.Sort()
        Return source
    End Function

    Public Function Sort(source As ArrayList, comparer As Object) As ArrayList
        Dim comparison As New Comparison(Of Object)(Function(x, y) CInt(InvokeFunction(comparer, {x, y})))
        source.Sort(CType(comparison, IComparer))
        Return source
    End Function

    Public Function Reverse(source As ArrayList) As ArrayList
        source.Reverse()
        Return source
    End Function

    Public Function Take(source As Object, count As Integer) As ArrayList
        Dim result As New ArrayList
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)

        If enumerable Is Nothing OrElse count <= 0 Then Return result

        Dim currCount As Integer = 0
        For Each item In enumerable
            If currCount >= count Then Exit For
            result.Add(item)
            currCount += 1
        Next

        Return result
    End Function

    Public Function Skip(source As Object, count As Integer) As ArrayList
        Dim result As New ArrayList
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)
        If enumerable Is Nothing Then Return result

        Dim currCount As Integer = 0
        For Each item In enumerable
            If currCount >= count Then result.Add(item)
            currCount += 1
        Next

        Return result
    End Function

    Public Function ToArray(source As Object) As Array
        Dim list As New ArrayList
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)

        If enumerable Is Nothing Then Return list.ToArray()
        For Each item In enumerable : list.Add(item) : Next

        Return list.ToArray()
    End Function

    Public Function ToArrayList(source As Object) As ArrayList
        Dim list As New ArrayList
        Dim enumerable As IEnumerable = TryCast(source, IEnumerable)

        If enumerable Is Nothing Then Return list
        For Each item In enumerable : list.Add(item) : Next
        Return list
    End Function

    Private Function InvokeFunction(func As Object, arguments As Object()) As Object
        If func Is Nothing Then Return Nothing

        Try
            Dim scriptObject = TryCast(func, IScriptObject)
            If scriptObject IsNot Nothing Then
                Return scriptObject.Invoke(False, arguments)
            End If

            Dim invokeMethod = func.GetType().GetMethod("Invoke")
            If invokeMethod IsNot Nothing Then
                Return invokeMethod.Invoke(func, arguments)
            End If

            Dim dynamicInvoke = func.GetType().GetMethod("DynamicInvoke")
            If dynamicInvoke IsNot Nothing Then
                Return dynamicInvoke.Invoke(func, arguments)
            End If
        Catch
        End Try

        Return Nothing
    End Function
End Module