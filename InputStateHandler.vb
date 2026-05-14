Public NotInheritable Class InputStateHandler
    Private _currentKeyState As New Dictionary(Of Keys, Boolean)
    Private _previousKeyState As New Dictionary(Of Keys, Boolean)
    Private _currentMouseState As New Dictionary(Of MouseButtons, Boolean)
    Private _previousMouseState As New Dictionary(Of MouseButtons, Boolean)
    Private _mousePos As Vec2i = Vec2i.Create(0, 0)
    
    Public Sub UpdatePreviousStates()
        ' Clear and copy current to previous
        _previousKeyState.Clear()
        For Each kvp In _currentKeyState
            _previousKeyState(kvp.Key) = kvp.Value
        Next
        
        _previousMouseState.Clear()
        For Each kvp In _currentMouseState
            _previousMouseState(kvp.Key) = kvp.Value
        Next
    End Sub
    
    Public Sub SetKeyState(key As Keys, pressed As Boolean)
        _currentKeyState(key) = pressed
    End Sub
    
    Public Sub SetMouseState(button As MouseButtons, pressed As Boolean)
        _currentMouseState(button) = pressed
    End Sub
    
    Public Sub SetMousePosition(x As Integer, y As Integer)
        _mousePos = Vec2i.Create(x, y)
    End Sub
    
    Public Function GetMousePosition() As Vec2i
        Return _mousePos
    End Function
    
    Public Function IsKeyDown(key As Keys) As Boolean
        Dim currentPressed As Boolean = False
        Dim previousPressed As Boolean = False
        
        Return (_currentKeyState.TryGetValue(key, currentPressed) AndAlso currentPressed) AndAlso
            (Not _previousKeyState.TryGetValue(key, previousPressed) OrElse Not previousPressed)
    End Function
    
    Public Function IsKeyHeld(key As Keys) As Boolean
        Dim pressed As Boolean = False
        Return _currentKeyState.TryGetValue(key, pressed) AndAlso pressed
    End Function
    
    Public Function IsKeyUp(key As Keys) As Boolean
        Dim currentPressed As Boolean = False
        Dim previousPressed As Boolean = False
        
        Return (_currentKeyState.TryGetValue(key, currentPressed) AndAlso Not currentPressed) AndAlso
            (_previousKeyState.TryGetValue(key, previousPressed) AndAlso previousPressed)
    End Function
    
    Public Function IsMouseDown(button As MouseButtons) As Boolean
        Dim currentPressed As Boolean = False
        Dim previousPressed As Boolean = False
        
        Return (_currentMouseState.TryGetValue(button, currentPressed) AndAlso currentPressed) AndAlso
            (Not _previousMouseState.TryGetValue(button, previousPressed) OrElse Not previousPressed)
    End Function
    
    Public Function IsMouseHeld(button As MouseButtons) As Boolean
        Dim pressed As Boolean = False
        Return _currentMouseState.TryGetValue(button, pressed) AndAlso pressed
    End Function
    
    Public Function IsMouseUp(button As MouseButtons) As Boolean
        Dim currentPressed As Boolean = False
        Dim previousPressed As Boolean = False
        
        Return (_currentMouseState.TryGetValue(button, currentPressed) AndAlso Not currentPressed) AndAlso
            (_previousMouseState.TryGetValue(button, previousPressed) AndAlso previousPressed)
    End Function
    
    Public Sub ResetStates()
        _currentKeyState.Clear()
        _previousKeyState.Clear()
        _currentMouseState.Clear()
        _previousMouseState.Clear()
    End Sub
End Class