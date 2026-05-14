Imports System.Threading
Imports Microsoft.ClearScript.Windows
Imports System.IO

Public NotInheritable Class AppMain
    Private Shared ReadOnly _instance As New Lazy(Of AppMain)(
        Function() New AppMain, LazyThreadSafetyMode.ExecutionAndPublication
    )

    Public Shared ReadOnly Property Instance As AppMain
        Get
            Return _instance.Value
        End Get
    End Property

    Private ReadOnly _scriptEngine As New VBScriptEngine
    Private ReadOnly _inputStateHandler As New InputStateHandler
    Private _gameLoop As System.Windows.Forms.Timer = Nothing
    Private _graphics As Graphics = Nothing
    Private _lastFrameTime As Long = 0
    Private _deltaTime As Single = 0.0F

    Private Sub New()
        InitializeComponent()
        DoubleBuffered = True
        Text = "vbs-revive: VBScript Game Engine"
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        StartPosition = FormStartPosition.Manual
        Location = New Point(0, 0)
        CreateAssetFolders()
        InitializeScriptEngine()
    End Sub

    Private Sub AppMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadGameScript()  ' Exception already handled
        StartGameLoop()
    End Sub

    Private Sub CreateAssetFolders()
        For Each folder In {"sounds", "images", "fonts", "music"}
            If Not Directory.Exists(folder) Then
                Directory.CreateDirectory(folder)
            End If
        Next folder

        If Not File.Exists("gamemain.vbs") OrElse
            String.IsNullOrWhiteSpace(File.ReadAllText("gamemain.vbs")) Then
            File.WriteAllText("gamemain.vbs", DEFAULT_GAME_SCRIPT)
        End If
    End Sub

    Private Const DEFAULT_GAME_SCRIPT As String = "Option Explicit
Dim playerRect, playerSpeed

Sub Initialize()
    AppMain.Instance.SetWindowSize 800, 600
    playerRect = Recti.CreateFromCenter(Vec2i.Create(400, 300), 50, 50)
    playerSpeed = 200  ' Unit: pixels per second
End Sub

Sub Update(dt)
    Dim actualSpeed
    actualSpeed = Convert.ToInt32(playerSpeed * dt)
    
    With AppMain.Instance
        If .IsKeyHeld(Keys.Left) Or .IsKeyHeld(Keys.A) Then
            playerRect.Offset -actualSpeed, 0
        End If
        If .IsKeyHeld(Keys.Right) Or .IsKeyHeld(Keys.D) Then
            playerRect.Offset actualSpeed, 0
        End If
        If .IsKeyHeld(Keys.Up) Or .IsKeyHeld(Keys.W) Then
            playerRect.Offset 0, -actualSpeed
        End If
        If .IsKeyHeld(Keys.Down) Or .IsKeyHeld(Keys.S) Then
            playerRect.Offset 0, actualSpeed
        End If
    End With
    With playerRect
        If .X < 0 Then .X = 0
        If .Right > 800 Then .X = 800 - playerRect.Width
        If .Y < 0 Then .Y = 0
        If .Bottom > 600 Then .Y = 600 - playerRect.Height
    End With
End Sub

Sub Render(g, dt)
    Dim font, text, fmtDt
    g.Clear Color.Black
    playerRect.DrawFilled g, Color.Cyan
    playerRect.DrawOutline g, 3
    Set font = FontAsset.Create(""Arial"")
    text = ""Welcome! Use W,A,S,D or Arrow Keys to move around.""
    font.DrawText g, text, 10, 10, Color.Cyan
    fmtDt = Strings.Format(dt * 1000, ""0.00"")
    text = ""Speed: "" & playerSpeed & "" px/s; Delta Time: "" & fmtDt & "" ms""
    font.DrawText g, text, 10, 30, Color.Cyan
End Sub
"

#Region "Public API for VBScript Scripting"
    Public Sub SetWindowSize(width As Integer, height As Integer)
        ClientSize = New Size(width, height)
    End Sub

    Public Function GetMousePosition() As Vec2i
        Return _inputStateHandler.GetMousePosition()
    End Function

    Public Function IsKeyDown(key As Keys) As Boolean
        Return _inputStateHandler.IsKeyDown(key)
    End Function

    Public Function IsKeyHeld(key As Keys) As Boolean
        Return _inputStateHandler.IsKeyHeld(key)
    End Function

    Public Function IsKeyUp(key As Keys) As Boolean
        Return _inputStateHandler.IsKeyUp(key)
    End Function

    Public Function IsMouseDown(button As MouseButtons) As Boolean
        Return _inputStateHandler.IsMouseDown(button)
    End Function

    Public Function IsMouseHeld(button As MouseButtons) As Boolean
        Return _inputStateHandler.IsMouseHeld(button)
    End Function

    Public Function IsMouseUp(button As MouseButtons) As Boolean
        Return _inputStateHandler.IsMouseUp(button)
    End Function
#End Region

    Private Sub InitializeScriptEngine()
        _scriptEngine.AddHostTypes(
            GetType(Vec2i), GetType(Vec2f), GetType(Recti), GetType(Rectf), GetType(MusicAsset),
            GetType(ImageAsset), GetType(SoundAsset), GetType(FontAsset), GetType(AppMain),
            GetType(Keys), GetType(MouseButtons), GetType(Color), GetType(FontStyle),
            GetType(Strings), GetType(Convert), GetType(VBMath), GetType(MathF)
        )
        _scriptEngine.AddHostObject("App", Me)
    End Sub

    Private Sub StopEngineWithMessage(message As String)
        MessageBox.Show(message, "SCRIPT ERROR - Engine will be stopped",
            MessageBoxButtons.OK, MessageBoxIcon.Error)
        Close()
    End Sub

    Private Sub LoadGameScript()
        Try
            If File.Exists("gamemain.vbs") Then
                Dim scriptCode As String = File.ReadAllText("gamemain.vbs")
                _scriptEngine.Execute(scriptCode)
                _scriptEngine.Invoke("Initialize")
            End If
        Catch ex As Exception
            StopEngineWithMessage($"Error loading script: {ex.Message}")
        End Try
    End Sub

    Private Sub StartGameLoop()
        _gameLoop = New System.Windows.Forms.Timer With {.Interval = 16}
        AddHandler _gameLoop.Tick, AddressOf GameLoop_Tick
        _gameLoop.Start()
    End Sub

    Private Sub GameLoop_Tick(sender As Object, e As EventArgs)
        ' Calculate delta time in seconds
        Dim currentTime As Long = Date.Now.Ticks \ TimeSpan.TicksPerMillisecond
        If _lastFrameTime = 0 Then
            _deltaTime = 0.016F ' Assume 60fps on first frame
        Else
            _deltaTime = (currentTime - _lastFrameTime) / 1000.0F
        End If
        _lastFrameTime = currentTime

        _inputStateHandler.UpdatePreviousStates()
        Try
            _scriptEngine.Invoke("Update", _deltaTime)
        Catch ex As Exception
            StopEngineWithMessage($"Error in Update call: {ex.Message}")
        End Try
        Invalidate()
    End Sub

    Private Sub AppMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        _inputStateHandler.SetKeyState(e.KeyCode, True)
    End Sub

    Private Sub AppMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        _inputStateHandler.SetKeyState(e.KeyCode, False)
    End Sub

    Private Sub AppMain_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        _inputStateHandler.SetMouseState(e.Button, True)
    End Sub

    Private Sub AppMain_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        _inputStateHandler.SetMouseState(e.Button, False)
    End Sub

    Private Sub AppMain_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        _inputStateHandler.SetMousePosition(e.X, e.Y)
    End Sub

    Private Sub AppMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        _gameLoop?.Stop()
        _inputStateHandler.ResetStates()
        _scriptEngine.Dispose()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        _graphics = e.Graphics
        Try
            _scriptEngine.Invoke("Render", _graphics, _deltaTime)
        Catch ex As Exception
            StopEngineWithMessage($"Error in Render call: {ex.Message}")
        End Try
    End Sub

    <STAThread()>
    Friend Shared Sub Main()
        Application.SetHighDpiMode(HighDpiMode.DpiUnaware)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(Instance)
    End Sub
End Class