Imports NAudio.Wave
Imports System.Drawing.Text
Imports System.IO

Public NotInheritable Class SoundAsset
    Private ReadOnly _fileName As String
    Private _waveOut As WaveOutEvent = Nothing
    Private _audioFile As AudioFileReader = Nothing
    Private _isLooping As Boolean = False

    Private Sub New(assetName As String)
        _fileName = $"sounds/{assetName}.wav"
    End Sub

    Public Shared Function Create(assetName As String) As SoundAsset
        Return New SoundAsset(assetName)
    End Function

    Public Sub Play()
        If File.Exists(_fileName) Then
            StopCurrentPlayback()
            _audioFile = New AudioFileReader(_fileName)
            _waveOut = New WaveOutEvent()
            AddHandler _waveOut.PlaybackStopped, AddressOf OnPlaybackStopped
            _waveOut.Init(_audioFile)
            _waveOut.Play()
        End If
    End Sub

    Public Sub PlayLooping()
        _isLooping = True
        Play()
    End Sub

    Private Sub OnPlaybackStopped(sender As Object, e As StoppedEventArgs)
        If _isLooping Then
            ' Restart the audio file from the beginning for looping
            If File.Exists(_fileName) Then
                _audioFile = New AudioFileReader(_fileName)
                _waveOut = New WaveOutEvent()
                AddHandler _waveOut.PlaybackStopped, AddressOf OnPlaybackStopped
                _waveOut.Init(_audioFile)
                _waveOut.Play()
            End If
        End If
    End Sub

    Public Sub [Stop]()
        _isLooping = False
        StopCurrentPlayback()
    End Sub

    Private Sub StopCurrentPlayback()
        _waveOut?.Stop()
        _audioFile?.Dispose()
        _waveOut?.Dispose()
        _waveOut = Nothing
        _audioFile = Nothing
    End Sub
End Class

Public NotInheritable Class MusicAsset
    Private ReadOnly _fileName As String
    Private _waveOut As WaveOutEvent = Nothing
    Private _audioFile As AudioFileReader = Nothing

    Private Sub New(assetName As String)
        _fileName = $"music/{assetName}.mp3"
    End Sub

    Public Shared Function Create(assetName As String) As MusicAsset
        Return New MusicAsset(assetName)
    End Function

    Private _isLooping As Boolean = False

    Public Sub Play()
        If File.Exists(_fileName) Then
            StopCurrentPlayback()
            _audioFile = New AudioFileReader(_fileName)
            _waveOut = New WaveOutEvent()
            AddHandler _waveOut.PlaybackStopped, AddressOf OnPlaybackStopped
            _waveOut.Init(_audioFile)
            _waveOut.Play()
        End If
    End Sub

    Public Sub PlayLooping()
        _isLooping = True
        Play()
    End Sub

    Private Sub OnPlaybackStopped(sender As Object, e As StoppedEventArgs)
        If _isLooping Then
            ' Restart the audio file from the beginning for looping
            If File.Exists(_fileName) Then
                _audioFile = New AudioFileReader(_fileName)
                _waveOut = New WaveOutEvent()
                AddHandler _waveOut.PlaybackStopped, AddressOf OnPlaybackStopped
                _waveOut.Init(_audioFile)
                _waveOut.Play()
            End If
        End If
    End Sub

    Public Sub [Stop]()
        _isLooping = False
        StopCurrentPlayback()
    End Sub

    Private Sub StopCurrentPlayback()
        _waveOut?.Stop()
        _audioFile?.Dispose()
        _waveOut?.Dispose()
        _waveOut = Nothing
        _audioFile = Nothing
    End Sub

    Public Sub SetVolume(volume As Single)
        If _audioFile IsNot Nothing Then
            _audioFile.Volume = Math.Clamp(volume, 0.0F, 1.0F)
        End If
    End Sub
End Class

Public NotInheritable Class FontAsset
    Private ReadOnly _fileName As String

    Private Sub New(assetName As String)
        _fileName = $"fonts/{assetName}.ttf"
    End Sub

    Public Shared Function Create(assetName As String) As FontAsset
        Return New FontAsset(assetName)
    End Function

    Private ReadOnly Property Font(size As Single) As Font
        Get
            If File.Exists(_fileName) Then
                Dim pfc As New PrivateFontCollection
                pfc.AddFontFile(_fileName)
                Return New Font(pfc.Families(0), size)
            Else
                Return New Font("Arial", size)
            End If
        End Get
    End Property

    Public Sub DrawText(g As Graphics, text As String, x As Integer, y As Integer,
            Optional color As Color? = Nothing, Optional size As Single = 10)
        Using brush = If(color.HasValue, New SolidBrush(color.Value), Brushes.White)
            g.DrawString(text, Font(size), brush, x, y)
        End Using
    End Sub
End Class

Public NotInheritable Class ImageAsset
    Private ReadOnly _fileName As String
    Private _image As Image = Nothing

    Private Sub New(assetName As String)
        _fileName = $"images/{assetName}.png"
    End Sub

    Public Shared Function Create(assetName As String) As ImageAsset
        Return New ImageAsset(assetName)
    End Function

    Private ReadOnly Property Image As Image
        Get
            If _image Is Nothing AndAlso File.Exists(_fileName) Then
                _image = Image.FromFile(_fileName)
            End If
            Return _image
        End Get
    End Property

    Public ReadOnly Property Width As Integer
        Get
            Return If(Image IsNot Nothing, Image.Width, 0)
        End Get
    End Property

    Public ReadOnly Property Height As Integer
        Get
            Return If(Image IsNot Nothing, Image.Height, 0)
        End Get
    End Property

    Public Sub Draw(g As Graphics, x As Integer, y As Integer)
        g.DrawImage(Image, x, y)
    End Sub
End Class