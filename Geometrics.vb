Public NotInheritable Class Vec2i
    Implements ICloneable

    Public Property X As Integer
    Public Property Y As Integer

    Public Shared Function Create(x As Integer, y As Integer) As Vec2i
        Return New Vec2i With {.X = x, .Y = y}
    End Function

    Public Shared Function CreateZero() As Vec2i
        Return New Vec2i With {.X = 0, .Y = 0}
    End Function

    Public Shared Function CreateOne() As Vec2i
        Return New Vec2i With {.X = 1, .Y = 1}
    End Function

    Public Shared Function CreateUnitX() As Vec2i
        Return New Vec2i With {.X = 1, .Y = 0}
    End Function

    Public Shared Function CreateUnitY() As Vec2i
        Return New Vec2i With {.X = 0, .Y = 1}
    End Function

    Public Function Add(other As Vec2i) As Vec2i
        Return Create(X + other.X, Y + other.Y)
    End Function

    Public Function Subtract(other As Vec2i) As Vec2i
        Return Create(X - other.X, Y - other.Y)
    End Function

    Public Function Multiply(scalar As Integer) As Vec2i
        Return Create(X * scalar, Y * scalar)
    End Function

    Public Function Divide(scalar As Integer) As Vec2i
        If scalar = 0 Then Return CreateZero()
        Return Create(X \ scalar, Y \ scalar)
    End Function

    Public Function Dot(other As Vec2i) As Integer
        Return X * other.X + Y * other.Y
    End Function

    Public Function LengthSquared() As Integer
        Return X * X + Y * Y
    End Function

    Public Function Length() As Single
        Return MathF.Sqrt(LengthSquared())
    End Function

    Public Function Normalize() As Vec2i
        Dim len As Single = Length()
        If len = 0 Then Return CreateZero()
        Return Create(CInt(X / len), CInt(Y / len))
    End Function

    Public Function DistanceTo(other As Vec2i) As Single
        Return Subtract(other).Length()
    End Function

    Public Function DistanceSquaredTo(other As Vec2i) As Integer
        Return Subtract(other).LengthSquared()
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Return New Vec2i With {.X = X, .Y = Y}
    End Function

    Public Overrides Function ToString() As String
        Return $"({X}, {Y})"
    End Function
End Class

Public NotInheritable Class Vec2f
    Implements ICloneable

    Public Property X As Single
    Public Property Y As Single

    Public Shared Function Create(x As Single, y As Single) As Vec2f
        Return New Vec2f With {.X = x, .Y = y}
    End Function

    Public Shared Function CreateZero() As Vec2f
        Return New Vec2f With {.X = 0.0F, .Y = 0.0F}
    End Function

    Public Shared Function CreateOne() As Vec2f
        Return New Vec2f With {.X = 1.0F, .Y = 1.0F}
    End Function

    Public Shared Function CreateUnitX() As Vec2f
        Return New Vec2f With {.X = 1.0F, .Y = 0.0F}
    End Function

    Public Shared Function CreateUnitY() As Vec2f
        Return New Vec2f With {.X = 0.0F, .Y = 1.0F}
    End Function

    Public Function Add(other As Vec2f) As Vec2f
        Return Create(X + other.X, Y + other.Y)
    End Function

    Public Function Subtract(other As Vec2f) As Vec2f
        Return Create(X - other.X, Y - other.Y)
    End Function

    Public Function Multiply(scalar As Single) As Vec2f
        Return Create(X * scalar, Y * scalar)
    End Function

    Public Function Divide(scalar As Single) As Vec2f
        If scalar = 0 Then Return CreateZero()
        Return Create(X / scalar, Y / scalar)
    End Function

    Public Function Dot(other As Vec2f) As Single
        Return X * other.X + Y * other.Y
    End Function

    Public Function Cross(other As Vec2f) As Single
        Return X * other.Y - Y * other.X
    End Function

    Public Function LengthSquared() As Single
        Return X * X + Y * Y
    End Function

    Public Function Length() As Single
        Return MathF.Sqrt(LengthSquared())
    End Function

    Public Function Normalize() As Vec2f
        Dim len As Single = Length()
        If len = 0 Then Return CreateZero()
        Return Create(X / len, Y / len)
    End Function

    Public Function DistanceTo(other As Vec2f) As Single
        Return Subtract(other).Length()
    End Function

    Public Function DistanceSquaredTo(other As Vec2f) As Single
        Return Subtract(other).LengthSquared()
    End Function

    Public Function Angle() As Single
        Return MathF.Atan2(Y, X)
    End Function

    Public Function AngleTo(other As Vec2f) As Single
        Dim dotValue As Single = Dot(other)
        Dim lenProduct As Single = Length() * other.Length()
        If lenProduct = 0 Then Return 0
        Dim ratio As Single = dotValue / lenProduct
        Return Math.Clamp(MathF.Acos(ratio), -1.0F, 1.0F)
    End Function

    Public Function Rotate(angle As Single) As Vec2f
        Dim cosAngle As Single = MathF.Cos(angle)
        Dim sinAngle As Single = MathF.Sin(angle)
        Return Create(
            X * cosAngle - Y * sinAngle,
            X * sinAngle + Y * cosAngle
        )
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Return New Vec2f With {.X = X, .Y = Y}
    End Function

    Public Overrides Function ToString() As String
        Return $"({X:F2}, {Y:F2})"
    End Function
End Class

Public NotInheritable Class Recti
    Implements ICloneable

    Public Property X As Integer
    Public Property Y As Integer
    Public Property Width As Integer
    Public Property Height As Integer

    Public ReadOnly Property Right As Integer
        Get
            Return X + Width
        End Get
    End Property

    Public ReadOnly Property Bottom As Integer
        Get
            Return Y + Height
        End Get
    End Property

    Public ReadOnly Property CenterX As Integer
        Get
            Return X + Width \ 2
        End Get
    End Property

    Public ReadOnly Property CenterY As Integer
        Get
            Return Y + Height \ 2
        End Get
    End Property

    Public ReadOnly Property Center As Vec2i
        Get
            Return Vec2i.Create(CenterX, CenterY)
        End Get
    End Property

    Public Shared Function Create(x As Integer, y As Integer, width As Integer, height As Integer) As Recti
        Return New Recti With {.X = x, .Y = y, .Width = width, .Height = height}
    End Function

    Public Shared Function CreateFromPoints(topLeft As Vec2i, bottomRight As Vec2i) As Recti
        Return Create(
            topLeft.X,
            topLeft.Y,
            bottomRight.X - topLeft.X,
            bottomRight.Y - topLeft.Y
        )
    End Function

    Public Shared Function CreateFromCenter(center As Vec2i, width As Integer, height As Integer) As Recti
        Return Create(
            center.X - width \ 2,
            center.Y - height \ 2,
            width,
            height
        )
    End Function

    Public Sub Offset(dx As Integer, dy As Integer)
        X += dx
        Y += dy
    End Sub

    Public Sub Offset(offset As Vec2i)
        X += offset.X
        Y += offset.Y
    End Sub

    Public Sub Inflate(amount As Integer)
        X -= amount
        Y -= amount
        Width += amount * 2
        Height += amount * 2
    End Sub

    Public Sub Inflate(horizontal As Integer, vertical As Integer)
        X -= horizontal
        Y -= vertical
        Width += horizontal * 2
        Height += vertical * 2
    End Sub

    Public Function Contains(x As Integer, y As Integer) As Boolean
        Return x >= x AndAlso x < Right AndAlso y >= y AndAlso y < Bottom
    End Function

    Public Function Contains(point As Vec2i) As Boolean
        Return Contains(point.X, point.Y)
    End Function

    Public Function Contains(rect As Recti) As Boolean
        Return rect.X >= X AndAlso rect.Right <= Right AndAlso rect.Y >= Y AndAlso rect.Bottom <= Bottom
    End Function

    Public Function Intersects(other As Recti) As Boolean
        Return X < other.Right AndAlso Right > other.X AndAlso Y < other.Bottom AndAlso Bottom > other.Y
    End Function

    Public Function Intersection(other As Recti) As Recti
        If Not Intersects(other) Then Return Create(0, 0, 0, 0)

        Dim interX As Integer = Math.Max(X, other.X)
        Dim interY As Integer = Math.Max(Y, other.Y)
        Dim interWidth As Integer = Math.Min(Right, other.Right) - interX
        Dim interHeight As Integer = Math.Min(Bottom, other.Bottom) - interY

        Return Create(interX, interY, interWidth, interHeight)
    End Function

    Public Function Union(other As Recti) As Recti
        Dim unionX As Integer = Math.Min(X, other.X)
        Dim unionY As Integer = Math.Min(Y, other.Y)
        Dim unionWidth As Integer = Math.Max(Right, other.Right) - unionX
        Dim unionHeight As Integer = Math.Max(Bottom, other.Bottom) - unionY

        Return Create(unionX, unionY, unionWidth, unionHeight)
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Return New Recti With {.X = X, .Y = Y, .Width = Width, .Height = Height}
    End Function

    Public Overrides Function ToString() As String
        Return $"Recti: X={X}, Y={Y}, Width={Width}, Height={Height}"
    End Function

    Public Sub DrawOutline(g As Graphics, thickness As Integer, Optional penColor As Color? = Nothing)
        If thickness <= 0 Then Throw New ArgumentException("Thickness must be greater than 0.")
        Using pen = If(penColor.HasValue, New Pen(penColor.Value, thickness), New Pen(Color.White, thickness))
            g.DrawRectangle(pen, X, Y, Width, Height)
        End Using
    End Sub

    Public Sub DrawFilled(g As Graphics, Optional color As Color? = Nothing)
        Using brush = If(color.HasValue, New SolidBrush(color.Value), Brushes.White)
            g.FillRectangle(brush, X, Y, Width, Height)
        End Using
    End Sub
End Class

Public NotInheritable Class Rectf
    Implements ICloneable

    Public Property X As Single
    Public Property Y As Single
    Public Property Width As Single
    Public Property Height As Single

    Public ReadOnly Property Right As Single
        Get
            Return X + Width
        End Get
    End Property

    Public ReadOnly Property Bottom As Single
        Get
            Return Y + Height
        End Get
    End Property

    Public ReadOnly Property CenterX As Single
        Get
            Return X + Width / 2.0F
        End Get
    End Property

    Public ReadOnly Property CenterY As Single
        Get
            Return Y + Height / 2.0F
        End Get
    End Property

    Public ReadOnly Property Center As Vec2f
        Get
            Return Vec2f.Create(CenterX, CenterY)
        End Get
    End Property

    Public Shared Function Create(x As Single, y As Single, width As Single, height As Single) As Rectf
        Return New Rectf With {.X = x, .Y = y, .Width = width, .Height = height}
    End Function

    Public Shared Function CreateFromPoints(topLeft As Vec2f, bottomRight As Vec2f) As Rectf
        Return Create(
            topLeft.X,
            topLeft.Y,
            bottomRight.X - topLeft.X,
            bottomRight.Y - topLeft.Y
        )
    End Function

    Public Shared Function CreateFromCenter(center As Vec2f, width As Single, height As Single) As Rectf
        Return Create(
            center.X - width / 2.0F,
            center.Y - height / 2.0F,
            width,
            height
        )
    End Function

    Public Sub Offset(dx As Single, dy As Single)
        X += dx
        Y += dy
    End Sub

    Public Sub Offset(offset As Vec2f)
        X += offset.X
        Y += offset.Y
    End Sub

    Public Sub Inflate(amount As Single)
        X -= amount
        Y -= amount
        Width += amount * 2.0F
        Height += amount * 2.0F
    End Sub

    Public Sub Inflate(horizontal As Single, vertical As Single)
        X -= horizontal
        Y -= vertical
        Width += horizontal * 2.0F
        Height += vertical * 2.0F
    End Sub

    Public Function Contains(x As Single, y As Single) As Boolean
        Return x >= x AndAlso x < Right AndAlso y >= y AndAlso y < Bottom
    End Function

    Public Function Contains(point As Vec2f) As Boolean
        Return Contains(point.X, point.Y)
    End Function

    Public Function Contains(rect As Rectf) As Boolean
        Return rect.X >= X AndAlso rect.Right <= Right AndAlso rect.Y >= Y AndAlso rect.Bottom <= Bottom
    End Function

    Public Function Intersects(other As Rectf) As Boolean
        Return X < other.Right AndAlso Right > other.X AndAlso Y < other.Bottom AndAlso Bottom > other.Y
    End Function

    Public Function Intersection(other As Rectf) As Rectf
        If Not Intersects(other) Then Return Create(0, 0, 0, 0)

        Dim interX As Single = Math.Max(X, other.X)
        Dim interY As Single = Math.Max(Y, other.Y)
        Dim interWidth As Single = Math.Min(Right, other.Right) - interX
        Dim interHeight As Single = Math.Min(Bottom, other.Bottom) - interY

        Return Create(interX, interY, interWidth, interHeight)
    End Function

    Public Function Union(other As Rectf) As Rectf
        Dim unionX As Single = Math.Min(X, other.X)
        Dim unionY As Single = Math.Min(Y, other.Y)
        Dim unionWidth As Single = Math.Max(Right, other.Right) - unionX
        Dim unionHeight As Single = Math.Max(Bottom, other.Bottom) - unionY

        Return Create(unionX, unionY, unionWidth, unionHeight)
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Return New Rectf With {.X = X, .Y = Y, .Width = Width, .Height = Height}
    End Function

    Public Overrides Function ToString() As String
        Return $"Rectf: X={X:F2}, Y={Y:F2}, Width={Width:F2}, Height={Height:F2}"
    End Function

    Public Sub DrawOutline(g As Graphics, thickness As Integer, Optional penColor As Color? = Nothing)
        If thickness <= 0 Then Throw New ArgumentException("Thickness must be greater than 0.")
        Using pen = If(penColor.HasValue, New Pen(penColor.Value, thickness), New Pen(Color.White, thickness))
            g.DrawRectangle(pen, X, Y, Width, Height)
        End Using
    End Sub

    Public Sub DrawFilled(g As Graphics, Optional color As Color? = Nothing)
        Using brush = If(color.HasValue, New SolidBrush(color.Value), Brushes.White)
            g.FillRectangle(brush, X, Y, Width, Height)
        End Using
    End Sub
End Class