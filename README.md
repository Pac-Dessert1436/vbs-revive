# vbs-revive: VBScript Game Engine Created With VB.NET WinForms

> **Important Note About VBScript Deprecation**
> 
> Microsoft has officially deprecated VBScript and will remove it from future versions of Windows, which is why this project exists (see [Why VBScript?](#why-vbscript) section). 
> 
> The project's executable (`vbs-revive.exe`) currently relies on the system's `vbscript.dll` to run game scripts. If you're concerned about long-term compatibility, you can back up the following files from your current Windows installation (Windows 10/11):
> 
> - `C:\Windows\System32\vbscript.dll` (64-bit)
> - `C:\Windows\SysWOW64\vbscript.dll` (32-bit)
> - Registry key: `HKEY_CLASSES_ROOT\VBScript`

![](screenshot.png)

## Description
`vbs-revive` is a lightweight game engine that allows developers to create games using VBScript as the scripting language. Despite Microsoft's official deprecation of VBScript, this engine remains a viable option for game development to this day.

The game engine, built with VB.NET and Windows Forms, not only provides a simple interface for rapid game prototyping, as well as for learning game development concepts, but also leverages the power of VBScript to create dynamic and interactive games, while remaining easy to use and extend.

> **v1.0.1 Stable Release**: Added music/sound looping support, along with LINQ-like functionality for arrays and collections. For more details, see [Available APIs](#available-apis).

## Why VBScript?
Yes, Microsoft is deprecating VBScript - that's exactly why this project exists. VBScript was once the easiest way for beginners to automate Windows and make simple games. _It required no compilers, no complex toolchains - just Notepad and a double-click. But as Microsoft moves on, that simplicity risks being lost forever._

`vbs-revive` is not a practical choice for production games. It's a preservation project, a hobbyist's playground, a mighty proof that even a "dead" language can still power a game loop, handle input, play sounds, and bring a little joy to those who grew up with it.

**Switching to Lua or JavaScript would make the engine more "viable," but it would also defeat the point.** Therefore, this project isn't about the best tool for the job. It's about keeping an old friend alive a little longer.

_**They said VBScript was deprecated, and they said it wouldn't run any longer, but the game engine has VBScript anyway.**_

## Features

- **VBScript-based scripting**: Write your game logic in familiar VBScript
- **Asset management**: Support for images, sounds, music, and fonts
- **Input handling**: Keyboard and mouse input support with state tracking
- **Graphics rendering**: 2D graphics with drawing primitives
- **Mathematical utilities**: Vector and rectangle classes for game geometry
- **Audio playback**: Support for sound effects and background music
- **Cross-platform dependencies**: Uses .NET 10.0 for modern compatibility
- **New Collections API**: Provides a rich set of collection classes for handling object arrays/hashsets, `ArrayList` and `Hashtable` with LINQ support (see [Collections Module](#collections-module-linq-like-functionality-new-in-101) section for details)
- **Extensibility**: Easily add custom game components and features

## Architecture Overview

The engine consists of several key components:

### Core Components
- **AppMain.vb**: The main application class that handles the game window, script engine, and game loop
- **InputStateHandler.vb**: Tracks keyboard and mouse states with support for key down, held, and up events
- **EngineAssets.vb**: Asset management classes for sounds, music, fonts, and images
- **Geometrics.vb**: Mathematical classes for 2D vectors and rectangles

### Asset Classes
- **`SoundAsset`**: Handles WAV audio files in the `sounds/` directory
- **`MusicAsset`**: Manages MP3 background music in the `music/` directory
- **`FontAsset`**: Loads TTF fonts from the `fonts/` directory
- **`ImageAsset`**: Supports PNG images in the `images/` directory

### Geometry Classes
- **`Vec2i` / `Vec2f`**: 2D vector classes for integer and floating-point coordinates
- **`Recti` / `Rectf`**: Rectangle classes with methods for collision detection and drawing

## Getting Started
```bash
git clone https://github.com/Pac-Dessert1436/vbs-revive.git
cd vbs-revive
```

1. **Prerequisites**: Ensure you have [.NET 10.0](https://dotnet.microsoft.com/download/dotnet/10.0) or later installed
2. **Build**: Restore dependencies with `dotnet restore` command, then compile the project with `dotnet build` command
3. **Run**: Execute the application using `dotnet run` command, or click the `vbs-revive.exe` file in the `bin/Debug/net10.0` directory
4. **Publish**: To create a standalone executable, use `dotnet publish` command

Upon first run, the engine creates default asset directories (`sounds`, `images`, `fonts`, `music`) and generates a sample `gamemain.vbs` script.

## Game Scripting API

> **Note**: `!` at the end of a variable name means it is a `Single` type. For more details, see [Available APIs](#available-apis) section.

Your game logic goes in `gamemain.vbs` with three main functions:

- `Initialize()`: Called once at startup
- `Update(dt!)`: Called every frame for game logic, where `dt` is delta time in seconds since last frame
- `Render(g As Graphics, dt!)`: Called every frame for drawing, where `g` is the Graphics object and `dt` is delta time in seconds since last frame

### Example Script
```vbscript
Option Explicit
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
    Set font = FontAsset.Create("Arial")
    text = "Welcome! Use W,A,S,D or Arrow Keys to move around."
    font.DrawText g, text, 10, 10, Color.Cyan
    fmtDt = Strings.Format(dt * 1000, "0.00")
    text = "Speed: " & playerSpeed & " px/s; Delta Time: " & fmtDt & " ms"
    font.DrawText g, text, 10, 30, Color.Cyan
End Sub
```

### Available APIs

> **Note**: Shorthands for writing type annotations: `%` = `As Integer`, `!` = `As Single`, `$` = `As String`. Parameters without type annotations are `Object` type.

#### Engine API (AppMain.Instance)
- `SetWindowSize(width%, height%)`: Set the game window size
- `GetMousePosition()`: Get current mouse position as Vec2i
- `IsKeyDown(key As Keys)`: Check if a key was just pressed; returns `Boolean`
- `IsKeyHeld(key As Keys)`: Check if a key is currently held; returns `Boolean`
- `IsKeyUp(key As Keys)`: Check if a key was just released; returns `Boolean`
- `IsMouseDown(button As MouseButtons)`: Check if a mouse button was just pressed; returns `Boolean`
- `IsMouseHeld(button As MouseButtons)`: Check if a mouse button is currently held; returns `Boolean`
- `IsMouseUp(button As MouseButtons)`: Check if a mouse button was just released; returns `Boolean`

#### Asset Classes

##### `SoundAsset` - WAV Audio Asset
- `SoundAsset.Create(name$)`: Create a sound asset from sounds/name.wav
- `.Play()`: Play the sound
- `.PlayLooping()`: Play the sound in a continuous loop (available in 1.0.1)
- `.Stop()`: Stop the sound playback

##### `MusicAsset` - MP3 Audio Asset
- `MusicAsset.Create(name$)`: Create a music asset from music/name.mp3
- `.Play()`: Play the music
- `.PlayLooping()`: Play the music in a continuous loop (available in 1.0.1)
- `.Stop()`: Stop the music playback
- `.SetVolume(volume)`: Set volume level (0.0 to 1.0)

##### `ImageAsset` - PNG Image Asset
- `ImageAsset.Create(name$)`: Create an image asset from images/name.png
- `.Draw(g As Graphics, x%, y%)`: Draw the image at position (x, y) on the Graphics object
- `.Width`: Get the image width (read-only property)
- `.Height`: Get the image height (read-only property)

##### `FontAsset` - TrueType Font Asset
- `FontAsset.Create(name$)`: Create a font asset from fonts/name.ttf (or use system font if not found)
- `.DrawText(g As Graphics, text$, x%, y%, [color As Color], [size%])`: Draw text at position (x, y) with optional color and size

#### Geometric Types

##### `Vec2i` - Integer 2D Vector
- `Vec2i.Create(x%, y%)`: Create a new Vec2i instance
- `Vec2i.CreateZero()`: Create a zero vector (0, 0)
- `Vec2i.CreateOne()`: Create a unit vector (1, 1)
- `Vec2i.CreateUnitX()`: Create unit X vector (1, 0)
- `Vec2i.CreateUnitY()`: Create unit Y vector (0, 1)
- `.X`, `.Y`: Coordinate properties
- `.Add(other As Vec2i)`: Add another vector and return new vector
- `.Subtract(other As Vec2i)`: Subtract another vector and return new vector
- `.Multiply(scalar%)`: Multiply by scalar and return new vector
- `.Divide(scalar%)`: Divide by scalar and return new vector
- `.Dot(other As Vec2i)`: Calculate dot product with another vector
- `.LengthSquared()`: Get squared length of the vector
- `.Length()`: Get length of the vector
- `.Normalize()`: Get normalized vector (unit vector)
- `.DistanceTo(other As Vec2i)`: Calculate distance to another vector
- `.DistanceSquaredTo(other As Vec2i)`: Calculate squared distance to another vector
- `.ToString()`: Get string representation

##### `Vec2f` - Float 2D Vector
- `Vec2f.Create(x!, y!)`: Create a new Vec2f instance
- `Vec2f.CreateZero()`: Create a zero vector (0.0, 0.0)
- `Vec2f.CreateOne()`: Create a unit vector (1.0, 1.0)
- `Vec2f.CreateUnitX()`: Create unit X vector (1.0, 0.0)
- `Vec2f.CreateUnitY()`: Create unit Y vector (0.0, 1.0)
- `.X`, `.Y`: Coordinate properties
- `.Add(other As Vec2f)`: Add another vector and return new vector
- `.Subtract(other As Vec2f)`: Subtract another vector and return new vector
- `.Multiply(scalar!)`: Multiply by scalar and return new vector
- `.Divide(scalar!)`: Divide by scalar and return new vector
- `.Dot(other As Vec2f)`: Calculate dot product with another vector
- `.Cross(other As Vec2f)`: Calculate cross product with another vector
- `.LengthSquared()`: Get squared length of the vector
- `.Length()`: Get length of the vector
- `.Normalize()`: Get normalized vector (unit vector)
- `.DistanceTo(other As Vec2f)`: Calculate distance to another vector
- `.DistanceSquaredTo(other As Vec2f)`: Calculate squared distance to another vector
- `.Angle()`: Get angle of the vector
- `.AngleTo(other As Vec2f)`: Get angle to another vector
- `.Rotate(angle!)`: Rotate vector by angle in radians
- `.ToString()`: Get string representation

##### `Recti` - Integer Rectangle
- `Recti.Create(x%, y%, width%, height%)`: Create a new Recti instance
- `Recti.CreateFromPoints(topLeft As Vec2i, bottomRight As Vec2i)`: Create from top-left and bottom-right points
- `Recti.CreateFromCenter(center As Vec2i, width%, height%)`: Create from center point and dimensions
- `.X`, `.Y`, `.Width`, `.Height`: Rectangle properties
- `.Right`, `.Bottom`: Read-only properties for right and bottom edges
- `.CenterX`, `.CenterY`: Read-only properties for center coordinates
- `.Center`: Read-only property for center as Vec2i
- `.Offset(dx%, dy%)`: Move rectangle by offset values
- `.Offset(vec As Vec2i)`: Move rectangle by Vec2i offset
- `.Inflate(amount%)`: Expand rectangle by amount in all directions
- `.Inflate(horizontal%, vertical%)`: Expand rectangle by different horizontal/vertical amounts
- `.Contains(x%, y%)`: Check if point is inside rectangle
- `.Contains(point As Vec2i)`: Check if Vec2i point is inside rectangle
- `.Contains(rect As Recti)`: Check if another rectangle is completely inside
- `.Intersects(other As Recti)`: Check if rectangle intersects with another
- `.Intersection(other As Recti)`: Get intersection rectangle with another
- `.Union(other As Recti)`: Get union rectangle with another
- `.DrawOutline(g As Graphics, thickness%, [color As Color])`: Draw rectangle outline on Graphics object
- `.DrawFilled(g As Graphics, [color As Color])`: Draw filled rectangle on Graphics object

##### `Rectf` - Float Rectangle
- `Rectf.Create(x!, y!, width!, height!)`: Create a new Rectf instance
- `Rectf.CreateFromPoints(topLeft As Vec2f, bottomRight As Vec2f)`: Create from top-left and bottom-right points
- `Rectf.CreateFromCenter(center As Vec2f, width!, height!)`: Create from center point and dimensions
- `.X`, `.Y`, `.Width`, `.Height`: Rectangle properties
- `.Right`, `.Bottom`: Read-only properties for right and bottom edges
- `.CenterX`, `.CenterY`: Read-only properties for center coordinates
- `.Center`: Read-only property for center as Vec2f
- `.Offset(dx!, dy!)`: Move rectangle by offset values
- `.Offset(vec As Vec2f)`: Move rectangle by Vec2f offset
- `.Inflate(amount!)`: Expand rectangle by amount in all directions
- `.Inflate(horizontal!, vertical!)`: Expand rectangle by different horizontal/vertical amounts
- `.Contains(x!, y!)`: Check if point is inside rectangle
- `.Contains(point As Vec2f)`: Check if Vec2f point is inside rectangle
- `.Contains(rect As Rectf)`: Check if another rectangle is completely inside
- `.Intersects(other As Rectf)`: Check if rectangle intersects with another
- `.Intersection(other As Rectf)`: Get intersection rectangle with another
- `.Union(other As Rectf)`: Get union rectangle with another
- `.DrawOutline(g As Graphics, thickness%, [color As Color])`: Draw rectangle outline on Graphics object
- `.DrawFilled(g As Graphics, [color As Color])`: Draw filled rectangle on Graphics object

#### Collections Module (LINQ-like functionality, NEW in 1.0.1)

The `Collections` module provides LINQ-like methods for working with arrays and collections in VBScript. All methods work with VBScript arrays and .NET collections. _Unlike the `AddressOf` keyword in VB.NET, we use **the built-in `GetRef` function** in VBScript with the function name, to transform functions into first-class objects._

**Creation Methods:**
- `Collections.ArrayOf(ParamArray items)`: Create an array from parameters
- `Collections.CreateArrayList([capacity%])`: Create an empty ArrayList (optional initial capacity)
- `Collections.CreateHashSet()`: Create an empty HashSet
- `Collections.CreateHashtable()`: Create an empty Hashtable

**Query Methods:**
- `Collections.Where(source, predicate)`: Filter items using a predicate function. Returns ArrayList.
- `Collections.Select(source, selector)`: Transform items using a selector function. Returns ArrayList.
- `Collections.First(source, [predicate])`: Get first matching item, or first item if no predicate.
- `Collections.Last(source, [predicate])`: Get last matching item, or last item if no predicate.
- `Collections.CountItems(source, [predicate])`: Count items, optionally filtered by predicate.
- `Collections.Any(source, [predicate])`: Check if any item exists (or matches predicate). Returns Boolean.
- `Collections.All(source, predicate)`: Check if all items match predicate. Returns Boolean.
- `Collections.Contains(source, value)`: Check if collection contains specified value. Returns Boolean.

**Manipulation Methods:**
- `Collections.Sort(source)`: Sort an ArrayList in place. Returns the sorted ArrayList.
- `Collections.SortWithComparer(source, comparer)`: Sort ArrayList with custom comparison function. Returns sorted ArrayList.
- `Collections.Reverse(source)`: Reverse an ArrayList in place. Returns the reversed ArrayList.
- `Collections.Take(source, count%)`: Take first N items from collection. Returns ArrayList.
- `Collections.Skip(source, count%)`: Skip first N items and return the rest. Returns ArrayList.

**Conversion Methods:**
- `Collections.ToArray(source)`: Convert collection to VBScript Array.
- `Collections.ToArrayList(source)`: Convert collection to ArrayList.

**Usage Example:**
```vbscript
' Create an array of enemies
enemies = Collections.ArrayOf( _
    Array("Goblin", 10), _
    Array("Orc", 25), _
    Array("Dragon", 100) _
)
' Filter strong enemies (HP > 20)
strongEnemies = Collections.Where(enemies, GetRef("FilterStrong"))
' Get enemy names only
names = Collections.Select(enemies, GetRef("GetName"))

' Helper functions
Function FilterStrong(enemy)
    FilterStrong = (enemy(1) > 20)
End Function

Function GetName(enemy)
    GetName = enemy(0)
End Function
```

#### Native .NET Framework Types Exposed

##### `Keys` (from namespace System.Windows.Forms)
Enum containing all keyboard keys that can be used with input functions.

##### `MouseButtons` (from namespace System.Windows.Forms)
Enum containing mouse buttons (Left, Right, Middle) for mouse input functions.

##### `Color` (from namespace System.Drawing)
Provides predefined colors (Color.Red, Color.Blue, etc.) and color creation methods.
- Predefined colors: Black, White, Red, Green, Blue, Cyan, Magenta, Yellow, etc.
- Methods: `Color.FromArgb(r%, g%, b%)`, `Color.FromArgb(a%, r%, g%, b%)`

##### `FontStyle` (from namespace System.Drawing)
Font styles for text rendering (Regular, Bold, Italic, Underline, Strikeout).

##### `Strings` (from namespace Microsoft.VisualBasic)
VB-specific string manipulation functions:
- `Strings.Format(expression$, format$)`: Format a value with specified format
- `Strings.Len(string$)`: Get length of string
- And many more string functions

##### `Convert` (from namespace System.Convert)
Conversion functions between different data types:
- `Convert.ToInt32(value)`: Converts an object to 32-bit integer
- `Convert.ToSingle(value)`: Converts an object to single precision float
- `Convert.ToDouble(value)`: Converts an object to double precision float
- `Convert.ToString(value)`: Converts an object to string
- And many more conversion methods

##### `VBMath` (from namespace Microsoft.VisualBasic)
Legacy mathematical functions, especially for random number generation:
- `VBMath.Rnd()`: Generate random number
- `VBMath.Randomize()`: Initialize random number generator

##### `MathF` (from namespace System.Math, used for floating-point math)
Mathematical functions for floating-point numbers:
- `MathF.Sin(value!)`: Sine function
- `MathF.Cos(value!)`: Cosine function
- `MathF.Tan(value!)`: Tangent function
- `MathF.Sqrt(value!)`: Square root
- `MathF.Abs(value!)`: Absolute value
- `MathF.Pow(base!, exponent!)`: Power function
- `MathF.Log(value!)`: Natural logarithm
- `MathF.Max(val1!, val2!)`: Maximum of two values
- `MathF.Min(val1!, val2!)`: Minimum of two values
- And many more mathematical functions

## Dependencies

- **Microsoft ClearScript**: For VBScript integration
- **NAudio**: For audio playback capabilities
- **System.Drawing**: For graphics rendering
- **Windows Forms**: For UI elements

## Project Structure

- `AppMain.vb`: Main application and game loop
- `EngineAssets.vb`: Asset management classes
- `Geometrics.vb`: Vector and rectangle classes
- `InputStateHandler.vb`: Input state management
- `vbs-revive.vbproj`: Project configuration

## License
This project is licensed under the BSD-3-Clause License. See the [LICENSE](LICENSE) file for details.