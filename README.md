# VB6 VideoPlayer

VB6 VideoPlayer is a simple, lightweight VB6 class that connects to the native Windows video subsystem and plays video directly onto a window.
It is intended for use in classic VB6 applications that need in-place video rendering without extra chrome or container frames.
The class only supports video formats that the host Windows installation natively supports (system codecs / DirectShow / Media Foundation).
Attempting to open unsupported files will raise an error.

---

## Key points / Philosophy

- Minimal, VB6-friendly class-object API.
- Plays video directly into any window.
- Uses only native Windows video support — no bundled codec or transcoding layers.

---

## Requirements

- Visual Basic 6.
- Windows with native video support (system codecs, DirectShow, or Media Foundation).

---

## Limitations

- Only plays files supported natively by the Windows installation. Unsupported files will cause an error when opened.
- No built-in controls (play/pause buttons, progress bars). The class provides programmatic control only.
- Not a full media framework — no subtitle parsing, no advanced audio routing, no transcoding, no format conversion.

---

## Interface

Datatypes referenced:
- VideoSize: same layout as a RECT structure (left, top, width, height)
- Multimedia_Command: enumeration of available commands: Open, Close, Play, Pause, Stop, Rewind

### Properties

- **hWndParent** (Long)  
  The window handle (hWnd) of the parent window where the video will be rendered.

- **FileName** (String)  
  A full path or URL to a video file. The file must be supported by the Windows installation.

- **VideoSize** (VideoSize datatype)  
  Read-Only. The rectangle that defines the video position/size relative to the parent window.

- **Left**, **Top**, **Width**, **Height** (Long)  
  After changing any of these values you must call ResizeVideo() to apply them.

- **Length** (Long)  
  Read-only. Returns the length of the currently opened video in seconds. If no file is opened, the value will be zero or unavailable.

- **Position** (Long)  
  Get or set the current playback position in seconds. Setting Position seeks the playback; getting returns the current position in seconds.

- **Stretch** (Boolean)  
  If True, the video will be stretched to exactly the dimensions you provided. If False, the player will preserve the video's aspect ratio while attempting to fit inside the assigned rectangle.

- **Command** (Multimedia_Command datatype)  
  Set this property to issue one of the available commands to the player. Valid commands include:
  - Open — open the file set in FileName (valid file required; Open will raise an error for unsupported files)
  - Play — start or resume playback
  - Pause — pause playback
  - Stop — stop playback
  - Rewind — seek to the beginning.

Important behavior note: a file must be opened (Command = Open) before playback actions are valid.
Open will raise a runtime error if the FileName refers to an unsupported format.

### Methods

- **ResizeVideo**()  
  After modifying Left/Top/Width/Height properties, call ResizeVideo() to apply the new rectangle to the running video surface.
  Changing the geometry properties alone has no effect until this method is invoked.

---

## License

[MIT License](LICENSE)  
Copyright © Ubehage

---

## Credits

Created by Ubehage  
[GitHub Profile](https://github.com/Ubehage)