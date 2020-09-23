<div align="center">

## An Idle Check


</div>

### Description

Idle Check tests your system whether it is in an idle state. Once the system is idle for a specified amount of time it performs a certain function. After it stops being idle a further further procedure is called. (Perfect for screensavers). This is done through checking any mouse cursor movements and any key presses.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Soluch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/soluch.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/soluch-an-idle-check__1-44355/archive/master.zip)

### API Declarations

```
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
```


### Source Code

```
Option Explicit
'BEFORE YOU BEGIN:
'- Place timer control on an empty form, name it "TimerIdle"
'- Set the interval on the timer to 1 (one)
'- Copy this code into the form
'- Ensure you can see the Immediate Window to see results
'Note: No error control, insert if you like
'    May encounter problems if computer passes midnight (timer resets)
'Peter Soluch - 2003
'Function to get state of keys
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Function to get position of mouse cursor
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'The time (in seconds) a computer must be idle before running sub
Private Const IDLESECONDS As Long = 5
'Type used with GetCursorPos
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Sub TimerIdle_Timer()
  Dim newMousePos As POINTAPI   'Var for values of "current" Mouse Position
  'Static variables
  Static oldMousePos As POINTAPI 'Old / Previous values of the mouse position
  Static isIdle As Boolean    'Checks if state is currently idle
  Static wasIdle As Boolean    'Checks if state was "declared" idle before
  Static idleStartTime As Single 'When did the idle first start
  Static idleTimeCount As Single 'Idle time counter
  Static idleTimeSecs As Single  'Idle time in seconds
  Static passedOnce As Boolean  'Used for first time timer started
  Dim i As Integer        'Just a counter
  'Check for first pass to set timer
  If passedOnce = False Then
    'Get what time the timer started
    idleStartTime = Timer
    passedOnce = True
  End If
  'Set that idle is true, check for mouse and keys movements, etc
  'If there are any then isIdle will become false
  isIdle = True
  'Check API for keypress
  For i = 1 To 256
    'If pressed state becomes -32767
    If GetAsyncKeyState(i) = -32767 Then
      isIdle = False
    End If
  Next i
  'Get CURRENT position of the mouse cursor
  GetCursorPos newMousePos
  'Compare mouse position with last time (has the mouse moved?)
  If newMousePos.X <> oldMousePos.X Or newMousePos.Y <> oldMousePos.Y Then
    'Mouse moved, not idle
    isIdle = False   'Not idle
    'Replace old coordinates with new ones to check next time
    oldMousePos.X = newMousePos.X
    oldMousePos.Y = newMousePos.Y
  End If
  '1. Check if computer WAS idle and user has come back
  If wasIdle And Not isIdle Then
    'Run procedure for when computer comes out of idle state
    IdleFinished
    'Reset wasIdle, so procedure does not run again till next idle time
    wasIdle = False
    'Clear timers
    idleTimeSecs = 0
    idleTimeCount = 0
    idleStartTime = Timer
  End If
  'Check for how long has been idle (seconds - i.e. convert to longs)
  If CLng(idleTimeSecs) > CLng(idleTimeCount) Then
    Debug.Print CLng(idleTimeSecs) & " second(s) have passed on idle"
    idleTimeCount = idleTimeSecs
  End If
  'Computer was not idle but has become idle after x seconds
  If Not wasIdle And isIdle And idleTimeSecs >= IDLESECONDS Then
    'Computer becomes idle, set wasIdle to true so can run
    'procedure after computer comes out of idle state
    wasIdle = True
    'Run procedure for "Idle"
    IdleStarted idleTimeSecs
  End If
  'If idle then update time that has been idle, else reset timers
  If isIdle Then
    idleTimeSecs = Timer - idleStartTime
  Else
    Debug.Print "User pressed a key or moved the mouse"
    idleTimeCount = 0
    idleStartTime = Timer
    idleTimeSecs = 0
  End If
End Sub
Private Sub IdleStarted(Optional ByVal numSeconds As Long)
  'Code when idling starts, i.e. user has gone away for x secs
  Debug.Print "Computer was declared idle at " & Now & " after " & numSeconds & " seconds"
  'Put your code here
End Sub
Private Sub IdleFinished()
  'Code when idling stops, i.e. user returns
  Debug.Print "Computer stopped being IDLE at " & Now
  'Put your code here
End Sub
```

