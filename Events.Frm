VERSION 5.00
Begin VB.Form Events 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   480
   ClientLeft      =   2400
   ClientTop       =   1710
   ClientWidth     =   2160
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   480
   ScaleWidth      =   2160
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1575
      Top             =   75
   End
End
Attribute VB_Name = "Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------
' Events.frm - by Francesco Balena
'
' add pseudo-events to your VB app
'-----------------------------------

Option Explicit

' Public properties

' used when Register is called without its "event" argument
Public DefaultEvents As Long
' delay time for showing tooltips (milliseconds)
Public TooltipDelay As Long

#If Win16 Then

' API structures
Private Type POINTAPI
    X As Integer
    Y As Integer
End Type
Private Type RECT
    Left    As Integer
    Top     As Integer
    Right   As Integer
    Bottom  As Integer
End Type

' API functions
Private Declare Function GetActiveWindow Lib "User" () As Integer
Private Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Private Declare Function GetCursorPos Lib "User" (lpPoint As POINTAPI) As Integer
Private Declare Function GetAsyncKeyState Lib "User" (ByVal vKey As Integer) As Integer
Private Declare Function ScreenToClient Lib "User" (ByVal hWnd As Integer, lpPoint As POINTAPI) As Integer
Private Declare Function WindowFromPoint Lib "User" (ByVal X As Integer, ByVal Y As Integer) As Integer
Private Declare Function GetWindowRectS Lib "User" Alias "GetWindowRect" (ByVal hWnd As Integer, ByVal values As String) As Long
Private Declare Function SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

' API constants
Const HWND_TOPMOST = -1
Const SWP_NOACTIVATE = &H10
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
Const GWW_HINSTANCE = (-6)

#Else

'API structures
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

' API functions
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRectS Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, ByVal values As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'API constants
Const HWND_TOPMOST = -1
Const SWP_NOACTIVATE = &H10
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
Const GWL_HINSTANCE = (-6)

#End If

'---------------------------------
' private form level variables
'---------------------------------

' the main structure is the regForms() array, that always holds
' info on all registered forms. The Register and Unregister methods
' add and delete items in this array. The RegisteredFormIndex private
' function scan this array to understand if a form is registered and
' retrieve additional info, such as registered events and position

' one record for each registered form
Private Type TRegForms
    frm         As Form     ' registered form
    regEvents   As Long     ' registered events (bit-coded)
    activeCtrl  As Control  ' current active control
    rectS       As String   ' current position and size (RECT coded in a string)
End Type
' regForms() keeps track of registered forms
Dim regForms() As TRegForms
Dim regFormsCount As Integer

' all events that are registered, by any form (bit-coded field)
' this is obtained OR-ing the regEvents fields for all the forms
' that are currently registered
' (is updated in Register and Unregister methods)
Dim registeredEvents As Long

' information on the active form
Dim currForm As Form
Dim saveForm As Form
'index into colRegisteredForms collection
Dim currFormIndex As Variant
' registered events for current form (bit-coded field)
Dim currFormEvents As Long
Dim saveFormEvents As Long
' window position and size as a 8- or 16-char string
' this is like a RECT structure stored in a string - this
' methods enables very quick comparisons when it is time
' to decide if the window has been moved or resized
Dim currFormRect As String

' the handle of the window that has the focus
' (might belong to another application)
Dim currWindow As Long
Dim saveWindow As Long

' information about instances
'
' the instance handle of this application
Dim thisInstance As Long
' the instance handle of the application that has currently the focus
Dim currInstance As Long
Dim saveInstance As Long
' True if this application has the focus
Dim hasFocus As Boolean
Dim saveHasFocus As Boolean

' information about the mouse
'
' mouse buttons and coordinates
Dim currButton As Integer
Dim saveButton As Integer
Dim currMousePoint As POINTAPI
Dim saveMousePoint As POINTAPI
' which kind of mouse activity as occurred
Dim mouseAction As Integer  '0=none, 1=down 2=up, 3=move
' shift key status
Dim currShift As Integer
' mouse position in local coordinate system
Dim localMouseX As Single
Dim localMouseY As Single
' the control under the mouse (on current form)
Dim currPointedCtrl As Control
Dim savePointedCtrl As Control

' this counts down from TooltipDelay to zero
Dim tooltipTimeout As Long

' these variables are used to track if internal subs
' have been executed in the current timer tick
Dim GetMouseDone As Boolean
Dim GetMouseXYDone As Boolean
Dim GetMouseControlDone As Boolean
Dim GetShiftDone As Boolean

' constants for Event activation - these are exposed as evXXXX readonly
' properties for use in the Register method
Const EV_APPACTIVATE = &H1      ' AppActivate & AppDeactivate events
Const EV_MOVE = &H4             ' Move event
Const EV_MOUSE = &H100          ' normal mouse events (mouse within the client area)
Const EV_MOUSENC = &H200        ' mouse events event outside client area
Const EV_MOUSECAPTURE = &H400   ' mouse events even if form is not active (but application is)
Const EV_MOUSEALWAYS = &H800    ' mouse events always
Const EV_MOUSEENTER = &H1000    ' MouseEnter & MouseExit events
Const EV_CHANGEFOCUS = &H2000   ' GotFocus & LostFocus events
Const EV_TOOLTIPS = &H4000      ' ShowTooltip & HideTooltip events
Const EV_TOOLTIPSAUTO = &H8000  ' show tooltip using the Tag property
' compounded constants (do not set from the outside)
Const EV_MOUSEALL = EV_MOUSE + EV_MOUSENC + EV_MOUSECAPTURE + EV_MOUSEALWAYS
Const EV_TOOLTIPSALL = EV_TOOLTIPS + EV_TOOLTIPSAUTO
Public Property Get evAppActivate() As Long
    evAppActivate = EV_APPACTIVATE
End Property
Public Property Get evMove() As Long
    evMove = EV_MOVE
End Property
Public Property Get evMouse() As Long
    evMouse = EV_MOUSE
End Property
Public Property Get evMouseNC() As Long
    evMouseNC = EV_MOUSENC
End Property
Public Property Get evMouseCapture() As Long
    evMouseCapture = EV_MOUSECAPTURE
End Property
Public Property Get evMouseAlways() As Long
    evMouseAlways = EV_MOUSEALWAYS
End Property
Public Property Get evMouseEnter() As Long
    evMouseEnter = EV_MOUSEENTER
End Property
Public Property Get evChangeFocus() As Long
    evChangeFocus = EV_CHANGEFOCUS
End Property
Public Property Get evTooltips() As Long
    evTooltips = EV_TOOLTIPS
End Property
Public Property Get evTooltipsAuto() As Long
    evTooltipsAuto = EV_TOOLTIPSAUTO
End Property

Public Property Get MouseWindowHWnd() As Long

    ' returns the handle of window pointed to by the mouse
    ' or zero if now window exists
    '
    ' IMPORTANT: only works if at least one MouseXXXX
    '            pseudo-event has been registered
    MouseWindowHWnd = WindowFromPoint(currMousePoint.X, currMousePoint.Y)

End Property

Public Sub ShowTooltip(Optional text As Variant)

    '---------------------------------------------------------
    ' show a tooltip at current mouse coordinates
    ' if argument is a null string or imitted, the tooltip is hidden
    '---------------------------------------------------------
    
    On Error Resume Next
    
    ' in all cases this form should be hidden
    Me.Visible = False
    Me.Cls
    
    If Len(text) = 0 Then
        ' this line is executed both if argument is missing
        ' or it is a null string - the form stays hidden
    Else
        Dim frm As Form
        Dim X As Single, Y As Single
        Dim fw As Single, fh As Single
        ' evaluate tooltip size
        fw = TextWidth(text) + 50
        fh = TextHeight(text) + 50
        ' get mouse position in local coordinates
        GetMouseXY Me, X, Y
        ' convert to screen coordinates, adjust for standard offset
        X = X + Me.Left - (fw / 2)
        Y = Y + Me.Top + 450
        ' prevent out of screen
        If X < 0 Then
            X = 0
        ElseIf X + fw > Screen.Width Then
            X = Screen.Width - fw
        End If
        If Y < 0 Then
            Y = 0
        ElseIf Y + fh > Screen.Height Then
            Y = Screen.Height - fh
        End If
        ' move and size the form
        Me.Move X, Y, fw, fh
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE + SWP_SHOWWINDOW
        Me.Print text;
        Me.Visible = True
    End If

End Sub

Private Sub Form_Initialize()

    ' initialize structures
    ReDim regForms(50) As TRegForms
    ' create a string 8- or 16-char long
    currFormRect = Space$(LenB("12345678"))

    ' initialize properties
    DefaultEvents = -1   ' all events
    TooltipDelay = 1000  ' one second

    ' retrieve the handle of this instance
    ' it will be used to understand if this application
    ' has lost the input focus
#If Win16 Then
    thisInstance = GetWindowWord(Me.hWnd, GWW_HINSTANCE)
#Else
    thisInstance = GetWindowLong(Me.hWnd, GWL_HINSTANCE)
#End If

    ' load the form (and activate the timer)
    ' the first time it is referenced
    Load Me
    
End Sub


Public Property Get Interval() As Long

    ' return internal timer's Interval property
    Interval = Timer1.Interval

End Property

Public Property Let Interval(value As Long)

    ' set internal timer's Interval property
    Timer1.Interval = value

End Property

Public Sub Register(frm As Form, Optional regEvents)

    '---------------------------------------------------------
    ' register a new form for receiving events
    ' if 2nd argument is missing or not numeric, the form will receive all events
    '---------------------------------------------------------
    
    On Error Resume Next
    
    Dim Index As Integer
    Dim formIndex As Integer
    
    ' account for optional argument
    If IsMissing(regEvents) Then
        regEvents = DefaultEvents
    ElseIf Not IsNumeric(regEvents) Then
        regEvents = DefaultEvents
    End If
    
    ' if the 2nd argument is zero, unregister the form
    If regEvents = 0 Then
        Unregister frm
        Exit Sub
    End If
    
    ' search this form in the regForms() array
    For Index = 1 To regFormsCount
        If regForms(Index).frm Is frm Then
            ' this form has been already registered
            formIndex = Index
            Exit For
        ElseIf formIndex = 0 Then
            ' if this item is empty, remember its position for later
            If regForms(Index).frm Is Nothing Then formIndex = Index
        End If
    Next
    
    If formIndex = 0 Then
        ' create some empty slots in the array if necessary
        If regFormsCount = UBound(regForms) Then
            ReDim Preserve regForms(regFormsCount + 10) As TRegForms
        End If
        regFormsCount = regFormsCount + 1
        formIndex = regFormsCount
    End If
    
    Set regForms(formIndex).frm = frm
    Set regForms(formIndex).activeCtrl = Nothing
    
    If regForms(formIndex).regEvents = 0 Then
        ' if this is an empty slot we can quickly evaluate the new
        ' value of registeredEvents
        regForms(formIndex).regEvents = regEvents
        registeredEvents = registeredEvents Or regEvents
    Else
        ' else we have rebuild the bit-coded field that store which
        ' events have been registered by at least one form
        regForms(formIndex).regEvents = regEvents
        registeredEvents = 0
        For Index = 1 To regFormsCount
            registeredEvents = registeredEvents Or regEvents
        Next
        ' if this is the current index, update currFormEvents
        If formIndex = currFormIndex Then
            currFormEvents = regEvents
        End If
    End If
    
End Sub

Public Sub Unregister(frm As Form)

    '---------------------------------------------------------
    ' unregister a form
    '---------------------------------------------------------
    
    Dim Index As Integer
    
    ' fire the last events, if any are pending
    FireEvents
    
    ' scan the regForms() array to look for the form
    ' at the same time, rebuild the registeredEvents variable
    registeredEvents = 0
    For Index = 1 To regFormsCount
        If frm Is regForms(Index).frm Then
            LSet regForms(Index) = regForms(0)
        Else
            ' rebuild registeredEvents for all remaining events
            registeredEvents = registeredEvents Or regForms(Index).regEvents
        End If
    Next
    
    ' now, try to decrease regFormsCount
    ' this will speed all subsequent lookups
    Do While regForms(regFormsCount).frm Is Nothing And regFormsCount > 0
        regFormsCount = regFormsCount - 1
    Loop
    ' if this was the last registered form, unload the module
    If regFormsCount = 0 Then
        Unload Me
    End If
    
End Sub

Public Property Get MouseX() As Integer

    ' return current mouse position in pixel / screen coordinates
    '
    ' IMPORTANT: only works if at least one MouseXXXX
    '            pseudo-event has been registered
    MouseX = currMousePoint.X

End Property

Public Property Get MouseY() As Integer

    ' return current mouse position in pixel / screen coordinates
    '
    ' IMPORTANT: only works if at least one MouseXXXX
    '            pseudo-event has been registered
    MouseY = currMousePoint.Y

End Property

Public Sub FireEvents()

    '---------------------------------------------------------
    ' fire all pending pseudo-events
    '---------------------------------------------------------
    
    On Error Resume Next
    
    ' reset all "done" flags
    GetMouseDone = False
    GetMouseXYDone = False
    GetShiftDone = False
    GetMouseControlDone = False
    
    ' save values before getting updated information
    Set saveForm = currForm
    saveWindow = currWindow
    saveInstance = currInstance
    saveHasFocus = hasFocus
    saveFormEvents = currFormEvents
    
    ' gather new values
    Set currForm = Screen.ActiveForm
    
    If Not (saveForm Is currForm) Then
        ' search the form index only if current form is different
        currFormIndex = RegisteredFormIndex(currForm)
        currFormEvents = regForms(currFormIndex).regEvents
    End If
    
    ' which application has the input focus?
#If Win16 Then
    currWindow = GetActiveWindow()
    If currWindow <> saveWindow Then
        currInstance = GetWindowWord(currWindow, GWW_HINSTANCE)
        hasFocus = (currInstance = thisInstance)
    End If
#Else
    currWindow = GetForegroundWindow()
    If currWindow <> saveWindow Then
        currInstance = GetWindowLong(currWindow, GWL_HINSTANCE)
        hasFocus = (currInstance = thisInstance)
    End If
#End If

    ' call procedures only for those events that have
    ' been registered at least by one form

    If registeredEvents And EV_APPACTIVATE Then
        CheckAppActivate
    End If
    
    If registeredEvents And EV_MOUSEALL Then
        CheckMouse
    End If
    
    If registeredEvents And EV_MOUSEENTER Then
        CheckMouseEnter
    End If
    
    If Not hasFocus Then Exit Sub
    
    ' all subsequent events are never called
    ' if this application does not have the focus
    
    If registeredEvents And EV_TOOLTIPSALL Then
        CheckTooltips
    End If
    
    If currFormEvents And EV_MOVE Then
        CheckMove
    End If
    
    If currFormEvents And EV_CHANGEFOCUS Then
        CheckChangeFocus
    End If
    
End Sub
Private Sub Timer1_Timer()

    FireEvents
    
End Sub

Private Function RegisteredFormIndex(frm As Form) As Integer

    '---------------------------------------------------------
    ' returns then the index of a form into regForms()
    '---------------------------------------------------------
    
    Dim Index As Integer
    
    RegisteredFormIndex = 0
    
    If Not frm Is Nothing Then
        ' we can only use linear search
        For Index = 1 To regFormsCount
            If frm Is regForms(Index).frm Then
                RegisteredFormIndex = Index
                Exit For
            End If
        Next
    End If
        
End Function

Private Sub GetMouse()

    '---------------------------------------------------------
    ' retrieve general information on mouse
    '---------------------------------------------------------

    ' execute this code only the first time after a timer click
    If GetMouseDone Then Exit Sub
    GetMouseDone = True
    
    ' save previous values
    LSet saveMousePoint = currMousePoint
    saveButton = currButton
    
    ' get cursor position in screen coordinates
    GetCursorPos currMousePoint
    
    ' get mouse button state - we have to use GetAsynchKeyState
    ' because it is the only function that works on
    ' a system wide basis under Win32
    currButton = 0
    If GetAsyncKeyState(vbLeftButton) Then currButton = vbLeftButton
    If GetAsyncKeyState(vbRightButton) Then currButton = currButton Or vbRightButton
    If GetAsyncKeyState(vbMiddleButton) Then currButton = currButton Or vbMiddleButton
    
    ' decide if mouse has been moved or clicked
    If currButton <> saveButton Then
        If (currButton Or saveButton) = currButton Then
            ' currButton has more "ones" than saveButton
            ' therefore a button must have been pressed
            mouseAction = 1
        Else
            ' else a button has been released
            mouseAction = 2
        End If
    ElseIf (currMousePoint.X <> saveMousePoint.X) Or (currMousePoint.Y <> saveMousePoint.Y) Then
        ' mouse has been moved
        mouseAction = 3
    Else
        ' no mouse action
        mouseAction = 0
    End If
        
End Sub


Private Sub GetMouseXY(frm As Form, X As Single, Y As Single)

    '---------------------------------------------------------
    ' retrieve information on mouse related to a form
    '---------------------------------------------------------

    If frm Is currForm Then
        ' since this procedure might be called many a time for the
        ' current form, it bufferizes results
        If GetMouseXYDone Then
            X = localMouseX
            Y = localMouseY
        Else
            GetMouseXYDone = True
            GoSub GetMouseXY_Sub
            localMouseX = X
            localMouseY = Y
        End If
    Else
        GoSub GetMouseXY_Sub
    End If
    
    Exit Sub

GetMouseXY_Sub:
    Dim lpPoint As POINTAPI
    With frm
        ' convert to client coordinates
        LSet lpPoint = currMousePoint
        ScreenToClient .hWnd, lpPoint
        ' convert to current scalemode
        X = .ScaleX(lpPoint.X, vbPixels, .ScaleMode) + .ScaleLeft
        Y = .ScaleY(lpPoint.Y, vbPixels, .ScaleMode) + .ScaleTop
    End With
    Return

End Sub


Private Sub GetShift()

    '---------------------------------------------------------
    ' retrieve the current status of Alt, Shift and Ctrl keys
    '---------------------------------------------------------
    
    If GetShiftDone Then Exit Sub
    GetShiftDone = True
    
    ' get shift key state - we have to use GetAsynchKeyState
    ' because it is the only keyboard function that works on
    ' a system wide basis under Win32
    currShift = 0
    If GetAsyncKeyState(vbKeyMenu) Then currShift = vbAltMask
    If GetAsyncKeyState(vbKeyShift) Then currShift = currShift Or vbShiftMask
    If GetAsyncKeyState(vbKeyControl) Then currShift = currShift Or vbCtrlMask

End Sub

Private Sub GetMouseControl()

    '---------------------------------------------------------
    ' get the control under the mouse on current form
    '---------------------------------------------------------

    ' exit if this routine has been already called in current tick
    If GetMouseControlDone Then Exit Sub
    GetMouseControlDone = True
    
    Dim X As Single, Y As Single
    Dim Index As Integer
    Dim hw As Long
    Dim ctrl As Control
    Dim pointedCtrl As Control
    Dim lpPoint As POINTAPI
    
    On Error Resume Next
    
    Set savePointedCtrl = currPointedCtrl
    
    ' return Nothing if the current form is not registered
    ' or if the current form is not the active form
    If currFormIndex = 0 Or currForm.hWnd <> currWindow Then
        Set currPointedCtrl = Nothing
        Exit Sub
    End If
    
    ' gather mouse info
    GetMouse
    ' if mouse has not been moved and no new form has been
    ' activate, the control under the mouse did not change
    If mouseAction <> 3 Then
        If currForm Is saveForm Then Exit Sub
    End If
    
    ' convert mouse position to local coordinates
    GetMouseXY currForm, X, Y
    
    ' see which control is under the mouse
    hw = MouseWindowHWnd
    If hw <> currForm.hWnd Then
        ' cursor is not directly on the form
        With currForm
            ' check if it is inside the client area of the form
            If X < .ScaleLeft Then
            ElseIf Y < .ScaleTop Then
            ElseIf X >= .ScaleWidth Then
            ElseIf Y >= .ScaleHeight Then
            Else
                ' the mouse is over a control with a hWnd property
                ' scan the Controls collection too see which control it is
                For Each ctrl In .Controls
                    ' the following comparison fails either if
                    ' control does not support hWnd property
                    ' or if the handle is different
                    If ctrl.hWnd <> hw Then
                        ' hWnd different or unsupported property
                    Else
                        Set pointedCtrl = ctrl
                        Exit For
                    End If
                Next
                ' now we need to check if the mouse is pointing
                ' to a graphical control inside "pointedCtrl"
                If Not pointedCtrl Is Nothing Then
                    ' make (x,y) relative to the container
                    Err = 0
                    Set ctrl = pointedCtrl
                    ' next loop will repeat as long as the parent
                    ' of "ctrl" is not the form itself
                    Do
                        X = X - ctrl.Left
                        Y = Y - ctrl.Top
                        Set ctrl = ctrl.Container
                    Loop Until Err
                    
                    For Each ctrl In .Controls
                        With ctrl
                            If ctrl.Container Is pointedCtrl Then
                                ' only check controls contained in "pointedCtrl"
                                ' the following IF structure fails either if any property
                                ' is not supported or it is not in valid range
                                If .Visible = 0 Or X < .Left Or Y < .Top Then
                                    ' *** unsupported or out-of-range properties
                                ElseIf X >= .Left + .Width Then
                                ElseIf Y >= .Top + .Height Then
                                Else
                                    ' all properties are supported and within range
                                    Set pointedCtrl = ctrl
                                    Exit For
                                End If
                            End If
                        End With
                    Next
                End If
            End If
        End With
        
    Else
        ' mouse is over the form, but we have to check all
        ' graphical controls on the form
        For Each ctrl In currForm.Controls
            With ctrl
                ' the next comparison is never satisfied
                ' (because we know that "hw" is the form handle)
                ' but if the control does not support hWnd property the
                ' execution flow goes into the IF block
                ' this way we automatically rule out all controls
                ' that expose a hWnd property
                If ctrl.hWnd = hw Then
                    ' following IF structure fails both if any property
                    ' is not supported or is not in valid range
                    If .Visible = 0 Or X < .Left Or Y < .Top Then
                    ElseIf X >= .Left + .Width Then
                    ElseIf Y >= .Top + .Height Then
                    Else
                        ' all properties are supported and within range
                        Set pointedCtrl = ctrl
                        Exit For
                    End If
                End If
            End With
        Next
    End If
    
    ' at this point pointedCtrl is the control pointed to by the mouse
    ' or nothing if there isn't such a control (mouse over the form)
    
    Set currPointedCtrl = pointedCtrl
    
End Sub


Private Sub CheckAppActivate()

    '-----------------------------------------------
    ' check if the application still hold focus
    '-----------------------------------------------
    
    Dim Index As Integer
    
    On Error Resume Next
    
    If hasFocus Then
        ' this app has the focus
        If Not saveHasFocus Then
            ' if it had not the focus before, it has just regained it
            ' so, fire Event_AppActivate in all forms that registered it
            For Index = 1 To regFormsCount
                If regForms(Index).regEvents And EV_APPACTIVATE Then
                    regForms(Index).frm.Event_AppActivate
                End If
            Next
        End If
    Else
        ' this app does not have the focus
        If saveHasFocus Then
            ' if it had the focus before, it has just lost it
            ' so, fire Event_AppDeactivate in all forms that registered it
            For Index = 1 To regFormsCount
                If regForms(Index).regEvents And EV_APPACTIVATE Then
                    regForms(Index).frm.Event_AppDeactivate
                End If
            Next
        End If
    End If
    
End Sub



Private Sub CheckMove()

    '-----------------------------------------------
    ' check if the active form has been moved
    '-----------------------------------------------
    
    On Error Resume Next
    
    ' get window position and size, into a 16-char string
    GetWindowRectS currForm.hWnd, currFormRect
    ' compare with previous position
    If currFormRect <> regForms(currFormIndex).rectS Then
        ' window has been moved or resized
        regForms(currFormIndex).rectS = currFormRect
        ' fire the Move event
        currForm.Event_Move
    End If

End Sub



Private Sub CheckMouse()

    '-----------------------------------------------
    ' check all mouse pseudo-Events
    '-----------------------------------------------
    
    Dim X As Single, Y As Single
    Dim argButton As Integer
    Dim Index As Integer
    Dim frm As Form
    Dim hw As Long
    ReDim processedForms(regFormsCount) As Boolean
    
    On Error Resume Next
    
    ' get mouse info, exit if no event is to be fired
    GetMouse
    If mouseAction = 0 Then Exit Sub
    
    ' if this is a mouseDown / Up action, the argument passed to the
    ' event is the button that has been clicked or released released
    argButton = saveButton Xor currButton
    
    ' get shift information, it will surely be useful later
    GetShift
    
    ' see if any form activated the EV_MOUSEALWAYS event
    If registeredEvents And EV_MOUSEALWAYS Then
        For Index = 1 To regFormsCount
            If regForms(Index).regEvents And EV_MOUSEALWAYS Then
                ' fire mouse event for corresponding form
                GoSub CheckMouse_FireEvent
            End If
        Next
    End If
    
    ' exit if this application does not have the input focus
    If Not hasFocus Then Exit Sub
    
    ' see if any form activated the EV_MOUSECAPTURE event
    If registeredEvents And EV_MOUSECAPTURE Then
        For Index = 1 To regFormsCount
            If regForms(Index).regEvents And EV_MOUSECAPTURE Then
                ' fire mouse event for corresponding form
                GoSub CheckMouse_FireEvent
            End If
        Next
    End If
                    
    If processedForms(Index) Then
        ' current form has been processed, do nothing
    ElseIf currFormEvents And EV_MOUSENC Then
        ' fire a mouse event for the current form, wherever the mouse is
        Index = currFormIndex
        GoSub CheckMouse_FireEvent
        
    ElseIf currFormEvents And EV_MOUSE Then
        With currForm
            ' exit if current form is not the foreground window
            If currWindow <> .hWnd Then Exit Sub
            ' fire a mouse event for the current form only if mouse
            ' is inside the client area
            GetMouseXY currForm, X, Y
            If X >= .ScaleLeft And Y >= .ScaleTop Then
                If X < .ScaleWidth And Y < .ScaleHeight Then
                    ' if inside the client area, fire the event
                    Index = currFormIndex
                    GoSub CheckMouse_FireEvent
                End If
            End If
        End With
    End If
    
    Exit Sub
    
    
CheckMouse_FireEvent:
    ' check if we already fired a mouse event for this form
    If processedForms(Index) Then Return
    
    ' fire event for form regForms(index).frm
    Set frm = regForms(Index).frm
    GetMouseXY frm, X, Y
    If mouseAction = 1 Then
        frm.Event_MouseDown argButton, currShift, X, Y
    ElseIf mouseAction = 2 Then
        frm.Event_MouseUp argButton, currShift, X, Y
    Else
        frm.Event_MouseMove currButton, currShift, X, Y
    End If
    ' remember we have processed this form
    processedForms(Index) = True
    Return

End Sub



Private Sub CheckMouseEnter()

    '------------------------------------------------
    ' check mouse entering or exiting a control
    '------------------------------------------------

    On Error Resume Next
    
    ' get control under mouse cursor
    GetMouseControl
    
    ' at this point pointedCtrl is the control pointed to by the mouse
    ' or nothing if there isn't such a control (mouse is over the form)
    
    If Not savePointedCtrl Is currPointedCtrl Then
        ' mouse moved on a different control (possibly of a different form)
        If (saveFormEvents And EV_MOUSEENTER) And Not (savePointedCtrl Is Nothing) Then
            ' fire the MouseExit event
            saveForm.Event_MouseExit savePointedCtrl
        End If
        If (currFormEvents And EV_MOUSEENTER) And Not (currPointedCtrl Is Nothing) Then
            ' fire the MouseExit event
            currForm.Event_MouseEnter currPointedCtrl
        End If
    End If
    
End Sub

Private Sub CheckTooltips()

    '------------------------------------------------
    ' check if tooltips are to be showed or hidden
    '------------------------------------------------
    
    On Error Resume Next
    
    ' get control under mouse cursor
    GetMouseControl
    
    ' at this point pointedCtrl is the control pointed to by the mouse
    ' or nothing if there isn't such a control (mouse is over the form)
    
    If Not (savePointedCtrl Is currPointedCtrl) Then
        If tooltipTimeout < 0 And Not (savePointedCtrl Is Nothing) Then
            ' mouse moved on a different control (possibly of a different form)
            ' and a tooltip is presently being showed
            ' therefore, we have to try to fire a TooltipHide event first
            Err = 0
            saveForm.Event_TooltipHide savePointedCtrl
            If Err And Me.Visible Then
                ' if no TooltipHide event was declared in current form
                ' but a tooltip is being showed then hide the tooltip
                ShowTooltip
            End If
        End If
        ' reset the timeout
        tooltipTimeout = TooltipDelay
   
    ElseIf Not (currPointedCtrl Is Nothing) Then
        ' mouse is on the same control as before
        If tooltipTimeout >= 0 Then
            ' no tooltip is presently being showed
            tooltipTimeout = tooltipTimeout - Timer1.Interval
            If tooltipTimeout < 0 Then
                ' timeout has expired, we have to show a tooltip
                ' first, try to fire the TooltipShow event
                If currFormEvents And EV_TOOLTIPS Then
                    Err = 0
                    currForm.Event_TooltipShow currPointedCtrl
                    If Err And (currFormEvents And EV_TOOLTIPSAUTO) Then
                        ' if there is not such an event and auto-mode is
                        ' activated, then show tooltip using internal method
                        ShowTooltip currPointedCtrl.Tag
                    End If
                ElseIf currFormEvents And EV_TOOLTIPSAUTO Then
                    ' only auto-tooltips were enabled
                    ShowTooltip currPointedCtrl.Tag
                End If
            End If
        End If
    End If
End Sub



Private Sub CheckChangeFocus()

    '------------------------------------------------
    ' check if the focus has gone to another control
    '------------------------------------------------

    On Error Resume Next
    
    ' compare current active control with previous one
    If Not (currForm.ActiveControl Is regForms(currFormIndex).activeCtrl) Then
        ' if focus changed then fire the LostFocus event
        ' passing the control that had the focus
        If Not (regForms(currFormIndex).activeCtrl Is Nothing) Then
            currForm.Event_LostFocus regForms(currFormIndex).activeCtrl
        End If
        ' check if the current control is still different from
        ' the previous current control - this way we let the programmer
        ' to set the focus back to the old control, without firing
        ' a useless GotFocus event
        If Not (currForm.ActiveControl Is regForms(currFormIndex).activeCtrl) Then
            If Not (currForm.ActiveControl Is Nothing) Then
                currForm.Event_GotFocus currForm.ActiveControl
            End If
            ' store the current control with focus into the structure
            ' doing this, the programmer is free to use SetFocus within
            ' the event, without causing another GotFocus event
            Set regForms(currFormIndex).activeCtrl = currForm.ActiveControl
        End If
    End If

End Sub


