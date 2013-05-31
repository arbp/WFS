#tag Class
Protected Class Console
	#tag Method, Flags = &h21
		Private Sub AttemptReadInput(handle as Integer)
		  #if TargetWin32
		    Declare Function ReadConsoleInput Lib "Kernel32" Alias "ReadConsoleInputA" ( handle as Integer, record as Ptr, length as Integer, ByRef eventsRead as Integer ) as Boolean
		    Declare Sub GetNumberOfConsoleInputEvents Lib "Kernel32" ( handle as Integer, ByRef numEvents as Integer )
		    
		    // First, figure out how many events there are to read for
		    // this handle
		    Dim numberOfEventsToRead as Integer
		    GetNumberOfConsoleInputEvents( handle, numberOfEventsToRead )
		    
		    // If we have nothing to do, then bail out
		    if numberOfEventsToRead <= 0 then return
		    
		    // Now that we know how many events to read, let's allocate a buffer
		    // that's big enough to hold all the events
		    Const kEventRecordSize = 20
		    Const kEventDataSize = 16
		    Dim records as new MemoryBlock( kEventRecordSize * numberOfEventsToRead )
		    
		    // Now read in the events
		    Dim eventsRead as Integer
		    call ReadConsoleInput( handle, records, numberOfEventsToRead, eventsRead )
		    
		    // Now process each of the events that
		    // we've been given, based on the type of the event
		    Dim bs as new BinaryStream( records )
		    for i as Integer = 0 to eventsRead - 1
		      // Read in a short for the event type
		      Dim eventType as Integer = bs.ReadShort
		      Const KEY_EVENT = &h1
		      Const MOUSE_EVENT = &h2
		      Const WINDOW_BUFFER_SIZE_EVENT = &h4
		      
		      // Read another short, which was stupid padding
		      call bs.ReadShort
		      
		      // Now that we know the event type, we can read
		      // in the record and pick which handler we wish to
		      // call to handle the data.
		      select case eventType
		      case KEY_EVENT
		        HandleKeyEvent( bs.Read( kEventDataSize ) )
		      case MOUSE_EVENT
		        HandleMouseEvent( bs.Read( kEventDataSize ) )
		      case WINDOW_BUFFER_SIZE_EVENT
		        HandleBufferEvent( bs.Read( kEventDataSize ) )
		      else
		        System.DebugLog( "Unknown event of type: 0x" + Hex( eventType ) )
		        
		        // We've gotten an event record we can't account
		        // for, so just read the data and throw it on the floor
		        call bs.Read( kEventDataSize )
		      end select
		    next i
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  #if TargetWin32
		    // Allocate the console that we need -- it's alright if this fails since
		    // it means that there's already a console allocated for us
		    Declare Sub AllocConsole Lib "Kernel32" ()
		    
		    AllocConsole
		    
		    // Get the significant number of mouse buttons
		    Declare Sub GetNumberOfConsoleMouseButtons Lib "Kernel32" ( ByRef buttons as Integer )
		    GetNumberOfConsoleMouseButtons( mSignificantMouseButtons )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  #if TargetWin32
		    Declare Sub FreeConsole Lib "Kernel32" ()
		    
		    FreeConsole
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetStdHandle(which as Integer) As Integer
		  #if TargetWin32
		    Declare Function MyGetStdHandle Lib "Kernel32" Alias "GetStdHandle" ( which as Integer ) as Integer
		    
		    return MyGetStdHandle( which )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleBufferEvent(data as MemoryBlock)
		  // The first two bytes are the rows, and the next two bytes are the cols.
		  // There's no other data in the structure
		  BufferChanged( data.Short( 0 ), data.Short( 2 ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleKeyEvent(data as MemoryBlock)
		  // The structure looks like this:
		  //
		  // typedef struct _KEY_EVENT_RECORD {
		  //   BOOL bKeyDown;
		  //   WORD wRepeatCount;
		  //   WORD wVirtualKeyCode;
		  //   WORD wVirtualScanCode;
		  //   union {
		  //      WCHAR UnicodeChar;
		  //      CHAR AsciiChar;
		  //   } uChar;
		  //   DWORD dwControlKeyState;
		  // } KEY_EVENT_RECORD;
		  
		  // First, figure out whether we're calling a KeyDown
		  // event, or a KeyUp event.
		  
		  if data.Long( 0 ) <> 0 then
		    // It's a key down event
		    KeyDown( data.Short( 6 ), data.Long( 12 ), data.Short( 4 ) )
		  else
		    // It's a key up event
		    KeyUp( data.Short( 6 ), data.Long( 12 ) )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleMouseEvent(data as MemoryBlock)
		  // The structure looks like this:
		  //
		  // typedef struct _MOUSE_EVENT_RECORD {
		  //   COORD dwMousePosition;
		  //   DWORD dwButtonState;
		  //   DWORD dwControlKeyState;
		  //   DWORD dwEventFlags;
		  // } MOUSE_EVENT_RECORD;
		  
		  // So first, get the X and Y coordinates
		  dim x, y as Integer
		  
		  x = data.Short( 0 )
		  y = data.Short( 2 )
		  
		  // Get the button mask
		  dim buttonMask as Integer = data.Long( 4 )
		  
		  // Then get the modifier key mask
		  dim keyMask as Integer = data.Long( 8 )
		  
		  // Finally, figure out which type of event it is
		  dim eventType as Integer = data.Long( 12 )
		  
		  Const MOUSE_BUTTON = &h0
		  Const DOUBLE_CLICK = &h2
		  Const MOUSE_MOVED = &h1
		  Const MOUSE_WHEELED = &h4
		  
		  select case eventType
		  case MOUSE_BUTTON
		    MouseButtonStateChanged( buttonMask, x, y, keyMask )
		  case DOUBLE_CLICK
		    MouseDoubleClick( x, y, keyMask )
		  case MOUSE_MOVED
		    MouseMove( x, y, keyMask )
		  case MOUSE_WHEELED
		    MouseWheel( x, y, keyMask )
		  else
		    System.DebugLog( "Unknown mouse event of type: 0x" + Hex( eventType ) )
		  end select
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function KeyStringFromVirtualKeyCode(virtualKeyCode as Integer) As String
		  // And be sure to add on the virtual key string
		  #if TargetWin32
		    Soft Declare Function MapVirtualKeyA Lib "User32" ( vk as Integer, type as Integer ) as Integer
		    Soft Declare Function MapVirtualKeyW Lib "User32" ( vk as Integer, type as Integer ) as Integer
		    
		    Soft Declare Function GetKeyNameTextA Lib "User32" ( lParam as Integer, name as Ptr, size as Integer ) as Integer
		    Soft Declare Function GetKeyNameTextW Lib "User32" ( lParam as Integer, name as Ptr, size as Integer ) as Integer
		    
		    // First, map the virtual key into a scan code
		    dim scanCode as Integer
		    if System.IsFunctionAvailable( "MapVirtualKeyW", "User32" ) then
		      scanCode = MapVirtualKeyA( virtualKeyCode, 0 )
		    else
		      scanCode = MapVirtualKeyW( virtualKeyCode, 0 )
		    end if
		    
		    // The GetKeyNameText function expects the
		    // scan code to occupy bits 16 thru 23, so we
		    // must shift ourselves over to occupy that space
		    scanCode = BItwise.ShiftLeft( scanCode, 16 )
		    
		    // Then get the key text
		    dim keyText as new MemoryBlock( 32 )
		    dim keyTextLen as Integer
		    
		    if System.IsFunctionAvailable( "GetKeyNameTextW", "User32" ) then
		      keyTextLen = GetKeyNameTextW( scanCode, keyText, keyText.SIze )
		      return keyText.WString( 0 )
		    else
		      keyTextLen = GetKeyNameTextA( scanCode, keyText, keyText.SIze )
		      return keyText.CString( 0 )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ListenForInputEvents()
		  if mConsoleHandles( kStdInIndex ) = 0 then OpenInputStream
		  
		  // Setup the input messages
		  SetupInput( mConsoleHandles( kStdInIndex ) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ModifierKeyPressed(whichModifier as Integer, modifierMask as Integer) As Boolean
		  // We just want to see whether a modifier key is pressed or
		  // not.  So check whether the modifier is in the mask.
		  
		  return Bitwise.BitAnd( modifierMask, whichModifier ) <> 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MouseButtonDown(positionFromLeft as Integer, buttonMask as Integer) As Boolean
		  if positionFromLeft = 0 then
		    return Bitwise.BitAnd( buttonMask, &h1 ) = &h1
		  elseif positionFromLeft = kRightmostButton then
		    return Bitwise.BitAnd( buttonMask, &h2 ) = &h2
		  else
		    dim maskFlag as Integer = 2 ^ (positionFromLeft + 1)
		    return Bitwise.BitAnd( buttonMask, maskFlag ) = maskFlag
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NumberOfSignificantButtons() As Integer
		  return mSignificantMouseButtons
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub OpenErrorStream()
		  Const STD_ERROR_HANDLE = -12
		  
		  Dim handle as Integer = GetStdHandle( STD_ERROR_HANDLE )
		  
		  // Stick this into our array
		  mConsoleHandles( kStdErrIndex ) = handle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub OpenInputStream()
		  Const STD_INPUT_HANDLE = -10
		  
		  Dim handle as Integer = GetStdHandle( STD_INPUT_HANDLE )
		  
		  // Stick this into our array
		  mConsoleHandles( kStdInIndex ) = handle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub OpenOutputStream()
		  Const STD_OUTPUT_HANDLE = -11
		  
		  Dim handle as Integer = GetStdHandle( STD_OUTPUT_HANDLE )
		  
		  // Stick this into our array
		  mConsoleHandles( kStdOutIndex ) = handle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Poll()
		  // We want to poll for the various console events that
		  // can be read in.
		  for each handle as Integer in mConsoleHandles
		    if handle <> 0 then
		      // Try to read and process some input from
		      // the handle
		      AttemptReadInput( handle )
		    end if
		  next handle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Read(numCharacters as Integer) As String
		  #if TargetWin32
		    Declare Sub ReadConsoleA Lib "Kernel32" ( handle as Integer, buf as Ptr, size as Integer, ByRef read as Integer, reserved as Integer )
		    
		    if mConsoleHandles( kStdInIndex ) = 0 then return ""
		    
		    Dim mb as new MemoryBlock( numCharacters )
		    Dim read as Integer
		    
		    ReadConsoleA( mConsoleHandles( kStdInIndex ), mb, numCharacters, read, 0 )
		    
		    return mb.StringValue( 0, read )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SetupInput(handle as Integer)
		  #if TargetWin32
		    Declare Function SetConsoleMode Lib "Kernel32" ( handle as Integer, mode as Integer ) as Boolean
		    Declare Sub GetConsoleMode Lib "Kernel32" ( handle as Integer, ByRef mode as Integer )
		    
		    Dim oldMode as Integer
		    GetConsoleMode( handle, oldMode )
		    
		    Const ENABLE_WINDOW_INPUT = &h8
		    Const ENABLE_MOUSE_INPUT = &h10
		    oldMode = BitwiseOr( oldMode, ENABLE_MOUSE_INPUT + ENABLE_WINDOW_INPUT )
		    
		    dim worked as Boolean = SetConsoleMode( handle, oldMode )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Write(data as String, toErrStream as Boolean = false)
		  #if TargetWin32
		    Dim handle as Integer
		    
		    if toErrStream then handle = mConsoleHandles( kStdErrIndex ) else handle = mConsoleHandles( kStdOutIndex )
		    if handle = 0 then return
		    
		    Declare Sub WriteConsoleA Lib "Kernel32" ( handle as Integer, buffer as Ptr, size as Integer, ByRef written as Integer, reserved as Integer )
		    
		    Dim mb as MemoryBlock = data
		    Dim written as Integer
		    WriteConsoleA( handle, mb, mb.Size, written, 0 )
		  #endif
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event BufferChanged(rows as Integer, cols as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event KeyDown(virtualKeyCode as Integer, modifierKeyStateMask as Integer, repeatRate as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event KeyUp(virtualKeyCode as Integer, modifierKeyStateMask as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseButtonStateChanged(buttonStateMask as Integer, x as Integer, y as Integer, modifierKeyStateMask as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseDoubleClick(x as Integer, y as Integer, modifierKeyStateMask as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseMove(x as Integer, y as Integer, modifierKeyStateMask as Integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseWheel(x as Integer, y as Integer, modifierKeyStateMask as Integer)
	#tag EndHook


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  #if TargetWin32
			    Declare Sub GetConsoleMode Lib "Kernel32" ( handle as Integer, ByRef mode as Integer )
			    
			    if mConsoleHandles( kStdInIndex ) = 0 then return false
			    
			    Dim mode as Integer
			    GetConsoleMode( mConsoleHandles( kStdInIndex ), mode )
			    
			    Const ENABLE_QUICK_EDIT_MODE = &h40
			    
			    return Bitwise.BitAnd( mode, ENABLE_QUICK_EDIT_MODE ) = ENABLE_QUICK_EDIT_MODE
			  #endif
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  #if TargetWin32
			    Declare Sub SetConsoleMode Lib "Kernel32" ( handle as Integer, mode as Integer )
			    Declare Sub GetConsoleMode Lib "Kernel32" ( handle as Integer, ByRef mode as Integer )
			    
			    if mConsoleHandles( kStdInIndex ) = 0 then return
			    
			    Dim mode as Integer
			    GetConsoleMode( mConsoleHandles( kStdInIndex ), mode )
			    
			    Const ENABLE_EXTENDED_FLAGS = &h80
			    Const ENABLE_QUICK_EDIT_MODE = &h40
			    
			    if value then
			      mode = Bitwise.BitOr( mode, ENABLE_EXTENDED_FLAGS + ENABLE_QUICK_EDIT_MODE )
			    else
			      mode = Bitwise.BitAnd( mode, Bitwise.OnesComplement( ENABLE_EXTENDED_FLAGS + ENABLE_QUICK_EDIT_MODE ) )
			    end if
			    
			    SetConsoleMode( mConsoleHandles( kStdInIndex ), mode )
			  #endif
			End Set
		#tag EndSetter
		AllowQuickEdit As Boolean
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  #if TargetWin32
			    Declare Sub GetConsoleMode Lib "Kernel32" ( handle as Integer, ByRef mode as Integer )
			    
			    if mConsoleHandles( kStdInIndex ) = 0 then return false
			    
			    Dim mode as Integer
			    GetConsoleMode( mConsoleHandles( kStdInIndex ), mode )
			    
			    Const ENABLE_LINE_INPUT = &h2
			    
			    return Bitwise.BitAnd( mode, ENABLE_LINE_INPUT ) = ENABLE_LINE_INPUT
			  #endif
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  #if TargetWin32
			    Declare Sub SetConsoleMode Lib "Kernel32" ( handle as Integer, mode as Integer )
			    Declare Sub GetConsoleMode Lib "Kernel32" ( handle as Integer, ByRef mode as Integer )
			    
			    if mConsoleHandles( kStdInIndex ) = 0 then return
			    
			    Dim mode as Integer
			    GetConsoleMode( mConsoleHandles( kStdInIndex ), mode )
			    
			    Const ENABLE_ECHO_INPUT = &h4
			    Const ENABLE_LINE_INPUT = &h2
			    
			    if value then
			      mode = Bitwise.BitOr( mode, ENABLE_ECHO_INPUT + ENABLE_LINE_INPUT )
			    else
			      mode = Bitwise.BitAnd( mode, Bitwise.OnesComplement( ENABLE_ECHO_INPUT + ENABLE_LINE_INPUT ) )
			    end if
			    
			    SetConsoleMode( mConsoleHandles( kStdInIndex ), mode )
			  #endif
			End Set
		#tag EndSetter
		AsyncRead As Boolean
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  #if TargetWin32
			    Declare Function GetConsoleWindow Lib "Kernel32" () as Integer
			    
			    return GetConsoleWindow
			  #endif
			End Get
		#tag EndGetter
		Handle As Integer
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		#tag Note
			Index          Purpose
			---------         ---------------
			0                 Standard Input
			1                 Standard Output
			2                 Standard Error
			
			If the handle is 0, then we don't care to use it
		#tag EndNote
		Private mConsoleHandles(3) As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSignificantMouseButtons As Integer
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  #if TargetWin32
			    Declare Function GetConsoleTitleA Lib "Kernel32" ( buf as Ptr, size as Integer ) as Integer
			    
			    Dim mb as new MemoryBlock( 64 * 1024 )
			    Dim numChars as Integer = GetConsoleTitleA( mb, mb.Size )
			    
			    return mb.CString( 0 )
			  #endif
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  #if TargetWin32
			    Declare Sub SetConsoleTitleA Lib "Kernel32" ( yeah as CString )
			    
			    SetConsoleTitleA( value )
			  #endif
			End Set
		#tag EndSetter
		Title As String
	#tag EndComputedProperty


	#tag Constant, Name = kAlt, Type = Double, Dynamic = False, Default = \"&h3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCapsLock, Type = Double, Dynamic = False, Default = \"&h80", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kCtrl, Type = Double, Dynamic = False, Default = \"&hC", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kEnhancedKey, Type = Double, Dynamic = False, Default = \"&h100", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kLeftAlt, Type = Double, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kLeftCtrl, Type = Double, Dynamic = False, Default = \"&h8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kRightAlt, Type = Double, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kRightCtrl, Type = Double, Dynamic = False, Default = \"&h4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kRightmostButton, Type = Double, Dynamic = False, Default = \"-1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kScrollLock, Type = Double, Dynamic = False, Default = \"&h40", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kShiftKey, Type = Double, Dynamic = False, Default = \"&h10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStdErrIndex, Type = Double, Dynamic = False, Default = \"2", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kStdInIndex, Type = Double, Dynamic = False, Default = \"0", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kStdOutIndex, Type = Double, Dynamic = False, Default = \"1", Scope = Private
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="AllowQuickEdit"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AsyncRead"
			Group="Behavior"
			InitialValue="0"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Handle"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Title"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
