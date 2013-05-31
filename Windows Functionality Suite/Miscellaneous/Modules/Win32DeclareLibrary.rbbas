#tag Module
Protected Module Win32DeclareLibrary
	#tag Method, Flags = &h1
		Protected Function CompressData(data as MemoryBlock) As MemoryBlock
		  #if TargetWin32
		    Soft Declare Sub RtlCompressBuffer Lib "ntdll" ( format as Integer, data as Ptr, length as Integer, _
		    destBuffer as Ptr, destLength as Integer, unknown as Integer, ByRef destSize as Integer, _
		    workspaceBuffer as Ptr )
		    
		    Soft Declare Sub RtlGetCompressionWorkSpaceSize Lib "ntdll" ( format as Integer, _
		    ByRef bufferSize as Integer, ByRef unknown as Integer )
		    
		    Const COMPRESSION_FORMAT_LZNT1 = &h2
		    dim neededSize, pageSize as Integer
		    pageSize = &h4000
		    RtlGetCompressionWorkSpaceSize( COMPRESSION_FORMAT_LZNT1, neededSize, pageSize )
		    
		    dim workspace as new MemoryBlock( neededSize )
		    dim ret as new MemoryBlock( data.Size )
		    neededSize = data.Size
		    RtlCompressBuffer( COMPRESSION_FORMAT_LZNT1, data, data.Size, ret, 0, _
		    &h4000, neededSize, workspace )
		    
		    ret = new MemoryBlock( neededSize )
		    RtlCompressBuffer( COMPRESSION_FORMAT_LZNT1, data, data.Size, ret, ret.Size, _
		    &h4000, neededSize, workspace )
		    
		    ret.Size = neededSize
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CreateGuid() As MemoryBlock
		  #if TargetWin32
		    Declare Sub CoCreateGuid Lib "Ole32" ( guid as Ptr )
		    
		    // Allocate a structure large enough for the GUID
		    Dim mb as new MemoryBlock( 20 )
		    // And create it
		    CoCreateGuid( mb )
		    
		    return mb
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CreateGuidString() As String
		  #if TargetWin32
		    // Create the GUID and stuff it into a memory block
		    Dim mb as MemoryBlock
		    mb = CreateGuid
		    
		    if mb = nil then return ""
		    
		    // Now make the memory block into a string
		    Declare Function StringFromGUID2 Lib "Ole32" ( guid as Ptr, theStr as Ptr, size as Integer ) as Integer
		    
		    dim numCharacters as Integer
		    dim retStr as new MemoryBlock( 1024 )
		    numCharacters = StringFromGUID2( mb, retStr, retStr.Size )
		    
		    // And return the UTF-16 string.  Remember that we were returned
		    // the number of characters, which is 2 * the number of bytes we want
		    // to pull from the memory block
		    return DefineEncoding( retStr.StringValue( 0, numCharacters * 2 ), Encodings.UTF16 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DecompressData(data as MemoryBlock, bufferSize as Integer = 10485760) As MemoryBlock
		  #if TargetWin32
		    Soft Declare Sub RtlDecompressBuffer Lib "ntdll" ( format as Integer, destBuffer as Ptr, _
		    destLength as Integer, sourceBuffer as Ptr, sourceLength as Integer, ByRef _
		    destSizeNeeded as Integer )
		    
		    Const COMPRESSION_FORMAT_LZNT1 = &h2
		    dim neededSize as Integer
		    dim ret as new MemoryBlock( bufferSize )
		    
		    RtlDecompressBuffer( COMPRESSION_FORMAT_LZNT1, ret, ret.Size, data, data.Size, neededSize )
		    
		    ret.Size = neededSize
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DontCallThisIdleHandler(nCode as Integer, wParam as Integer, lParam as Integer) As Integer
		  if nCode >= 0 then
		    if mIdleHandler <> nil then
		      ' Call the idle handler for the user
		      mIdleHandler.Idle
		    end
		  end if
		  
		  #if TargetWin32
		    Soft Declare Function CallNextHookEx Lib "User32" ( hookHandle as Integer, code as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    
		    ' And make sure we call the next hook in the list
		    return CallNextHookEx( mIdleHandlerHook, nCode, wParam, lParam )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub EmptyClipboard()
		  #if TargetWin32
		    Declare Sub EmptyClipboard Lib "User32" ()
		    Declare Sub OpenClipboard Lib "User32" (wnd as Integer )
		    Declare Sub CloseClipboard Lib "User32" ()
		    
		    OpenClipboard( 0 )
		    EmptyClipboard
		    CloseClipboard
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub EndRestorePoint(cancel as Boolean)
		  if mRestorePointID = 0 then return
		  
		  Const END_SYSTEM_CHANGE = 101
		  Const CANCELLED_OPERATION = 13
		  
		  if cancel then
		    call SRSetRestorePoint( END_SYSTEM_CHANGE, CANCELLED_OPERATION, mRestorePointID )
		  else
		    call SRSetRestorePoint( END_SYSTEM_CHANGE, 0, mRestorePointID )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FillString(char as String, numChars as Integer) As String
		  #if TargetWin32
		    Declare Sub memset lib "msvcrt" ( dest as Ptr, char as Integer, count as Integer )
		    
		    dim mb as new MemoryBlock( LenB( char ) * numChars )
		    memset( mb, AscB( char ), numChars )
		    
		    return DefineEncoding( mb, Encoding( char ) )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FormatErrorMessage() As String
		  ///Gets lasterror and formats system message
		  ///if format fails, return lasterror
		  dim ret as integer
		  dim buffer as memoryBlock
		  Soft Declare Function FormatMessageA Lib "kernel32" (dwFlags As integer, lpSource As integer, dwMessageId As integer, dwLanguageId As integer, lpBuffer As ptr, nSize As integer, Arguments As integer) As integer
		  Soft Declare Function FormatMessageW Lib "kernel32" (dwFlags As integer, lpSource As integer, dwMessageId As integer, dwLanguageId As integer, lpBuffer As ptr, nSize As integer, Arguments As integer) As integer
		  
		  Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
		  
		  if System.IsFunctionAvailable( "FormatMessageW", "Kernel32" ) then
		    buffer = new MemoryBlock( 2048 )
		    ret = FormatMessageW( FORMAT_MESSAGE_FROM_SYSTEM, 0, GetLastError, 0 , Buffer, Buffer.Size, 0 )
		    if ret <> 0 then
		      return Buffer.WString( 0 )
		    end if
		  else
		    buffer = new MemoryBlock( 1024 )
		    ret = FormatMessageA( FORMAT_MESSAGE_FROM_SYSTEM, 0, GetLastError, 0 , Buffer, Buffer.Size, 0 )
		    if ret <> 0 then
		      return Buffer.cstring( 0 )
		    end if
		  end if
		  
		  return str( GetLastError )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub GenerateKeyDown(virtualKeyCode as Integer, extendedKey as Boolean = false)
		  #if TargetWin32
		    Declare Sub keybd_event Lib "User32" ( keyCode as Integer, scanCode as Integer, _
		    flags as Integer, extraData as Integer )
		    
		    dim flags as Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    if extendedKey then
		      flags = KEYEVENTF_EXTENDEDKEY
		    end
		    
		    ' Press the key
		    keybd_event( virtualKeyCode, 0, flags, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub GenerateKeyUp(virtualKeyCode as Integer, extendedKey as Boolean = false)
		  #if TargetWin32
		    Declare Sub keybd_event Lib "User32" ( keyCode as Integer, scanCode as Integer, _
		    flags as Integer, extraData as Integer )
		    
		    dim flags as Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    if extendedKey then
		      flags = KEYEVENTF_EXTENDEDKEY
		    end
		    ' Release the key
		    Const KEYEVENTF_KEYUP = &h2
		    flags = BitwiseOr( flags, KEYEVENTF_KEYUP )
		    keybd_event( virtualKeyCode, 0, flags, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetLastError() As Integer
		  #if TargetWin32
		    Declare Function GetLastError Lib "Kernel32" () as Integer
		    
		    return GetLastError
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub InstallIdleHandler(i as IdleHandler)
		  ' If we already have an idle handler, then we
		  ' cannot install a new one.  Just bail out
		  if mIdleHandlerHook <> 0 then return
		  
		  #if TargetWin32
		    Declare Function SetWindowsHookExA Lib "User32" ( hookType as Integer, proc as Ptr, _
		    instance as Integer, threadID as Integer ) as Integer
		    Declare Function GetCurrentThreadId Lib "Kernel32" () as Integer
		    
		    ' Store the idle handler
		    mIdleHandler = i
		    
		    Const WH_FOREGROUNDIDLE = 11
		    
		    // Well if this isn't about as strange as you can get...  I tried turning this into a
		    // Unicode-savvy function, but couldn't make a go of it.  Using the exact same
		    // code (only the W version of SetWindowsHookEx) causes a crash to occur
		    // immediately after the call returns.  I wasn't able to find any information about
		    // why the crash was happening, and it doesn't make any sense (the Windows is
		    // a Unicode window, so there's no mixed-type calls).  Since this function doesn't
		    // deal with strings, there's no real benefit to making it Unicode-savvy, so I'm leaving
		    // the function as-is.
		    
		    ' And install the handler
		    mIdleHandlerHook= SetWindowsHookExA( WH_FOREGROUNDIDLE, AddressOf DontCallThisIdleHandler, 0, GetCurrentThreadId )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsAlphabetic(s as String) As Boolean
		  #if TargetWin32
		    Declare Function isalpha Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return isalpha( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsHexDigit(s as String) As Boolean
		  #if TargetWin32
		    Declare Function isxdigit Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return isxdigit( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsInf(d as Double) As Boolean
		  #if TargetWin32
		    Declare Function _finite Lib "msvcrt" ( d as Double ) as Boolean
		    
		    return not _finite( d )
		  #endif
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsLowerCase(s as String) As Boolean
		  #if TargetWin32
		    Declare Function islower Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return islower( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsNaN(d as Double) As Boolean
		  #if TargetWin32
		    Declare Function _isnan Lib "msvcrt" ( d as double ) as Boolean
		    
		    return _isnan( d )
		  #endif
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsNumber(s as String) As Boolean
		  #if TargetWin32
		    Declare Function isdigit Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return isdigit( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsPunctuation(s as String) As Boolean
		  #if TargetWin32
		    Declare Function ispunct Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return ispunct( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsUpperCase(s as String) As Boolean
		  #if TargetWin32
		    Declare Function isupper Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return isupper( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsWhiteSpace(s as String) As Boolean
		  #if TargetWin32
		    Declare Function isspace Lib "msvcrt" ( char as Integer ) as Integer
		    
		    dim mb as MemoryBlock
		    mb = Left( s, 1 )
		    
		    try
		      return isspace( mb.Byte( 0 ) ) <> 0
		    catch
		      return false
		    end try
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function LongLongToDouble(mb as MemoryBlock) As Double
		  dim ret as Double
		  
		  // Take the high 4 bytes and shift them
		  ret = mb.Long( 4 ) * (2 ^32)
		  
		  // Then add in the low 4 bytes
		  ret = ret + mb.Long( 0 )
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MakeIntResource(resNum as Integer) As String
		  // Windows is stupid and wants us to typecast an integer as a string
		  dim mb as new MemoryBlock( 4 )
		  mb.LittleEndian = true
		  
		  // MAKEINTRESOURCE does this: (LPTSTR)((DWORD)((WORD) (i)))
		  
		  mb.Short( 0 ) = resNum
		  mb.Long( 0 ) = mb.Short( 0 )
		  
		  return mb.StringValue( 0, 4 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub PressKey(virtualKeyCode as Integer, extendedKey as Boolean = false)
		  #if TargetWin32
		    Declare Sub keybd_event Lib "User32" ( keyCode as Integer, scanCode as Integer, _
		    flags as Integer, extraData as Integer )
		    
		    dim flags as Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    if extendedKey then
		      flags = KEYEVENTF_EXTENDEDKEY
		    end
		    
		    ' Press the key
		    keybd_event( virtualKeyCode, 0, flags, 0 )
		    ' Release the key
		    Const KEYEVENTF_KEYUP = &h2
		    flags = BitwiseOr( flags, KEYEVENTF_KEYUP )
		    keybd_event( virtualKeyCode, 0, flags, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RemoveIdleHandler()
		  ' If we don't have an idle handler, then we
		  ' can just bail out
		  if mIdleHandlerHook = 0 then return
		  
		  #if TargetWin32
		    Declare Sub UnhookWindowsHookEx Lib "User32" ( hookHandle as Integer )
		    
		    UnhookWindowsHookEx( mIdleHandlerHook )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function RestorePointID() As Int64
		  return mRestorePointID
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SRSetRestorePoint(change as Integer, reason as Integer, id as Int64 = 0, description as String = "") As Integer
		  #if TargetWin32
		    Soft Declare Function SRSetRestorePointW Lib "SrClient" ( spec as Ptr, status as Ptr ) as Boolean
		    Soft Declare Function SRSetRestorePointA Lib "SrClient" ( spec as Ptr, status as Ptr ) as Boolean
		    
		    Const BEGIN_SYSTEM_CHANGE = 100
		    Const APPLICATION_INSTALL = 0
		    Const ERROR_SUCCESS = 0
		    if System.IsFunctionAvailable( "SRSetRestorePointW", "SrClient" ) then
		      dim mb as new MemoryBlock( 528 )
		      mb.Long( 0 ) = change
		      mb.Long( 4 ) = reason
		      mb.Int64Value( 8 ) = id
		      if description <> "" then
		        mb.WString( 16 ) = description
		      end if
		      
		      dim status as new MemoryBlock( 12 )
		      dim ret as Boolean = SRSetRestorePointW( mb, status )
		      if not ret or status.Long( 0 ) <> ERROR_SUCCESS then return -1
		      
		      // Save the id off so we know what to close later
		      return status.Int64Value( 4 )
		    else
		      dim mb as new MemoryBlock( 80 )
		      mb.Long( 0 ) = change
		      mb.Long( 4 ) = reason
		      mb.Int64Value( 8 ) = id
		      if description <> "" then
		        mb.CString( 16 ) = description
		      end if
		      
		      dim status as new MemoryBlock( 12 )
		      dim ret as Boolean = SRSetRestorePointA( mb, status )
		      if not ret or status.Long( 0 ) <> ERROR_SUCCESS then return -1
		      
		      // Save the id off so we know what to close later
		      return status.Int64Value( 4 )
		    end if
		  #endif
		  
		  return -1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function StartRestorePoint(type as Integer, description as String) As Boolean
		  // NOTES:
		  // 1) You can only make one restore point at a time, so do not call this recursively.
		  // 2) On ME, you cannot end a restore point while there are still pending renames (which means you have
		  // to reboot the machine before calling the EndRestorePoint method).  So if you have any pending
		  // moves, deletes or renames, you should store the RestorePointID somewhere until after the reboot.  This
		  // isn't something you will encounter often, but it's something for you to be aware of.
		  // 3) This requires 2006r1 or up to compile due to the Int64 support.  You could modify the function to
		  // work with a Double as well, but it's more of a pain than I felt like dealing with.
		  
		  if mRestorePointID <> 0 then return false
		  
		  Const BEGIN_SYSTEM_CHANGE = 100
		  mRestorePointID = SRSetRestorePoint( BEGIN_SYSTEM_CHANGE, type, 0, description )
		  
		  if mRestorePointID = -1 then
		    mRestorePointID = 0
		    return false
		  end if
		  
		  return true
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mIdleHandler As IdleHandler
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIdleHandlerHook As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mRestorePointID As Int64
	#tag EndProperty


	#tag Constant, Name = kRestorePointApplicationInstall, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kRestorePointApplicationUninstall, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kRestorePointDriverInstall, Type = Double, Dynamic = False, Default = \"10", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kRestorePointModifySettings, Type = Double, Dynamic = False, Default = \"12", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateHide, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateMaximize, Type = Integer, Dynamic = False, Default = \"3", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateMinimize, Type = Double, Dynamic = False, Default = \"6", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateRestore, Type = Integer, Dynamic = False, Default = \"9", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateShow, Type = Double, Dynamic = False, Default = \"5", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateShowMaximized, Type = Double, Dynamic = False, Default = \"3", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateShowMinimized, Type = Integer, Dynamic = False, Default = \"2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kWindowStateShowNormal, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant


	#tag ViewBehavior
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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
