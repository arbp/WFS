#tag Module
Protected Module SystemInformation
	#tag Method, Flags = &h1
		Protected Function CPUUsage() As Double
		  Soft Declare Sub NtQuerySystemInformation Lib "ntdll" ( infoClass as Integer, data as Ptr, length as Integer, ByRef retLength as Integer )
		  
		  if not System.IsFunctionAvailable( "NtQuerySystemInformation", "ntdll" ) then return -1.0
		  
		  Const SYSTEM_PERFORMANCEINFORMATION = 2
		  Const SYSTEM_TIMEINFORMATION = 3
		  
		  Dim retStructSize as Integer
		  
		  // First, get the system time information
		  Dim sysTimeInfo as new MemoryBlock( 32 )
		  NtQuerySystemInformation( SYSTEM_TIMEINFORMATION, sysTimeInfo, sysTimeInfo.Size, retStructSize )
		  
		  // Then get the system's idle time information
		  Dim sysIdleTimeInfo as new MemoryBlock( 312 )
		  NtQuerySystemInformation( SYSTEM_PERFORMANCEINFORMATION, sysIdleTimeInfo, sysIdleTimeInfo.Size, retStructSize )
		  
		  // Current value = New value - Old value
		  Static sOldIdleTime, sOldSystemTime as Double
		  
		  // Note, we can just pass this MemoryBlock in because the
		  // LongLongToDouble API assumes that the long long is the
		  // first 8 bytes of the block, which it just so happens to be
		  // in the sysIdleTimeInfo structure.
		  Dim dbIdleTime as Double
		  #if RBVersion < 2006.01 then
		    dbIdleTime = LongLongToDouble( sysIdleTimeInfo ) - sOldIdleTime
		  #else
		    dbIdleTime = sysIdleTimeInfo.UInt64Value( 0 ) - sOldIdleTime
		  #endif
		  
		  // We're not so lucky here -- we need to get the long long into
		  // a new memory block
		  Dim tempMb as new MemoryBlock( 8 )
		  tempMb.Long( 0 ) = sysTimeInfo.Long( 8 )
		  tempMb.Long( 4 ) = sysTimeInfo.Long( 12 )
		  Dim dbSystemTime as Double
		  #if RBVersion < 2006.01 then
		    dbSystemTime = LongLongToDouble( tempMb ) - sOldSystemTime
		  #else
		    dbSystemTime = tempMb.UInt64Value( 0 ) - sOldSystemTime
		  #endif
		  
		  Dim retVal as Double
		  
		  // Divide the idle by the system time
		  if dbSystemTime <> 0 then
		    retVal = dbIdleTime / dbSystemTime
		  else
		    retVal = dbIdleTime
		  end if
		  
		  // Get the number of processors, but cache the information
		  // since it's not going to change while the application is running.
		  Static sNumProcessors as Integer
		  if sNumProcessors = 0 then sNumProcessors = NumberOfProcessors
		  
		  // Now get the true CPU usage time
		  retVal = 100 - retVal * 100 / sNumProcessors + .5
		  
		  // Then store the values for the next query
		  #if RBVersion < 2006.001 then
		    sOldIdleTime = LongLongToDouble( sysIdleTimeInfo )
		    sOldSystemTime = LongLongToDouble( tempMb )
		  #else
		    sOldIdleTime = sysIdleTimeInfo.UInt64Value( 0 )
		    sOldSystemTime = tempMb.UInt64Value( 0 )
		  #endif
		  return retVal / 100
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetComputerName() As String
		  #if TargetWin32
		    Soft Declare Sub GetComputerNameA Lib "Kernel32" ( name as Ptr, ByRef size as Integer )
		    Soft Declare Sub GetComputerNameW Lib "Kernel32" ( name as Ptr, ByRef size as Integer )
		    
		    dim mb as new MemoryBlock( 1024 )
		    dim size as Integer = mb.Size()
		    
		    if System.IsFunctionAvailable( "GetComputerNameW", "Kernel32" ) then
		      GetComputerNameW( mb, size )
		      
		      return mb.WString( 0 )
		    else
		      GetComputerNameA( mb, size )
		      
		      return mb.CString( 0 )
		    end if
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetDefaultPrinter() As String
		  // Function GetWin32DefaultPrinter() As String
		  // DŽbut.: 31/05/2004
		  // Fin...: 01/06/2004
		  // Notes.: Return the default printer name
		  
		  #if TargetWin32
		    // First, try to use the OS-provided function for doing this
		    Soft Declare Function GetDefaultPrinterW Lib "WinSpool.drv" ( buffer as Ptr, ByRef size as Integer ) as Boolean
		    Declare Function GetLastError Lib "Kernel32" () as Integer
		    if System.IsFunctionAvailable( "GetDefaultPrinterW", "WinSpool.drv" ) then
		      // Awesome, let's call it to see how much space we need to allocate
		      Const ERROR_INSUFFICIENT_BUFFER = 122
		      Const ERROR_FILE_NOT_FOUND = 2
		      
		      dim numCharacters as Integer
		      Call GetDefaultPrinterW( nil, numCharacters )
		      select case GetLastError
		      case ERROR_INSUFFICIENT_BUFFER
		        // Now we know how many charactets to allocate so we can continue on
		        dim mb as new MemoryBlock( numCharacters * 2 )
		        if GetDefaultPrinterW( mb, numCharacters ) then
		          // Ta da!  We're golden
		          return mb.WString( 0 )
		        end if
		      case ERROR_FILE_NOT_FOUND
		        // If we got here, then we know there's no default printer and we can just bail
		        return ""
		      end select
		      
		      // The fall-through case means that we're just going to try the profile
		      // string stuff
		    end if
		    
		    Dim slength As Integer                             // receives length of string read from WIN.INI
		    Dim thePrinterName As String
		    dim lpReturnedString As new MemoryBlock( 2048 )
		    
		    Soft Declare Function GetProfileStringW Lib "kernel32" (lpSection As WString, lpKeyName As WString, lpDefault As Integer, lpReturnedString As Ptr, nSize As Integer) As Integer
		    Soft Declare Function GetProfileStringA Lib "kernel32" (lpSection As CString, lpKeyName As CString, lpDefault As Integer, lpReturnedString As Ptr, nSize As Integer) As Integer
		    
		    if System.IsFunctionAvailable( "GetProfileStringW", "Kernel32" ) then
		      slength = GetProfileStringW( "Windows", "Device", 0, lpReturnedString, lpReturnedString.Size \ 2 )
		      
		      if slength > 0 then
		        dim theProfileString as String = lpReturnedString.WString( 0 )
		        thePrinterName = NthField( theProfileString, ",", 1 )
		      end if
		    else
		      slength = GetProfileStringA( "Windows", "Device", 0, lpReturnedString, lpReturnedString.Size )
		      
		      if slength > 0 then
		        thePrinterName = NthField( lpReturnedString.CString( 0 ), ",", 1 )
		      end if
		    end if
		    
		    return thePrinterName
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetDoubleClickTime() As Integer
		  #if TargetWin32
		    Declare Function GetDoubleClickTime Lib "User32" () as Integer
		    
		    return GetDoubleClickTime
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetHostName() As String
		  #if TargetWin32
		    Declare Function gethostname Lib "ws2_32" ( name as Ptr, size as Integer ) as Integer
		    
		    dim mb as new MemoryBlock( 1024 )
		    if gethostname( mb, mb.Size ) = 0 then
		      return mb.CString( 0 )
		    end
		    
		    return ""
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetLoggedInUserName() As String
		  #if TargetWin32
		    Soft Declare Sub GetUserNameA Lib "AdvApi32" ( name as Ptr, ByRef size as Integer )
		    Soft Declare Sub GetUserNameW Lib "AdvApi32" ( name as Ptr, ByRef size as Integer )
		    
		    dim mb as new MemoryBlock( 256 )
		    dim size as Integer = mb.Size()
		    
		    if System.IsFunctionAvailable( "GetUserNameW", "AdvApi32" ) then
		      GetUserNameW( mb, size )
		      
		      return mb.WString( 0 )
		    else
		      GetUserNameA( mb, size )
		      
		      return mb.CString( 0 )
		    end if
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetPrinters() As String()
		  // Function GetWin32DefaultPrinter() As String
		  // DŽbut.: 31/05/2004
		  // Fin...: 02/06/2004
		  // Notes.: Retrun the printers list, delimited by CHR(0)
		  
		  #IF targetWin32
		    Dim slength As Integer                                            // receives length of string read from WIN.INI
		    Dim strAppName, strKeyName, strDefault, thePrinterList As String
		    dim lpAppName, lpKeyName, lpDefault, lpReturnedString As MemoryBlock
		    
		    Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (lpAppName As Ptr, lpKeyName As Integer, lpDefault As Ptr, lpReturnedString As Ptr, nSize As Integer) As Integer
		    
		    strAppName = "PrinterPorts"
		    'strKeyName = ""
		    strDefault = ""
		    
		    lpAppName = new MemoryBlock(lenB(strAppName)+1)
		    'lpKeyName = new MemoryBlock(lenB(strKeyName)+1)
		    lpDefault = new MemoryBlock(lenB(strDefault)+1)
		    lpReturnedString = new MemoryBlock(255)
		    
		    lpAppName.CString(0) = strAppName
		    'lpKeyName.CString(0) = strKeyName
		    lpDefault.CString(0) = strDefault
		    
		    slength = GetProfileString(lpAppName, 0, lpDefault, lpReturnedString, 255)
		    
		    If slength > 0 Then
		      thePrinterList = lpReturnedString.StringValue(0, slength)
		    End If
		    
		    dim ret() as String
		    ret = Split( thePrinterList, Chr( 0 ) )
		    
		    return ret
		    
		  #ENDIF
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetSystemName(root as FolderItem) As String
		  #if TargetWin32
		    Soft Declare Function GetVolumeInformationA Lib "Kernel32" ( root as CString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    Soft Declare Function GetVolumeInformationW Lib "Kernel32" ( root as WString, _
		    volName as Ptr, volNameSize as Integer, ByRef volSer as Integer, ByRef _
		    maxCompLength as Integer, ByRef sysFlags as Integer, sysName as Ptr, _
		    sysNameSize as Integer ) as Boolean
		    
		    
		    dim volName as new MemoryBlock( 256 )
		    dim sysName as new MemoryBlock( 256 )
		    dim volSerial, maxCompLength, sysFlags as Integer
		    
		    if System.IsFunctionAvailable( "GetVolumeInformationW", "Kernel32" ) then
		      Call GetVolumeInformationW( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		      
		      return sysName.WString( 0 )
		    else
		      Call GetVolumeInformationA( left( root.AbsolutePath, 3 ), volName, 256, volSerial, maxCompLength, _
		      sysFlags, sysName, 256 )
		      
		      return sysName.CString( 0 )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsPlatformNT() As Boolean
		  #if TargetWin32
		    Declare Sub GetVersionExA lib "Kernel32" ( info as Ptr )
		    
		    dim info as new MemoryBlock( 148 )
		    info.Long( 0 ) = info.Size
		    
		    GetVersionExA( info )
		    dim str as String
		    
		    if info.Long( 4 ) = 4 then
		      return false
		    elseif info.Long( 4 ) = 3 then
		      return false ' who cares about NT 3?
		    elseif info.Long( 4 ) = 5 then
		      return true
		    end
		    
		    return false
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function NumberOfProcessors() As Integer
		  #if TargetWin32
		    Declare Sub GetSystemInfo Lib "Kernel32" ( info as Ptr )
		    
		    dim info as new MemoryBlock( 9 * 4 )
		    GetSystemInfo( info )
		    
		    return info.Long( 20 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function OSVersionString() As String
		  #if TargetWin32
		    Soft Declare Sub GetVersionExA lib "Kernel32" ( info as Ptr )
		    Soft Declare Sub GetVersionExW lib "Kernel32" ( info as Ptr )
		    
		    dim info as MemoryBlock
		    
		    if System.IsFunctionAvailable( "GetVersionExW", "Kernel32" ) then
		      info =  new MemoryBlock( 20 + (2 * 128) )
		      info.Long( 0 ) = info.Size
		      
		      GetVersionExW( info )
		    else
		      info =  new MemoryBlock( 148 )
		      info.Long( 0 ) = info.Size
		      
		      GetVersionExA( info )
		    end if
		    
		    dim str as String
		    
		    if info.Long( 4 ) = 4 then
		      if info.Long( 8 ) = 0 then
		        str = "Windows 95/NT 4.0"
		      elseif info.Long( 8 ) = 10 then
		        str = "Windows 98"
		      elseif info.Long( 8 ) = 90 then
		        str = "Windows Me"
		      end
		    elseif info.Long( 4 ) = 3 then
		      str = "Windows NT 3.51"
		    elseif info.Long( 4 ) = 5 then
		      if info.Long( 8 ) = 0 then
		        str = "Windows 2000"
		      elseif info.Long( 8 ) = 1 then
		        str = "Windows XP"
		      elseif info.Long( 8 ) = 2 then
		        str = "Windows Server 2003"
		      end
		    end
		    
		    str = str + " Build " + Str( info.Long( 12 ) )
		    
		    if System.IsFunctionAvailable( "GetVersionExW", "Kernel32" ) then
		      str = str + " " + Trim( info.WString( 20 ) )
		    else
		      str = str + " " + Trim( info.CString( 20 ) )
		    end if
		    
		    
		    return str
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SetDefaultPrinter(strPrinterName As String) As Boolean
		  // Function SetWin32DefaultPrinter(strPrinterName As String) As Boolean
		  // DŽbut.: 31/05/2004
		  // Fin...: 31/05/2004
		  // Notes.: Set the default printer to strPrinterName
		  
		  #IF targetWin32
		    Soft Declare Function SetDefaultPrinterW Lib "WinSpool" ( printer as WString ) as Boolean
		    if System.IsFunctionAvailable( "SetDefaultPrinterW", "WinSpool" ) then
		      return SetDefaultPrinterW( strPrinterName )
		    end if
		    
		    // The fall-through case means that we're just going to try the profile
		    // string stuff
		    
		    Dim slength As Integer                             // receives length of string read from WIN.INI
		    Dim strBuffer, strDeviceLine As String
		    dim lpSection, lpKeyName, lpDefault, lpReturnedString, lpString As MemoryBlock
		    dim ok As Boolean
		    
		    Const HWND_BROADCAST  = &HFFFF
		    Const WM_WININICHANGE = &H1A
		    
		    Soft Declare Function GetProfileStringA Lib "kernel32" (lpSection As Ptr, lpKeyName As Ptr, lpDefault As Ptr, lpReturnedString As Ptr, nSize As Integer) As Integer
		    Soft Declare Function WriteProfileStringA Lib "kernel32" (lpSection As Ptr, lpKeyName As Ptr, lpString As Ptr) As Boolean
		    Soft Declare Function SendMessageA Lib "user32" (hwnd As Integer, wMsg As Integer, wParam As Integer, lparam As Ptr) As Boolean
		    Soft Declare Function GetProfileStringW Lib "kernel32" (lpSection As Ptr, lpKeyName As Ptr, lpDefault As Ptr, lpReturnedString As Ptr, nSize As Integer) As Integer
		    Soft Declare Function WriteProfileStringW Lib "kernel32" (lpSection As Ptr, lpKeyName As Ptr, lpString As Ptr) As Boolean
		    Soft Declare Function SendMessageW Lib "user32" (hwnd As Integer, wMsg As Integer, wParam As Integer, lparam As Ptr) As Boolean
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "SendMessageW", "User32" )
		    
		    // Get the full device string
		    // reading PrinterPorts section from win.ini
		    // **************************
		    lpSection = new MemoryBlock( 256 )
		    lpKeyName = new MemoryBlock( 256 )
		    lpDefault = new MemoryBlock( 256 )
		    lpReturnedString = new MemoryBlock(1024)
		    
		    if unicodeSavvy then
		      lpSection.WString(0) = "PrinterPorts"
		      lpKeyName.WString(0) = strPrinterName
		      lpDefault.WString(0) = ""
		      
		      slength = GetProfileStringW(lpSection, lpKeyName, lpDefault, lpReturnedString, 1024)
		    else
		      lpSection.CString(0) = "PrinterPorts"
		      lpKeyName.CString(0) = strPrinterName
		      lpDefault.CString(0) = ""
		      
		      slength = GetProfileStringA(lpSection, lpKeyName, lpDefault, lpReturnedString, 1024)
		    end if
		    
		    
		    // Write out this new printer information in WIN.INI file for DEVICE item
		    // Returned string should be of the form "driver,port,timeout,timeout",
		    // i.e. "winspool,LPT1:,15,45"
		    // **********************************************************************
		    If slength > 0 Then
		      if unicodeSavvy then
		        strBuffer = lpReturnedString.StringValue( 0, slength * 2 )
		      else
		        strBuffer = lpReturnedString.StringValue(0, slength)
		      end if
		      
		      strDeviceLine = strPrinterName + "," + NthField(strBuffer, Chr(0), 1) + "," + NthField(strBuffer, Chr(0), 2)
		      
		      lpSection = new MemoryBlock( 256 )
		      lpKeyName = new MemoryBlock( 256 )
		      lpString  = new MemoryBlock( 1024 )
		      
		      if unicodeSavvy then
		        lpSection.WString(0) = "Windows"
		        lpKeyName.WString(0) = "Device"
		        lpString.WString(0)  = strDeviceLine
		        
		        ok = WriteProfileStringW(lpSection, lpKeyName, lpString)
		      else
		        lpSection.CString(0) = "Windows"
		        lpKeyName.CString(0) = "Device"
		        lpString.CString(0)  = strDeviceLine
		        
		        ok = WriteProfileStringA(lpSection, lpKeyName, lpString)
		      end if
		      
		      
		      // Notify the changes
		      // ******************
		      // Below is optional, and should be done. It updates the existing windows
		      // so the "default" printer icon changes. If you don't do the below..then
		      // you will often see more than one printer as the default! The reason *not*
		      // to do the SendMessage is that many open applications will now sense the change
		      // in printer. I vote to leave it in..but your case you might not want this.
		      
		      if ok then
		        if unicodeSavvy then
		          ok = SendMessageW(HWND_BROADCAST, WM_WININICHANGE, 0, lpSection)
		        else
		          ok = SendMessageA(HWND_BROADCAST, WM_WININICHANGE, 0, lpSection)
		        end if
		      end if
		      
		    Else
		      ok = False
		    End If
		    
		    RETURN ok
		    
		  #ENDIF
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SetScreenResolution(width as Integer, height as Integer, depth as Integer) As Integer
		  Soft Declare Function EnumDisplaySettingsW Lib "user32" ( null as Integer, iModeNum As Integer, lpDevMode As Ptr ) As Integer
		  Soft Declare Function EnumDisplaySettingsA Lib "user32" ( null as Integer, iModeNum As Integer, lpDevMode As Ptr ) As Integer
		  
		  Soft Declare Function ChangeDisplaySettingsW Lib "user32" ( lpDevMode As Ptr,  dwFlags As Integer ) As Integer
		  Soft Declare Function ChangeDisplaySettingsA Lib "user32" ( lpDevMode As Ptr,  dwFlags As Integer ) As Integer
		  
		  Const ENUM_CURRENT_SETTINGS = -1
		  
		  Const CDS_UPDATEREGISTRY = &H1
		  Const CDS_TEST = &H2
		  Const DISP_CHANGE_SUCCESSFUL = 0
		  Const DISP_CHANGE_RESTART = 1
		  
		  // First, figure out whether we're unicode savvy or not
		  Dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "EnumDisplaySettingsW", "User32" )
		  
		  // Now, allocate a DEVMODE structure and set its dmSize property
		  dim devMode as MemoryBlock
		  if unicodeSavvy then
		    devMode = new MemoryBlock( 188 )
		    devMode.Long( 68 ) = devMode.Size
		  else
		    devMode = new MemoryBlock( 124 )
		    devMode.Long( 36 ) = devMode.Size
		  end if
		  
		  dim offset as Integer = devMode.Size - 20
		  
		  dim retVal as Integer
		  if unicodeSavvy then
		    retVal = EnumDisplaySettingsW( 0, ENUM_CURRENT_SETTINGS, devMode )
		  else
		    retVal = EnumDisplaySettingsA( 0, ENUM_CURRENT_SETTINGS, devMode )
		  end if
		  
		  if retVal = 0 then return kResolutionChangeFailed
		  
		  // Make the requested changes
		  devMode.Long( offset ) = depth
		  devMode.Long( offset + 4 ) = width
		  devMode.Long( offset + 8 ) = height
		  
		  // Check to see whether we can make the change
		  if unicodeSavvy then
		    retVal = ChangeDisplaySettingsW( devMode, CDS_TEST )
		  else
		    retVal = ChangeDisplaySettingsA( devMode, CDS_TEST )
		  end if
		  
		  if retVal <> DISP_CHANGE_SUCCESSFUL then return kResolutionChangeFailed
		  
		  // We're able to make the change, so let's actually do the change
		  if unicodeSavvy then
		    retVal = ChangeDisplaySettingsW( devMode, CDS_UPDATEREGISTRY )
		  else
		    retVal = ChangeDisplaySettingsA( devMode, CDS_UPDATEREGISTRY )
		  end if
		  
		  select case retVal
		  case DISP_CHANGE_RESTART
		    return kResolutionChangeRequiresReboot
		    
		  case DISP_CHANGE_SUCCESSFUL
		    return kResolutionChangeSuccess
		    
		  else
		    return kResolutionChangeFailed
		  end select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Shutdown(mode as Integer) As Boolean
		  #if TargetWin32
		    Declare Function ExitWindowsEx Lib "User32" ( flags as Integer, zero as Integer ) as Boolean
		    
		    return ExitWindowsEx( mode, 0 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Shutdown(owner as Window)
		  // This is similar to the other shutdown with UI function, except that it doesn't
		  // tell you about the user's selection, and it shows the same dialog you woud see
		  // from going to Start->Shutdown  There's no way to figure out what the user selected
		  
		  #if TargetWin32
		    // Let's start hacking!  Load the function pointer up via an ordinal.  Note, this function is
		    // not documented by Microsoft, nor is it officially sanctioned.  We'll use it anyways because we're cool like that.
		    Soft Declare Sub ExitWindowsDialog Lib "Shell32" Alias "#60" ( owner as Integer )
		    
		    dim ownerHandle as Integer
		    if owner <> nil then ownerHandle = owner.WinHWND
		    
		    ExitWindowsDialog( ownerHandle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Shutdown(owner as Window, reason as String, type as Integer) As Boolean
		  #if TargetWin32
		    // Let's start hacking!  Load the function pointer up via an ordinal.  Note, this function is
		    // not documented by Microsoft, nor is it officially sanctioned.  We'll use it anyways because we're cool like that.
		    Soft Declare Function RestartDialog Lib "Shell32" Alias "#59" ( owner as Integer, reason as Ptr, flags as Integer ) as Integer
		    
		    dim ownerHandle as Integer
		    if owner <> nil then ownerHandle = owner.WinHWND
		    
		    reason = reason + EndOfLine + EndOfLine
		    
		    dim retVal as Integer
		    dim msg as MemoryBlock
		    if IsPlatformNT then
		      msg = ConvertEncoding( reason + Chr( 0 ), Encodings.UTF16 )
		    else
		      msg = ConvertEncoding( reason + Chr( 0 ), Encodings.ASCII )
		    end if
		    
		    retVal = RestartDialog( ownerHandle, msg, type )
		    
		    Const IDOK = 1
		    Const IDCANCEL = 2
		    
		    return retVal = IDOK
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SystemMasterVolume() As Integer
		  #if TargetWin32
		    Declare Function mixerOpen Lib "winmm" ( ByRef handle as Integer, id as Integer, _
		    callback as Integer, instance as Integer, open as Integer ) as Integer
		    Declare Function mixerGetNumDevs Lib "winmm" () as Integer
		    Declare Function mixerGetControlDetailsA Lib "winmm" ( handle as Integer, details as Ptr, flags as Integer ) as Integer
		    Declare Sub mixerGetLineInfoA Lib "winmm" ( handle as Integer, line as Ptr, flags as Integer )
		    Declare Sub mixerGetLineControlsA Lib "winmm" ( handle as Integer, lineCtl as Ptr, flags as Integer )
		    
		    dim i, count as Integer
		    count = mixerGetNumDevs - 1
		    
		    dim device as Integer
		    for i = 0 to count
		      if mixerOpen( device, i, 0, 0, 0 ) = 0 then
		        exit
		      end
		    next
		    
		    ' Get the line information for the Speakers
		    dim lineThinger as new MemoryBlock( 80 + 40 + 32 + 16 )
		    lineThinger.Long( 0 ) = lineThinger.Size
		    lineThinger.Long( 24 ) = 4
		    mixerGetLineInfoA( device, lineThinger, 3 )
		    
		    ' Get the volume control for the speakers
		    dim otherLineThinger as new MemoryBlock( 24 )
		    dim mixerControl as new MemoryBlock( 80 + (18 * 4))
		    otherLineThinger.Long( 0 ) = otherLineThinger.Size
		    otherLineThinger.Long( 4 ) = lineThinger.Long( 12 )
		    otherLineThinger.Long( 8 ) = &h50000000 + &h30000 + 1
		    otherLineThinger.Long( 12 ) = 1
		    otherLineThinger.Long( 16 ) = mixerControl.Size
		    otherLineThinger.Ptr( 20 ) = mixerControl
		    mixerGetLineControlsA( device, otherLineThinger, 2 )
		    
		    dim details as new MemoryBlock( 24 )
		    dim vals as new MemoryBlock( 4 )
		    
		    details.Long( 0 ) = details.Size
		    details.Long( 4 ) = mixerControl.Long( 4 )
		    details.Long( 8 ) = 1
		    details.Long( 16 ) = 4
		    details.Ptr( 20 ) = vals
		    
		    Call mixerGetControlDetailsA( device, details, 0 )
		    
		    return vals.Long( 0 )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SystemMasterVolume(assigns vol as Integer)
		  #if TargetWin32
		    Declare Function mixerOpen Lib "winmm" ( ByRef handle as Integer, id as Integer, _
		    callback as Integer, instance as Integer, open as Integer ) as Integer
		    Declare Function mixerGetNumDevs Lib "winmm" () as Integer
		    Declare Function mixerSetControlDetails Lib "winmm" ( handle as Integer, details as Ptr, flags as Integer ) as Integer
		    Declare Sub mixerGetLineInfoA Lib "winmm" ( handle as Integer, line as Ptr, flags as Integer )
		    Declare Sub mixerGetLineControlsA Lib "winmm" ( handle as Integer, lineCtl as Ptr, flags as Integer )
		    
		    dim i, count as Integer
		    count = mixerGetNumDevs - 1
		    
		    dim device as Integer
		    for i = 0 to count
		      if mixerOpen( device, i, 0, 0, 0 ) = 0 then
		        exit
		      end
		    next
		    
		    ' Get the line information for the Speakers
		    dim lineThinger as new MemoryBlock( 80 + 40 + 32 + 16 )
		    lineThinger.Long( 0 ) = lineThinger.Size
		    lineThinger.Long( 24 ) = 4
		    mixerGetLineInfoA( device, lineThinger, 3 )
		    
		    ' Get the volume control for the speakers
		    dim otherLineThinger as new MemoryBlock( 24 )
		    dim mixerControl as new MemoryBlock( 80 + (18 * 4))
		    otherLineThinger.Long( 0 ) = otherLineThinger.Size
		    otherLineThinger.Long( 4 ) = lineThinger.Long( 12 )
		    otherLineThinger.Long( 8 ) = &h50000000 + &h30000 + 1
		    otherLineThinger.Long( 12 ) = 1
		    otherLineThinger.Long( 16 ) = mixerControl.Size
		    otherLineThinger.Ptr( 20 ) = mixerControl
		    mixerGetLineControlsA( device, otherLineThinger, 2 )
		    
		    dim details as new MemoryBlock( 24 )
		    dim vals as new MemoryBlock( 4 )
		    vals.Long( 0 ) = vol
		    
		    details.Long( 0 ) = details.Size
		    details.Long( 4 ) = mixerControl.Long( 4 )
		    details.Long( 8 ) = 1
		    details.Long( 16 ) = 4
		    details.Ptr( 20 ) = vals
		    
		    dim ret as Integer
		    ret = mixerSetControlDetails( device, details, 0 )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function TimeZoneInformation() As TimeZoneInformation
		  #if TargetWin32
		    Declare Function GetTimeZoneInformation Lib "Kernel32" ( tzi as Ptr ) as Integer
		    
		    Dim tzi as new MemoryBlock( 172 )
		    dim current as Integer
		    current = GetTimeZoneInformation( tzi )
		    
		    if current = -1 then return nil
		    
		    dim ret as new TimeZoneInformation
		    
		    ret.Bias = tzi.Long( 0 )
		    
		    ret.StandardName = DefineEncoding( tzi.StringValue( 4, 64 ), Encodings.UTF16 )
		    ret.StandardBias = tzi.Long( 84  )
		    
		    ret.DaylightName = DefineEncoding( tzi.StringValue( 88, 64 ), Encodings.UTF16 )
		    ret.DaylightBias = tzi.Long( 168 )
		    ret.Current = current
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WavVolume() As Integer
		  #if TargetWin32
		    Declare Sub waveOutGetVolume Lib "winmm" ( devid as Integer, ByRef vol as Integer )
		    Declare Function waveOutOpen Lib "winmm" ( ByRef out as Integer, devid as Integer, format as Ptr, _
		    callback as Integer, instance as Integer, open as Integer ) as Integer
		    Declare Function waveOutGetNumDevs Lib "winmm" () as Integer
		    
		    dim format as new MemoryBlock( 18 )
		    format.Short( 0 ) = 1                ' wFormatTag
		    format.Short( 2 ) = 1                ' nChannels
		    format.Long( 4 ) = 44100       ' nSamplesPerSecond
		    format.Short( 14 ) = 8              ' wBitsPerSample
		    format.Short( 12 ) = 2              ' nBlockAlign
		    format.Long( 8 ) = 2 * 44100 ' nAvgBytesPerSec
		    
		    dim i, count, ret as Integer
		    dim device as Integer
		    count = waveOutGetNumDevs - 1
		    
		    for i = 0 to count
		      ret = waveOutOpen( device, i, format, 0, 0, 1 )
		      if ret = 0 then
		        exit
		      end
		    next
		    
		    dim vol as Integer
		    waveOutGetVolume( device, vol )
		    
		    return vol
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WavVolume(assigns vol as Integer)
		  #if TargetWin32
		    Declare Function waveOutOpen Lib "winmm" ( ByRef out as Integer, devid as Integer, format as Ptr, _
		    callback as Integer, instance as Integer, open as Integer ) as Integer
		    Declare Function waveOutGetNumDevs Lib "winmm" () as Integer
		    Declare Function waveOutSetVolume Lib "winmm" ( handle as Integer,  vol as Integer ) as Integer
		    
		    dim format as new MemoryBlock( 18 )
		    format.Short( 0 ) = 1                ' wFormatTag
		    format.Short( 2 ) = 1                ' nChannels
		    format.Long( 4 ) = 44100       ' nSamplesPerSecond
		    format.Short( 14 ) = 8              ' wBitsPerSample
		    format.Short( 12 ) = 2              ' nBlockAlign
		    format.Long( 8 ) = 2 * 44100 ' nAvgBytesPerSec
		    
		    dim i, count, ret as Integer
		    dim device as Integer
		    count = waveOutGetNumDevs - 1
		    
		    for i = 0 to count
		      ret = waveOutOpen( device, i, format, 0, 0, 1 )
		      if ret = 0 then
		        exit
		      end
		    next
		    
		    Call waveOutSetVolume( device, vol )
		  #endif
		End Sub
	#tag EndMethod


	#tag Constant, Name = kResolutionChangeFailed, Type = Double, Dynamic = False, Default = \"-1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kResolutionChangeRequiresReboot, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kResolutionChangeSuccess, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kShutdownModeLogoff, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kShutdownModePowerOff, Type = Double, Dynamic = False, Default = \"3", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kShutdownModeReboot, Type = Double, Dynamic = False, Default = \"2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kShutdownModeShutdown, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
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
