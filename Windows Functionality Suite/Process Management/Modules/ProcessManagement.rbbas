#tag Module
Protected Module ProcessManagement
	#tag Method, Flags = &h1
		Protected Function ApplicationPriority() As Integer
		  #if TargetWin32
		    Declare Function OpenProcess Lib "Kernel32" ( access as Integer, inherit as Boolean, procID as Integer ) as Integer
		    Declare Function GetPriorityClass Lib "Kernel32" ( handle as Integer ) as Integer
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    
		    ' Get a handle to the current process
		    dim processHandle as Integer
		    const PROCESS_QUERY_INFORMATION = &h400
		    processHandle = OpenProcess( PROCESS_QUERY_INFORMATION, false, GetCurrentProcessID )
		    
		    ' Get the priority
		    dim ret as Integer
		    ret = GetPriorityClass( processHandle )
		    
		    ' And close the handle to the module
		    CloseHandle( processHandle )
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ApplicationPriority(assigns level as Integer)
		  #if TargetWin32
		    Declare Function OpenProcess Lib "Kernel32" ( access as Integer, inherit as Boolean, procID as Integer ) as Integer
		    Declare Sub SetPriorityClass Lib "Kernel32" ( handle as Integer, priority as Integer )
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    
		    ' Get a handle to the current process
		    dim processHandle as Integer
		    const PROCESS_SET_INFORMATION = &h200
		    processHandle = OpenProcess( PROCESS_SET_INFORMATION, false, GetCurrentProcessID )
		    
		    ' And set the priority
		    SetPriorityClass( processHandle, level )
		    
		    ' And close the handle to the module
		    CloseHandle( processHandle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub BringToFront(process as ProcessInformation)
		  #if TargetWin32
		    // First, make sure we've loaded our thread information
		    if UBound( process.LoadedThreads ) = -1 then process.LoadThreads
		    
		    // Loop over all the application's threads, telling them to bring
		    // their window handles to the front
		    dim i as Integer
		    for i = 0 to UBound( process.LoadedThreads )
		      Declare Sub EnumThreadWindows Lib "User32" ( threadID as Integer, proc as Ptr, cookie as Integer )
		      
		      EnumThreadWindows( process.LoadedThreads( i ).ThreadID, AddressOf BringToFrontCallback, 0 )
		    next i
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BringToFrontCallback(hwnd as Integer, cookie as Integer) As Boolean
		  #if TargetWin32
		    Declare Sub BringWindowToTop Lib "User32" ( hwnd as Integer )
		    
		    BringWindowToTop( hwnd )
		    
		    Return true
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetActiveProcesses() As ProcessInformation()
		  #if TargetWin32
		    Soft Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (flags as Integer, id as Integer ) as Integer
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    Soft Declare Sub Process32First Lib "Kernel32" ( handle as Integer, entry as Ptr )
		    Soft Declare Function Process32Next Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    Soft Declare Sub Process32FirstW Lib "Kernel32" ( handle as Integer, entry as Ptr )
		    Soft Declare Function Process32NextW Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    
		    dim snapHandle as Integer
		    Const TH32CS_SNAPPROCESS = &h2
		    snapHandle = CreateToolhelp32Snapshot( TH32CS_SNAPPROCESS, 0 )
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "Process32FirstW", "Kernel32" )
		    dim mb as MemoryBlock
		    if unicodeSavvy then
		      mb = new MemoryBlock( (260 * 2) + 36 )
		    else
		      mb = new MemoryBlock( 260 + 36 )
		    end if
		    
		    dim entry as ProcessInformation
		    dim err as Integer
		    dim ret() as ProcessInformation
		    
		    mb.Long( 0 ) = mb.Size
		    if unicodeSavvy then
		      Process32FirstW( snapHandle, mb )
		    else
		      Process32First( snapHandle, mb )
		    end if
		    
		    dim good as Boolean
		    
		    do
		      entry = new ProcessInformation( mb, unicodeSavvy )
		      
		      ret.Append( entry )
		      
		      if unicodeSavvy then
		        good = Process32NextW( snapHandle, mb )
		      else
		        good = Process32Next( snapHandle, mb )
		      end if
		    loop until not good
		    CloseHandle( snapHandle )
		    
		    return ret
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetActiveProcessNames() As String()
		  #if TargetWin32
		    dim ret() as String
		    dim entries() as ProcessInformation
		    dim entry as ProcessInformation
		    
		    // Get all of the processes
		    entries = GetActiveProcesses()
		    
		    // Add their names to our return list
		    for each entry in entries
		      ret.Append( entry.Name )
		    next entry
		    
		    return ret
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCurrentProcessID() As Integer
		  #if TargetWin32
		    Declare Function MyGetCurrentProcessId Lib "Kernel32" Alias "GetCurrentProcessId" () as Integer
		    
		    return MyGetCurrentProcessID
		  #endif
		  return 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetFrontmostWindowHandle() As Integer
		  #if TargetWin32
		    Declare Function GetForegroundWindow Lib "User32" () as Integer
		    
		    return GetForegroundWindow
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetFrontmostWindowProcessInformation() As ProcessInformation
		  #if TargetWin32
		    // We want to get the frontmost window based
		    dim frontWindowHWND as Integer
		    frontWindowHWND = GetFrontmostWindowHandle
		    
		    if frontWindowHWND <= 0 then return nil
		    
		    // Now figure out what process ID owns the window
		    Declare Sub GetWindowThreadProcessId  Lib "User32" ( hwnd as Integer, ByRef procId as Integer )
		    dim processID as Integer
		    GetWindowThreadProcessId( frontWindowHWND, processId )
		    
		    // Now we get a list of all the processes, and see if we can find
		    // one with a match.
		    dim processes(-1) as ProcessInformation
		    processes = GetActiveProcesses
		    
		    dim ret as ProcessInformation
		    for each ret in processes
		      if ret.ProcessID = processID then
		        return ret
		      end if
		    next
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub HideWin9xProcess(procID as Integer, hide as Boolean = true)
		  #if TargetWin32
		    Declare Sub RegisterServiceProcess Lib "Kernel32" ( id as Integer, reg as Boolean )
		    
		    RegisterServiceProcess( procID, hide )
		  #endif
		End Sub
	#tag EndMethod


	#tag Constant, Name = kPriorityHigh, Type = Double, Dynamic = False, Default = \"&h80", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kPriorityIdle, Type = Double, Dynamic = False, Default = \"&h40", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kPriorityNormal, Type = Double, Dynamic = False, Default = \"&h20", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kPriorityRealTime, Type = Double, Dynamic = False, Default = \"&h100", Scope = Protected
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
