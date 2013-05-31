#tag Class
Protected Class ProcessInformation
	#tag Method, Flags = &h0
		Sub BringToFront()
		  #if TargetWin32
		    // First, make sure we've loaded our thread information
		    if UBound( LoadedThreads ) = -1 then LoadThreads
		    
		    // Loop over all the application's threads, telling them to bring
		    // their window handles to the front
		    dim i as Integer
		    for i = 0 to UBound( LoadedThreads )
		      Declare Sub EnumThreadWindows Lib "User32" ( threadID as Integer, proc as Ptr, cookie as Integer )
		      
		      EnumThreadWindows( LoadedThreads( i ).ThreadID, AddressOf BringToFrontCallback, 0 )
		    next i
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(mb as MemoryBlock, unicodeSavvy as Boolean)
		  // Construct one from a memory block
		  
		  me.UnicodeSavvy = unicodeSavvy
		  
		  // First four bytes are the size of the structure, so
		  // just ignore them.
		  
		  // Four bytes of reference count
		  ReferenceCount = mb.Long( 4 )
		  
		  // Four bytes of process id
		  ProcessID = mb.Long( 8 )
		  
		  // Pointer to the default heap id
		  DefaultHeapID = mb.Long( 12 )
		  
		  // Module ID
		  ModuleID = mb.Long( 16 )
		  
		  // Number of threads
		  NumberOfThreads = mb.Long( 20 )
		  
		  // Parent id
		  ParentProcessID = mb.Long( 24 )
		  
		  // The base priority for threads
		  BaseThreadPriority = mb.Long( 28 )
		  
		  // Ignore the next four bytes, they're
		  // reserved.
		  
		  // And finally, the name
		  if unicodeSavvy then
		    Name = mb.WString( 36 )
		  else
		    Name = mb.CString( 36 )
		  end if
		  if Name = "[System Process]" then Name = "System Idle Process"
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadHeapLists()
		  #if TargetWin32
		    Soft Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (flags as Integer, id as Integer ) as Integer
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    Soft Declare Function Heap32ListFirst Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    Soft Declare Function Heap32ListNext Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    
		    Const TH32CS_SNAPHEAPLIST = &h1
		    
		    dim snapHandle as Integer
		    snapHandle = CreateToolhelp32Snapshot( TH32CS_SNAPHEAPLIST, ProcessID )
		    
		    dim mb as new MemoryBlock( 16 )
		    
		    dim entry as HeapListInformation
		    
		    mb.Long( 0 ) = mb.Size
		    if not Heap32ListFirst( snapHandle, mb ) then return
		    
		    dim good as Boolean
		    
		    do
		      entry = new HeapListInformation( mb )
		      
		      LoadedHeapLists.Append( entry )
		      
		      good = Heap32ListNext( snapHandle, mb )
		    loop until not good
		    CloseHandle( snapHandle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadModules()
		  // Because we know things like this process' ID number, we can load
		  // all the modules for the process as well
		  #if TargetWin32
		    Soft Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (flags as Integer, id as Integer ) as Integer
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    Soft Declare Function Module32First Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    Soft Declare Function Module32Next Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    Soft Declare Function Module32FirstW Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    Soft Declare Function Module32NextW Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    
		    Const TH32CS_SNAPMODULE = &h8
		    
		    dim snapHandle as Integer
		    snapHandle = CreateToolhelp32Snapshot( TH32CS_SNAPMODULE, ProcessID )
		    
		    dim mb as MemoryBlock
		    if unicodeSavvy then
		      mb = new MemoryBlock( 32 + 512 + 520 )
		    else
		      mb = new MemoryBlock( 32 + 256 + 260 )
		    end if
		    
		    dim entry as ModuleInformation
		    
		    mb.Long( 0 ) = mb.Size
		    if unicodeSavvy then
		      if not Module32FirstW( snapHandle, mb ) then return
		    else
		      if not Module32First( snapHandle, mb ) then return
		    end if
		    
		    dim good as Boolean
		    
		    do
		      entry = new ModuleInformation( mb, unicodeSavvy )
		      
		      LoadedModules.Append( entry )
		      
		      if unicodeSavvy then
		        good = Module32NextW( snapHandle, mb )
		      else
		        good = Module32Next( snapHandle, mb )
		      end if
		    loop until not good
		    CloseHandle( snapHandle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadThreads()
		  // Because we know things like this process' ID number, we can load
		  // all the thread information for the process as well
		  
		  #if TargetWin32
		    Soft Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (flags as Integer, id as Integer ) as Integer
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    Soft Declare Function Thread32First Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    Soft Declare Function Thread32Next Lib "Kernel32" ( handle as Integer, entry as Ptr ) as Boolean
		    
		    Const TH32CS_SNAPTHREAD = &h4
		    
		    dim snapHandle as Integer
		    snapHandle = CreateToolhelp32Snapshot( TH32CS_SNAPTHREAD, ProcessID )
		    
		    dim mb as new MemoryBlock( 28 )
		    
		    dim entry as ThreadInformation
		    
		    mb.Long( 0 ) = mb.Size
		    if not Thread32First( snapHandle, mb ) then return
		    
		    dim good as Boolean
		    
		    do
		      entry = new ThreadInformation( mb )
		      
		      // For whatever reason, the system will tell us about every thread running
		      // in the entire OS even though we specify the process ID.  This is documented
		      // behavior even though it makes no sense to me.  So we check the thread entry's
		      // process ID to see if it's the same as ours.  If it is, then we keep the thread around.
		      if entry.OwnerProcessID = ProcessID then
		        LoadedThreads.Append( entry )
		      end if
		      
		      good = Thread32Next( snapHandle, mb )
		    loop until not good
		    CloseHandle( snapHandle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Terminate(exitCode as Integer = 0)
		  // Use caution when calling this method because terminating an
		  // application in this way can cause major problems (things aren't
		  // cleaned up properly, the user may lose data, etc).
		  
		  #if TargetWin32
		    Declare Sub TerminateProcess Lib "Kernel32" ( handle as Integer, exitCode as Integer )
		    Declare Function OpenProcess Lib "Kernel32" ( access as Integer, inheritHandle as Boolean, processId as Integer ) as Integer
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    
		    // Open the process up for termination
		    Const PROCESS_TERMINATE = &h1
		    Dim processHandle as Integer = OpenProcess( PROCESS_TERMINATE, false, ProcessID )
		    if processHandle <> 0 then
		      // Terminate the process
		      TerminateProcess( processHandle, exitCode )
		      
		      // Close our handle to the process
		      CloseHandle( processHandle )
		    end if
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		BaseThreadPriority As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		DefaultHeapID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		LoadedHeapLists(-1) As HeapListInformation
	#tag EndProperty

	#tag Property, Flags = &h0
		LoadedModules(-1) As ModuleInformation
	#tag EndProperty

	#tag Property, Flags = &h0
		LoadedThreads(-1) As ThreadInformation
	#tag EndProperty

	#tag Property, Flags = &h0
		ModuleID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberOfThreads As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ParentProcessID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ProcessID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ReferenceCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected unicodeSavvy As Boolean
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="BaseThreadPriority"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DefaultHeapID"
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
			Name="ModuleID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="NumberOfThreads"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ParentProcessID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ProcessID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ReferenceCount"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
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
End Class
#tag EndClass
