#tag Class
Protected Class Service
	#tag Method, Flags = &h1
		Protected Sub InternalControlService(handle as Integer, controlCode as Integer)
		  #if TargetWin32
		    Declare Sub ControlService Lib "AdvApi32" ( handle as Integer, code as Integer, status as Ptr )
		    
		    dim status as new MemoryBlock( 7 * 4 )
		    ControlService( handle, controlCode, status )
		    
		    // Reset all the state variables
		    mIsStopped = False
		    mIsRunning = False
		    mIsStarting = False
		    mIsStopping = False
		    mIsPaused = False
		    mIsPausing = False
		    mIsContinuing = False
		    
		    // Check the status byte
		    select case status.Byte( 4 )
		    case 1
		      mIsStopped = True
		    case 2
		      mIsStarting = True
		    case 3
		      mIsStopping = True
		    case 4
		      mIsRunning = True
		    case 5
		      mIsContinuing = True
		    case 6
		      mIsPausing = True
		    case 7
		      mIsPaused = True
		    end select
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsContinuing() As Boolean
		  Query
		  
		  return mIsContinuing
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsPaused() As Boolean
		  Query
		  
		  return mIsPaused
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsPausing() As Boolean
		  Query
		  
		  return mIsPausing
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsRunning() As Boolean
		  Query
		  
		  return mIsRunning
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsStarting() As Boolean
		  Query
		  
		  return mIsStarting
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsStopped() As Boolean
		  Query
		  
		  return mIsStopped
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsStopping() As Boolean
		  Query
		  
		  return mIsStopping
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Pause()
		  InternalControlService( Handle, kControlPause )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Query()
		  InternalControlService( Handle, kControlInterrogate )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Resume()
		  InternalControlService( Handle, kControlContinue )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Start(ParamArray params as String)
		  #if TargetWin32
		    Dim theParams, param as String
		    Dim numParams as Integer
		    dim paramsPtr as new MemoryBlock( 4 )
		    dim paramsCStr, paramsWStr as MemoryBlock
		    
		    // Check to see whether we have already gotten a handle to the
		    // service
		    if Handle = 0 then
		      return
		    else
		      Soft Declare Sub StartServiceA Lib "AdvApi32" ( handle as Integer, numParams as Integer, _
		      params as Ptr )
		      Soft Declare Sub StartServiceW Lib "AdvApi32" ( handle as Integer, numParams as Integer, _
		      params as Ptr )
		      
		      numParams = UBound( params )
		      
		      if numParams < 0 then numParams = 0
		      
		      for each param in params
		        theParams = theParams + param + Chr( 0 )
		      next
		      
		      if System.IsFunctionAvailable( "StartServiceW", "AdvApi32" ) then
		        paramsWStr = new MemoryBlock( (Len( theParams ) + 1) * 2 )
		        paramsWStr.WString( 0 ) = theParams
		        paramsPtr.Ptr( 0 ) = paramsWStr
		        StartServiceW( Handle, numParams, paramsPtr )
		      else
		        paramsCStr = new MemoryBlock( Len( theParams ) + 1 )
		        paramsCStr.CString( 0 ) = theParams
		        paramsPtr.Ptr( 0 ) = paramsCStr
		        StartServiceA( Handle, numParams, paramsPtr )
		      end if
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Stop()
		  InternalControlService( Handle, kControlStop )
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Dependencies As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Description As String
	#tag EndProperty

	#tag Property, Flags = &h0
		DisplayName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		ErrorControl As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ExecutableFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Handle As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		LoadOrderGroup As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsContinuing As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsPaused As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsPausing As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsRunning As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsStarting As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsStopped As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsStopping As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Password As String
	#tag EndProperty

	#tag Property, Flags = &h0
		StartName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		StartType As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		StartupParameters As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Type As Integer
	#tag EndProperty


	#tag Constant, Name = kControlContinue, Type = Double, Dynamic = False, Default = \"&H3", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kControlInterrogate, Type = Double, Dynamic = False, Default = \"&h4", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kControlPause, Type = Integer, Dynamic = False, Default = \"&H2", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kControlStop, Type = Integer, Dynamic = False, Default = \"&H1", Scope = Private
	#tag EndConstant

	#tag Constant, Name = kErrorControlCritical, Type = Integer, Dynamic = False, Default = \"&h3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kErrorControlIgnore, Type = Integer, Dynamic = False, Default = \"&h0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kErrorControlNormal, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kErrorControlSevere, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStartTypeAuto, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStartTypeBoot, Type = Integer, Dynamic = False, Default = \"&h0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStartTypeDemand, Type = Integer, Dynamic = False, Default = \"&h3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStartTypeDisabled, Type = Integer, Dynamic = False, Default = \"&h4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStartTypeSystem, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTypeFileSystemDriver, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTypeInteractiveProcess, Type = Integer, Dynamic = False, Default = \"&h100", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTypeKernelDriver, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTypeOwnProcess, Type = Integer, Dynamic = False, Default = \"&h10", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTypeShareProcess, Type = Integer, Dynamic = False, Default = \"&h20", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Dependencies"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Description"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DisplayName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ErrorControl"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
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
			Name="LoadOrderGroup"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Password"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="StartName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="StartType"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="StartupParameters"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
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
		#tag ViewProperty
			Name="Type"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
