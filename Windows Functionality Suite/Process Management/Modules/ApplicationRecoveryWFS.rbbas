#tag Module
Protected Module ApplicationRecoveryWFS
	#tag Method, Flags = &h21
		Private Function RecoveryCallback(param as UInt32) As UInt32
		  // This callback method must be a StdCall
		  #pragma X86CallingConvention StdCall
		  
		  mCallback.RecoveryCallback( param )
		  
		Exception
		  // If anything bad happened, we failed
		  RecoveryFinished( false )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RecoveryFinished(success as Boolean)
		  #if TargetWin32
		    
		    Soft Declare Sub ApplicationRecoveryFinished Lib "Kernel32" ( success as Boolean )
		    if System.IsFunctionAvailable( "ApplicationRecoveryFinished", "Kernel32" ) then
		      ApplicationRecoveryFinished( success )
		    end if
		    
		  #else
		    
		    #pragma unused success
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function RecoveryInProgress() As Boolean
		  #if TargetWin32
		    Soft Declare Sub ApplicationRecoveryInProgress Lib "Kernel32" ( ByRef cancelled as Boolean )
		    
		    if System.IsFunctionAvailable( "ApplicationRecoveryInProgress", "Kernel32" ) then
		      dim ret as Boolean
		      ApplicationRecoveryInProgress( ret )
		      return ret
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RegisterRecoveryCallback(callback as ApplicationRecoveryCallbackProviderWFS, param as UInt32)
		  #if TargetWin32
		    
		    Soft Declare Sub RegisterApplicationRecoveryCallback Lib "Kernel32" ( callback as Ptr, param as UInt32, ping as UInt32, flags as UInt32 )
		    
		    if callback <> nil and System.IsFunctionAvailable( "RegisterApplicationRecoveryCallback", "Kernel32" ) then
		      mCallback = callback
		      RegisterApplicationRecoveryCallback( AddressOf RecoveryCallback, param, 0, 0 )
		    end if
		    
		  #else
		    
		    #pragma unused callback
		    #pragma unused param
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RegisterRestart(commandLine as String, flags as Integer)
		  #if TargetWin32
		    
		    Soft Declare Sub RegisterApplicationRestart Lib "Kernel32" ( cmdLine as WString, flags as UInt32 )
		    
		    if System.IsFunctionAvailable( "RegisterApplicationRestart", "Kernel32" ) then
		      RegisterApplicationRestart( commandLine, flags )
		    end if
		    
		  #else
		    
		    #pragma unused commandLine
		    #pragma unused flags
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub UnregisterRestart()
		  #if TargetWin32
		    Soft Declare Sub UnregisterApplicationRestart Lib "Kernel32" ()
		    
		    if System.IsFunctionAvailable( "UnregisterApplicationRestart", "Kernel32" ) then
		      UnregisterApplicationRestart
		    end if
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mCallback As ApplicationRecoveryCallbackProviderWFS
	#tag EndProperty


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
