#tag Class
Protected Class Win32Mutex
	#tag Method, Flags = &h1
		Protected Sub Close()
		  #if TargetWin32
		    Declare Sub CloseHandle Lib "Kernel32" ( handle as Integer )
		    
		    if mHandle <> 0 then
		      CloseHandle( mHandle )
		    end
		  #endif
		  
		  mHandle = 0
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(name as String, globalMutex as Boolean = false)
		  mHandle = 0
		  mName = name
		  mGlobal = globalMutex
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  Release
		  Close
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Handle() As Integer
		  return mHandle
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Obtain() As Boolean
		  #if TargetWin32
		    Soft Declare Function CreateMutexA Lib "Kernel32" ( security as Integer, _
		    owner as Boolean, name as CString ) as Integer
		    Soft Declare Function CreateMutexW Lib "Kernel32" ( security as Integer, _
		    owner as Boolean, name as WString ) as Integer
		    
		    
		    dim useName as String
		    if mGlobal then
		      useName = "Global\" + mName
		    else
		      useName = mName
		    end
		    
		    if System.IsFunctionAvailable( "CreateMutexW", "Kernel32" ) then
		      mHandle = CreateMutexW( 0, true, useName )
		    else
		      mHandle = CreateMutexA( 0, true, useName )
		    end if
		    
		    const ERROR_ALREADY_EXISTS = 183
		    dim error as Integer
		    if mHandle = 0 then
		      return false
		    else
		      error = Win32DeclareLibrary.GetLastError
		      if error = ERROR_ALREADY_EXISTS then
		        return false
		      end
		      return true
		    end
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Release()
		  if mHandle = 0 then return
		  
		  #if TargetWin32
		    Declare Sub ReleaseMutex Lib "Kernel32" ( handle as Integer )
		    
		    ReleaseMutex( mHandle )
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected mGlobal As Boolean
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mName As String
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
End Class
#tag EndClass
