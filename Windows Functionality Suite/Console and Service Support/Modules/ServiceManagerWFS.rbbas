#tag Module
Protected Module ServiceManagerWFS
	#tag Method, Flags = &h1
		Protected Sub AddService(ByRef serv as ServiceWFS)
		  #if TargetWin32
		    
		    Soft Declare Function CreateServiceA Lib "AdvApi32" ( manager as Integer, name as CString, _
		    displayName as CString, access as Integer, serviceType as Integer, startType as Integer, _
		    errorControl as Integer, binaryPath as CString, loadOrder as CString, tagID as Integer, _
		    dependencies as CString, accountName as CString, password as CString ) as Integer
		    Soft Declare Function CreateServiceW Lib "AdvApi32" ( manager as Integer, name as WString, _
		    displayName as WString, access as Integer, serviceType as Integer, startType as Integer, _
		    errorControl as Integer, binaryPath as WString, loadOrder as WString, tagID as Integer, _
		    dependencies as WString, accountName as WString, password as WString ) as Integer
		    
		    if mManager = 0 then return
		    
		    dim path as String
		    path = """" + serv.ExecutableFile.AbsolutePath + """"
		    
		    if serv.StartupParameters <> "" then
		      path = path + " " + serv.StartupParameters
		    end
		    
		    if System.IsFunctionAvailable( "CreateServiceW", "AdvApi32" ) then
		      serv.Handle = CreateServiceW( mManager, serv.Name, serv.DisplayName, kAccessServiceAll, serv.Type, _
		      serv.StartType, serv.ErrorControl, path, serv.LoadOrderGroup, _
		      0, serv.Dependencies, serv.StartName, serv.Password )
		    else
		      serv.Handle = CreateServiceA( mManager, serv.Name, serv.DisplayName, kAccessServiceAll, serv.Type, _
		      serv.StartType, serv.ErrorControl, path, serv.LoadOrderGroup, _
		      0, serv.Dependencies, serv.StartName, serv.Password )
		    end if
		    
		    if serv.Handle <> 0 then
		      ModifyService( serv )
		    end
		    
		  #else
		    
		    #pragma unused serv
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub CloseService(serv as ServiceWFS)
		  #if TargetWin32
		    
		    CloseServiceHandle( serv.Handle )
		    
		  #else
		    
		    #pragma unused serv
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub CloseServiceHandle(handle as Integer)
		  #if TargetWin32
		    
		    Declare Sub CloseServiceHandle Lib "AdvApi32" ( handle as Integer )
		    
		    CloseServiceHandle( handle )
		    
		  #else
		    
		    #pragma unused handle
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub FillServiceInformation(ByRef serv as ServiceWFS)
		  #if TargetWin32
		    
		    Soft Declare Function QueryServiceConfigA Lib "AdvApi32" ( handle as Integer, _
		    config as Ptr, size as Integer, ByRef bytesNeeded as Integer ) as Boolean
		    Soft Declare Function QueryServiceConfigW Lib "AdvApi32" ( handle as Integer, _
		    config as Ptr, size as Integer, ByRef bytesNeeded as Integer ) as Boolean
		    
		    dim mb as new MemoryBlock( 4096 )
		    dim bytesNeeded as Integer
		    dim err as Boolean
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "QueryServiceConfigW", "AdvApi32" )
		    
		    if unicodeSavvy then
		      err = QueryServiceConfigW( serv.Handle, mb, mb.Size, bytesNeeded )
		    else
		      err = QueryServiceConfigA( serv.Handle, mb, mb.Size, bytesNeeded )
		    end if
		    
		    if not err and bytesNeeded > 0 then
		      mb = new MemoryBlock( bytesNeeded )
		      if unicodeSavvy then
		        err = QueryServiceConfigW( serv.Handle, mb, mb.Size, bytesNeeded )
		      else
		        err = QueryServiceConfigA( serv.Handle, mb, mb.Size, bytesNeeded )
		      end if
		      
		      if not err then return
		    end
		    
		    dim path as String
		    dim params as String
		    dim num as Integer
		    try
		      serv.Type = mb.Long( 0 )
		      serv.StartType = mb.Long( 4 )
		      serv.ErrorControl = mb.Long( 8 )
		      // Get the path
		      if unicodeSavvy then
		        path = mb.Ptr( 12 ).WString( 0 )
		      else
		        path = mb.Ptr( 12 ).CString( 0 )
		      end if
		      
		      // Get the startup parameters and remove the "'s
		      num = CountFields( path, """" )
		      params = Trim( NthField( path, """", num ) )
		      path = mid( path, Len( params ) )
		      path = Trim( ReplaceAll( path, """", "" ) )
		      
		      serv.ExecutableFile = new FolderItem( path )
		      serv.StartupParameters = params
		      if unicodeSavvy then
		        serv.LoadOrderGroup = mb.Ptr( 16 ).WString( 0 )
		        serv.Dependencies = mb.Ptr( 24 ).WString( 0 )
		        serv.StartName = mb.Ptr( 28 ).WString( 0 )
		      else
		        serv.LoadOrderGroup = mb.Ptr( 16 ).CString( 0 )
		        serv.Dependencies = mb.Ptr( 24 ).CString( 0 )
		        serv.StartName = mb.Ptr( 28 ).CString( 0 )
		      end if
		      
		      if serv.StartName = "" then
		        serv.StartName = "LocalSystem"
		      end
		      
		      if unicodeSavvy then
		        serv.DisplayName = mb.Ptr( 32 ).WString( 0 )
		      else
		        serv.DisplayName = mb.Ptr( 32 ).CString( 0 )
		      end if
		    catch excerr as RuntimeException
		      
		    end
		    
		    Soft Declare Function QueryServiceConfig2A Lib "AdvApi32" ( handle as Integer, level as Integer, _
		    buffer as Ptr, size as Integer, ByRef bytesNeeded as Integer ) as Boolean
		    Soft Declare Function QueryServiceConfig2W Lib "AdvApi32" ( handle as Integer, level as Integer, _
		    buffer as Ptr, size as Integer, ByRef bytesNeeded as Integer ) as Boolean
		    
		    unicodeSavvy = System.IsFunctionAvailable( "QueryServiceConfig2W", "AdvApi32" )
		    
		    dim description as new MemoryBlock( 4 )
		    if unicodeSavvy then
		      err = QueryServiceConfig2W( serv.Handle, &h1, description, description.Size, bytesNeeded )
		    else
		      err = QueryServiceConfig2A( serv.Handle, &h1, description, description.Size, bytesNeeded )
		    end if
		    
		    if not err and bytesNeeded > 0 then
		      description = new MemoryBlock( bytesNeeded )
		      if unicodeSavvy then
		        err = QueryServiceConfig2W( serv.Handle, &h1, description, description.Size, bytesNeeded )
		      else
		        err = QueryServiceConfig2A( serv.Handle, &h1, description, description.Size, bytesNeeded )
		      end if
		      if not err then return
		    end
		    
		    try
		      if unicodeSavvy then
		        serv.Description = description.Ptr( 0 ).WString( 0 )
		      else
		        serv.Description = description.Ptr( 0 ).CString( 0 )
		      end if
		    catch excerr2 as RuntimeException
		      
		    end
		    
		  #else
		    
		    #pragma unused serv
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Finalize()
		  // Close the manager
		  CloseServiceHandle( mManager )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Initialize()
		  // Open the manager
		  mManager = OpenManager
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ModifyService(serv as ServiceWFS)
		  // You have to be sure that the service was opened with the ability to change
		  // the configuration.
		  #if TargetWin32
		    
		    Soft Declare Sub ChangeServiceConfigA Lib "AdvApi32" ( handle as Integer, type as Integer, _
		    startType as Integer, errorControl as Integer, path as CString, loadGroup as CString, _
		    tag as Integer, dependencies as CString, startName as CString, password as CString, _
		    displayName as CString )
		    Soft Declare Sub ChangeServiceConfigW Lib "AdvApi32" ( handle as Integer, type as Integer, _
		    startType as Integer, errorControl as Integer, path as WString, loadGroup as WString, _
		    tag as Integer, dependencies as WString, startName as WString, password as WString, _
		    displayName as WString )
		    
		    dim path as String
		    path = """" + serv.ExecutableFile.AbsolutePath + """"
		    
		    if serv.StartupParameters <> "" then
		      path = path + " " + serv.StartupParameters
		    end
		    
		    if System.IsFunctionAvailable( "ChangeServiceConfigW", "AdvApi32" ) then
		      ChangeServiceConfigW( serv.Handle, serv.Type, serv.StartType, serv.ErrorControl, _
		      path, serv.LoadOrderGroup, 0, serv.Dependencies, serv.StartName, serv.Password, _
		      serv.DisplayName )
		    else
		      ChangeServiceConfigA( serv.Handle, serv.Type, serv.StartType, serv.ErrorControl, _
		      path, serv.LoadOrderGroup, 0, serv.Dependencies, serv.StartName, serv.Password, _
		      serv.DisplayName )
		    end if
		    
		    Soft Declare Sub ChangeServiceConfig2A Lib "AdvApi32" ( handle as Integer, level as Integer, buffer as Ptr )
		    Soft Declare Sub ChangeServiceConfig2W Lib "AdvApi32" ( handle as Integer, level as Integer, buffer as Ptr )
		    
		    if System.IsFunctionAvailable( "ChangeServiceConfig2W", "AdvApi32" ) then
		      dim description as new MemoryBlock( Len( serv.Description ) * 2 + 2 )
		      dim descHandle as new MemoryBlock( 4 )
		      description.WString( 0 ) = serv.Description
		      descHandle.Ptr( 0 ) = description
		      ChangeServiceConfig2W( serv.Handle, &h1, descHandle )
		    else
		      dim description as new MemoryBlock( Len( serv.Description ) + 1 )
		      dim descHandle as new MemoryBlock( 4 )
		      description.CString( 0 ) = serv.Description
		      descHandle.Ptr( 0 ) = description
		      ChangeServiceConfig2A( serv.Handle, &h1, descHandle )
		    end if
		    
		  #else
		    
		    #pragma unused serv
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function OpenManager() As Integer
		  #if TargetWin32
		    Soft Declare Function OpenSCManagerA Lib "AdvApi32" ( computer as Integer, database as Integer, access as Integer ) as Integer
		    Soft Declare Function OpenSCManagerW Lib "AdvApi32" ( computer as Integer, database as Integer, access as Integer ) as Integer
		    
		    Const SC_MANAGER_CONNECT = &h1
		    Const SC_MANAGER_CREATE_SERVICE = &h2
		    Const SC_MANAGER_ENUMERATE_SERVICE = &h4
		    Const SC_MANAGER_LOCK = &h8
		    Const SC_MANAGER_QUERY_LOCK_STATUS = &h10
		    Const SC_MANAGER_MODIFY_BOOT_CONFIG = &h20
		    Const STANDARD_RIGHTS_REQUIRED = &h000F0000
		    Const SC_MANAGER_ALL_ACCESS = &hF003F
		    
		    if System.IsFunctionAvailable( "OpenSCManagerW", "AdvApi32" ) then
		      return OpenSCManagerW( 0, 0, SC_MANAGER_ALL_ACCESS )
		    else
		      return OpenSCManagerA( 0, 0, SC_MANAGER_ALL_ACCESS )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function OpenService(name as String, access as Integer) As ServiceWFS
		  dim serv as ServiceWFS
		  
		  #if TargetWin32
		    
		    Soft Declare Function OpenServiceA Lib "AdvApi32" ( manager as Integer, name as CString, _
		    access as Integer ) as Integer
		    Soft Declare Function OpenServiceW Lib "AdvApi32" ( manager as Integer, name as WString, _
		    access as Integer ) as Integer
		    
		    if mManager = 0 then return nil
		    
		    serv = new Service
		    if System.IsFunctionAvailable( "OpenServiceW", "AdvApi32" ) then
		      serv.Handle = OpenServiceW( mManager, name, access )
		    else
		      serv.Handle = OpenServiceA( mManager, name, access )
		    end if
		    
		    if serv.Handle = 0 then return nil
		    
		    if Bitwise.BitAnd( access, kAccessServiceQueryConfig ) <> 0 then
		      FillServiceInformation( serv )
		    end
		    
		  #else
		    
		    #pragma unused name
		    #pragma unused access
		    
		  #endif
		  
		  return serv
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RemoveService(s as ServiceWFS)
		  #if TargetWin32
		    
		    // We want to delete the service, so this assumes that we
		    // have already opened the service.  If we don't have a Handle
		    // then we will try to use the Name
		    
		    Declare Sub DeleteService Lib "AdvApi32" ( handle as Integer )
		    
		    dim handle as Integer
		    dim closeTheHandle as Boolean
		    if s.Handle > 0 then
		      // We can just delete as normal
		      handle = s.Handle
		    else
		      // We don't have a valid handle, so open up the
		      // service by name
		      handle = OpenService( s.Name, &hF0000 ).Handle
		      
		      // And be sure to close the handle when we're done
		      closeTheHandle = true
		    end
		    
		    if handle > 0 then
		      DeleteService( handle )
		    end
		    
		    if closeTheHandle then
		      CloseServiceHandle( handle )
		    end
		    
		  #else
		    
		    #pragma unused s
		    
		  #endif
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mManager As Integer
	#tag EndProperty


	#tag Constant, Name = kAccessServiceAll, Type = Integer, Dynamic = False, Default = \"&h000F01FF", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceChangeConfig, Type = Integer, Dynamic = False, Default = \"&h2", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceEnumerateDependents, Type = Integer, Dynamic = False, Default = \"&h8", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceInterrogate, Type = Integer, Dynamic = False, Default = \"&h80", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServicePauseContinue, Type = Integer, Dynamic = False, Default = \"&h40", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceQueryConfig, Type = Integer, Dynamic = False, Default = \"&h1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceQueryStatus, Type = Integer, Dynamic = False, Default = \"&h4", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceStart, Type = Integer, Dynamic = False, Default = \"&h10", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceStop, Type = Integer, Dynamic = False, Default = \"&h20", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kAccessServiceUserDefinedControl, Type = Integer, Dynamic = False, Default = \"&h100", Scope = Protected
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
