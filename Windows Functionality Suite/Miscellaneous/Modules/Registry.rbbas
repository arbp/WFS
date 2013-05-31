#tag Module
Protected Module Registry
	#tag Method, Flags = &h1
		Protected Function AllowActiveDesktop() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoActiveDesktopChanges" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowActiveDesktop(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoActiveDesktopChanges" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowChangeStartMenu() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoChangeStartMenu" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowChangeStartMenu(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoChangeStartMenu" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowCloseLogoff() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoLogoff" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowCloseLogoff(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoLogoff" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowControlPanels() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "System" )
		    
		    return (policies.Value( "NoDispCpl" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowControlPanels(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    try
		      policies = policies.Child( "System" )
		    catch
		      policies = policies.AddFolder( "System" )
		    end
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoDispCpl" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowDesktop() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoDesktop" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowDesktop(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoDesktop" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowDrive(driveNum as Integer) As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim oldMask as Integer
		    
		    try
		      oldMask = policies.Value( "NoDrives" )
		    end
		    
		    return BitwiseAnd( oldMask, Pow( 2, driveNum ) ) = 0
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowDrive(driveNum as Integer, assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim oldMask as Integer
		    
		    try
		      oldMask = policies.Value( "NoDrives" )
		    end
		    
		    if not set then
		      oldMask = BitwiseOr( oldMask, Pow( 2, driveNum ) )
		    else
		      oldMask = BitwiseAnd( oldMask, Bitwise.OnesComplement( Pow( 2, driveNum ) ) )
		    end
		    
		    policies.Value( "NoDrives" ) = oldMask
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowFavoritesMenu() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoFavoritesMenu" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowFavoritesMenu(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoFavoritesMenu" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowFileAssociate() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoFileAssociate" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowFileAssociate(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoFileAssociate" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowFind() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoFind" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowFind(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoFind" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowFolderOptions() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoFolderOptions" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowFolderOptions(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoFolderOptions" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowInternetIcon() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoInternetIcon" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowInternetIcon(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoInternetIcon" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowRecentDocumentsMenu() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoRecentDocsMenu" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowRecentDocumentsMenu(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoRecentDocsMenu" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowRegistryTools() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "System" )
		    
		    return (policies.Value( "DisableRegistryTools" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowRegistryTools(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    try
		      policies = policies.Child( "System" )
		    catch
		      policies = policies.AddFolder( "System" )
		    end
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "DisableRegistryTools" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowRunMenu() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoRun" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowRunMenu(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoRun" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowSpecialFolders() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoSetFolders" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowSpecialFolders(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoSetFolders" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowTaskbarBands() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoMovingBands" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowTaskbarBands(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoMovingBands" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowTaskbarProperties() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoSetTaskbar" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowTaskbarProperties(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoSetTaskbar" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowTaskManager() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "System" )
		    
		    return (policies.Value( "DisableTaskMgr" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowTaskManager(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    try
		      policies = policies.Child( "System" )
		    catch
		      policies = policies.AddFolder( "System" )
		    end
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "DisableTaskMgr" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function AllowTrayContextMenu() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return true
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "NoTrayContextMenu" ) = 0)
		  #endif
		  
		exception
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub AllowTrayContextMenu(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if not set then setInt = 1
		    
		    policies.Value( "NoTrayContextMenu" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ClearRecentDocsOnExit() As Boolean
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return false
		    
		    policies = policies.Child( "Explorer" )
		    
		    return (policies.Value( "ClearRecentDocsOnExit" ) = 1)
		  #endif
		  
		exception
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ClearRecentDocsOnExit(assigns set as Boolean)
		  #if TargetWin32
		    dim policies as RegistryItem
		    
		    policies = CurrentUsersPolicies
		    
		    if policies = nil then return
		    
		    policies = policies.Child( "Explorer" )
		    
		    dim setInt as Integer
		    if set then setInt = 1
		    
		    policies.Value( "ClearRecentDocsOnExit" ) = setInt
		  #endif
		  
		exception
		  return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CurrentUsersPolicies() As RegistryItem
		  #if TargetWin32
		    dim ret as RegistryItem
		    
		    try
		      ret = new RegistryItem( "HKEY_CURRENT_USER" )
		      ret = ret.Child( "Software" ).Child( "Microsoft" )
		      ret = ret.Child( "Windows" ).Child( "CurrentVersion" )
		      ret = ret.Child( "Policies" )
		    catch
		      ret = nil
		    end
		    
		    return ret
		  #endif
		  
		  return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetBIOSDate() As String
		  #if TargetWin32 then
		    
		    Dim sysInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System" )
		    
		    if reg <> Nil then
		      sysInfo = reg.Value( "SystemBiosDate" )
		    end if
		    
		    Return sysInfo
		    
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetBIOSVersion() As String
		  #if TargetWin32 then
		    Dim sysInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System" )
		    
		    if reg <> Nil then
		      sysInfo = reg.Value( "SystemBiosVersion" )
		    end if
		    
		    Return sysInfo
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCPUIdentifier(ProcessorNumber as Integer) As String
		  #if TargetWin32 then
		    Dim cpuInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\" + _
		    CStr( ProcessorNumber ) )
		    
		    if reg <> Nil then
		      cpuInfo = reg.Value( "Identifier" )
		    end if
		    
		    Return cpuInfo
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCPUName(ProcessorNumber as Integer) As String
		  #if TargetWin32 then
		    Dim cpuInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\" + _
		    CStr( ProcessorNumber ) )
		    
		    if reg <> Nil then
		      cpuInfo = reg.Value( "ProcessorNameString" )
		    end if
		    
		    Return cpuInfo
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCPUSpeed(ProcessorNumber as Integer) As String
		  #if TargetWin32 then
		    Dim cpuSpeed As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\" + _
		    CStr( ProcessorNumber ) )
		    
		    dim mb as new MemoryBlock( 4 )
		    if reg <> Nil then
		      cpuSpeed = reg.Value( "~MHZ" )
		    end if
		    
		    Return cpuSpeed
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCPUType() As String
		  #if TargetWin32
		    Declare Sub GetSystemInfo Lib "Kernel32" ( info as Ptr )
		    
		    dim info as new MemoryBlock( 9 * 4 )
		    GetSystemInfo( info )
		    
		    select case info.Long( 24 )
		      
		    case 386
		      
		      Return "Intel 386"
		      
		    case 486
		      
		      Return "Intel 486"
		      
		    case 586
		      
		      Return "Intel Pentium"
		      
		    case 4000
		      
		      Return "MIPS 4000"
		      
		    case 21064
		      
		      Return "Alpha"
		      
		    case 601
		      
		      Return "PPC 601"
		      
		    case 603
		      
		      Return "PPC 603"
		      
		    case 604
		      
		      Return "PPC 604"
		      
		    case 620
		      
		      Return "PPC 620"
		      
		    case 10003
		      
		      Return "Hitachi SH3"
		      
		    case 10004
		      
		      Return "Hitachi SH3E"
		      
		    case 10005
		      
		      Return "Hitachi SH4"
		      
		    Case 821
		      
		      Return "Motorola 821"
		      
		    case 103
		      
		      Return "SHx SH3"
		      
		    case 104
		      
		      Return "SHx SH4"
		      
		    case 2577
		      
		      Return "STRONGARM"
		      
		    case 1824
		      
		      Return "ARM 720"
		      
		    case 2080
		      
		      Return "ARM 820"
		      
		    case 2336
		      
		      Return "ARM 920"
		      
		    case 70001
		      
		      Return "ARM 7TDMI"
		      
		    end select
		    
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetCPUVendor(ProcessorNumber as Integer) As String
		  #if TargetWin32 then
		    Dim cpuInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\" + _
		    CStr( ProcessorNumber ) )
		    
		    if reg <> Nil then
		      cpuInfo = reg.Value( "VendorIdentifier" )
		    end if
		    
		    Return cpuInfo
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetVideoBIOSDate() As String
		  #if TargetWin32 then
		    
		    Dim sysInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System" )
		    
		    if reg <> Nil then
		      sysInfo = reg.Value( "VideoBiosDate" )
		    end if
		    
		    Return sysInfo
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetVideoBIOSVersion() As String
		  #if TargetWin32 then
		    
		    Dim sysInfo As String
		    
		    dim reg as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System" )
		    
		    if reg <> Nil then
		      sysInfo = reg.Value( "VideoBiosVersion" )
		    end if
		    
		    Return sysInfo
		  #endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function WallpaperStyle() As Integer
		  Dim dskControlPanel as RegistryItem
		  dskControlPanel = new RegistryItem( "HKEY_CURRENT_USERS" )
		  dskControlPanel = dskControlPanel.Child( "Control Panel" ).Child( "Desktop" )
		  
		  if dskControlPanel = nil then return -1
		  
		  dim wallStyle as String = dskControlPanel.Value( "WallpaperStyle" ).StringValue
		  dim tileValue as String = dskControlPanel.Value( "TileWallpaper" ).StringValue
		  
		  if tileValue = "1" and wallStyle = "1" then
		    return kTiled
		  elseif tileValue = "0" and wallStyle = "2" then
		    return kStretched
		  else
		    return kCentered
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WallpaperStyle(assigns newStyle as Integer)
		  Dim DesktopControlPanel as RegistryItem
		  DesktopControlPanel = new RegistryItem( "HKEY_CURRENT_USERS" )
		  DesktopControlPanel = DesktopControlPanel.Child( "Control Panel" ).Child( "Desktop" )
		  
		  if DesktopControlPanel = nil then return
		  
		  dim wallStyle, tileValue as String
		  
		  select case newStyle
		  case kTiled
		    wallStyle = "1"
		    tileValue = "1"
		    
		  case kStretched
		    wallStyle = "2"
		    tileValue = "0"
		    
		  case kCentered
		    wallStyle = "0"
		    tileValue = "0"
		    
		  end select
		  
		  DesktopControlPanel.Value( "WallpaperStyle" ) = wallStyle
		  DesktopControlPanel.Value( "TileWallpaper" ) = tileValue
		End Sub
	#tag EndMethod


	#tag Constant, Name = kCentered, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kStretched, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTiled, Type = Double, Dynamic = False, Default = \"0", Scope = Public
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
