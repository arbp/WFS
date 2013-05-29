#tag Module
Protected Module PlugAndPlay
	#tag Method, Flags = &h21
		Private Sub CreateInvisibleWindow()
		  #if TargetWin32
		    Soft Declare Sub RegisterClassA Lib "User32" ( wc as Ptr )
		    Soft Declare Sub RegisterClassW Lib "User32" ( wc as Ptr )
		    Soft Declare Function CreateWindowExA Lib "User32" ( zero as Integer, className as CString, _
		    title as CString, style as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, parent as Integer, menu as Integer, hInst as Integer, lParam as Integer ) as Integer
		    Soft Declare Function CreateWindowExW Lib "User32" ( zero as Integer, className as WString, _
		    title as WString, style as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, parent as Integer, menu as Integer, hInst as Integer, lParam as Integer ) as Integer
		    Declare Function GetModuleHandleA Lib "Kernel32" ( null as Integer ) as Integer
		    
		    dim unicodeSavvy as Boolean = System.IsFunctionAvailable( "RegisterClassW", "User32" )
		    mUnicodeSavvy = unicodeSavvy
		    
		    // Get a handle to our application instance
		    dim hInstance as Integer
		    hInstance = GetModuleHandleA( 0 )
		    
		    // And make the class to register
		    dim wc as new MemoryBlock( 40 )
		    dim className as new MemoryBlock( 1024 )
		    if unicodeSavvy then
		      className.WString( 0 ) = "InvisibleWnd"
		    else
		      className.CString( 0 ) = "InvisibleWnd"
		    end if
		    wc.Ptr( 4 ) = AddressOf PlugAndPlayWndProc
		    wc.Long( 16 ) = hInstance
		    wc.Ptr( 36 ) = className
		    
		    if unicodeSavvy then
		      RegisterClassW( wc )
		      mWnd = CreateWindowExW( 0, "InvisibleWnd", "", 0, 0, 0, 0, 0, 0, 0, hInstance, 0 )
		    else
		      RegisterClassA( wc )
		      mWnd = CreateWindowExA( 0, "InvisibleWnd", "", 0, 0, 0, 0, 0, 0, 0, hInstance, 0 )
		    end if
		    
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetDeviceTypeString(deviceType as Integer) As String
		  Const DBT_DEVTYP_OEM                        = &h00000000  // oem-defined device type
		  Const DBT_DEVTYP_DEVNODE             = &h00000001  // devnode number
		  Const DBT_DEVTYP_VOLUME                = &h00000002  // logical volume
		  Const DBT_DEVTYP_PORT                     = &h00000003  // serial, parallel
		  Const DBT_DEVTYP_NET                        = &h00000004  // network resource
		  Const DBT_DEVTYP_DEVICEINTERFACE = &h00000005  // device interface class
		  Const DBT_DEVTYP_HANDLE               = &h00000006  // file system handle
		  
		  dim deviceTypeStr as String
		  
		  select case deviceType
		  case DBT_DEVTYP_OEM
		    deviceTypeStr = kDeviceTypeOEM
		  case DBT_DEVTYP_DEVNODE
		    deviceTypeStr = kDeviceTypeNode
		  case DBT_DEVTYP_VOLUME
		    deviceTypeStr = kDeviceTypeVolume
		  case DBT_DEVTYP_PORT
		    deviceTypeStr = kDeviceTypePort
		  case DBT_DEVTYP_NET
		    deviceTypeStr = kDeviceTypeNetworkResource
		  case DBT_DEVTYP_DEVICEINTERFACE
		    deviceTypeStr = kDeviceTypeInterface
		  case DBT_DEVTYP_HANDLE
		    deviceTypeStr = kDeviceTypeFileSystem
		  else
		    deviceTypeStr = kDeviceTypeUnknown
		  end
		  
		  return deviceTypeStr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetDrivesFromMask(driveBitmask as Integer) As String()
		  // The bitmask goes like this.  Bit 1 = A, bit 2 = B, and
		  // so on.
		  dim i as Integer
		  dim driveLetters(-1) as String
		  for i = 0 to 25
		    if BitwiseAnd( driveBitmask, Bitwise.ShiftLeft( 1, i ) ) <> 0 then
		      // This bit is set, so let's add the drive letter
		      driveLetters.Append( Chr( 65 + i ) + ":\"  )
		    end
		  next
		  
		  return driveLetters
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GuidToString(Data1 as Integer, Data2 as Integer, Data3 as Integer, Data4() as Integer) As String
		  dim guid_string as String
		  guid_string = "{" + Pad( Hex( Data1 ), 8 ) + "-" + Pad( Hex( Data2 ), 4 ) + "-" +_
		  Pad( Hex( Data3 ), 4 ) + "-" + Pad( Hex( Data4( 0 ) ), 2 ) + Pad( Hex( Data4( 1 ) ), 2 ) + "-" + _
		  Pad( Hex( Data4( 2 ) ), 2 ) + Pad( Hex( Data4( 3 ) ), 2 ) + Pad( Hex( Data4( 4 ) ), 2 ) + _
		  Pad( Hex( Data4( 5 ) ), 2 ) + Pad( Hex( Data4( 6 ) ), 2 ) + Pad( Hex( Data4( 7 ) ), 2 ) + "}"
		  
		  return guid_string
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleDeviceAdded(lParam as Integer)
		  // A device was just added, so let's gather
		  // information about it and alert the user
		  
		  // First, get the device header
		  dim dev_broadcast_header as MemoryBlock
		  dev_broadcast_header = LParamToMemoryBlock( lParam )
		  
		  // Now, we can figure out what the device type is
		  // from the header
		  dim deviceType as Integer
		  deviceType = dev_broadcast_header.Long( 4 )
		  
		  dim deviceTypeStr as String
		  deviceTypeStr = GetDeviceTypeString( deviceType )
		  
		  // Now that we know what device was added, let's
		  // let the observers know.  First, let the specific
		  // observers know.  We can pass in the broadcast
		  // header because it's a generic structure that
		  // each specific type has extra data bolted on to
		  select case deviceTypeStr
		  case kDeviceTypePort
		    HandlePortAdded( dev_broadcast_header )
		  case kDeviceTypeVolume
		    HandleVolumeAdded( dev_broadcast_header )
		  case kDeviceTypeInterface
		    HandleDeviceInterfaceAdded( dev_broadcast_header )
		  case kDeviceTypeOEM
		    HandleOEMAdded( dev_broadcast_header )
		  end
		  
		  // And then, let the generic observers have a
		  // stab at it
		  dim observer as GenericDeviceStateObserver
		  for each observer in mObservers
		    observer.DeviceAdded( deviceTypeStr )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleDeviceInterfaceAdded(dev_broadcast_deviceinterface as MemoryBlock)
		  // A device interface was added to the system, so
		  // we should report it.
		  
		  // A GUID has the following structure:
		  //typedef struct _GUID {
		  //   ULONG   Data1;
		  //   unsigned short Data2;
		  //   unsigned short Data3;
		  //   unsigned char Data4[8];
		  //} GUID;
		  
		  // First, read in the GUID
		  dim i, Data1, Data2, Data3 as Integer
		  Data1 = dev_broadcast_deviceinterface.Long( 12 )
		  Data2 = dev_broadcast_deviceinterface.UShort( 16 )
		  Data3 = dev_broadcast_deviceinterface.UShort( 18 )
		  dim Data4(8) as Integer
		  for i = 0 to 7
		    Data4( i ) = dev_broadcast_deviceinterface.Byte( 20 + i )
		  next
		  
		  // Now, define it as a string for the user
		  dim guid_string as String
		  guid_string = GuidToString( Data1, Data2, Data3, Data4 )
		  
		  // And grab the friendly name of the interface
		  dim friendlyName as String
		  if mUnicodeSavvy then
		    friendlyName = dev_broadcast_deviceinterface.WString( 28 )
		  else
		    friendlyName = dev_broadcast_deviceinterface.CString( 28 )
		  end if
		  
		  // And alert the observers
		  dim observer as DeviceInterfaceStateObserver
		  for each observer in mDeviceInterfaceObservers
		    observer.DeviceInterfaceAdded( guid_string, friendlyName )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleDeviceInterfaceRemoved(dev_broadcast_deviceinterface as MemoryBlock)
		  // A device interface was removed from the system, so
		  // we should report it.
		  
		  // A GUID has the following structure:
		  //typedef struct _GUID {
		  //   ULONG   Data1;
		  //   unsigned short Data2;
		  //   unsigned short Data3;
		  //   unsigned char Data4[8];
		  //} GUID;
		  
		  // First, read in the GUID
		  dim i, Data1, Data2, Data3 as Integer
		  Data1 = dev_broadcast_deviceinterface.Long( 12 )
		  Data2 = dev_broadcast_deviceinterface.UShort( 16 )
		  Data3 = dev_broadcast_deviceinterface.UShort( 18 )
		  dim Data4(8) as Integer
		  for i = 0 to 7
		    Data4( i ) = dev_broadcast_deviceinterface.Byte( 20 + i )
		  next
		  
		  // Now, define it as a string for the user
		  dim guid_string as String
		  guid_string = GuidToString( Data1, Data2, Data3, Data4 )
		  
		  // And grab the friendly name of the interface
		  dim friendlyName as String
		  if mUnicodeSavvy then
		    friendlyName = dev_broadcast_deviceinterface.WString( 28 )
		  else
		    friendlyName = dev_broadcast_deviceinterface.CString( 28 )
		  end if
		  
		  // And alert the observers
		  dim observer as DeviceInterfaceStateObserver
		  for each observer in mDeviceInterfaceObservers
		    observer.DeviceInterfaceRemoved( guid_string, friendlyName )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleDeviceQueryRemove(lParam as Integer) As Boolean
		  // The system wants to know whether it's OK to remove
		  // a device or not
		  
		  // First, get the device header
		  dim dev_broadcast_header as MemoryBlock
		  dev_broadcast_header = LParamToMemoryBlock( lParam )
		  
		  // Now, we can figure out what the device type is
		  // from the header
		  dim deviceType as Integer
		  deviceType = dev_broadcast_header.Long( 4 )
		  
		  dim deviceTypeStr as String
		  deviceTypeStr = GetDeviceTypeString( deviceType )
		  
		  // Now that we know what device is being removed, let's
		  // let the observers know.  First, let the specific
		  // observers know.  We can pass in the broadcast
		  // header because it's a generic structure that
		  // each specific type has extra data bolted on to
		  dim ret as Boolean
		  select case deviceTypeStr
		  case kDeviceTypePort
		    ret = HandleQueryPortRemoved( dev_broadcast_header )
		  case kDeviceTypeVolume
		    ret = HandleQueryVolumeRemoved( dev_broadcast_header )
		  case kDeviceTypeInterface
		    ret = HandleQueryDeviceInterfaceRemoved( dev_broadcast_header )
		  case kDeviceTypeOEM
		    ret = HandleQueryOemRemoved( dev_broadcast_header )
		  end
		  
		  // One of the specific handlers handled the function
		  if ret then return true
		  
		  // Now that we know what device is, let's
		  // ask the observers if it's ok to remove the device
		  dim observer as GenericDeviceStateObserver
		  for each observer in mObservers
		    ret = observer.CancelDeviceRemove( deviceTypeStr )
		    if ret = true then return true
		  next
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleDeviceRemoved(lParam as Integer)
		  // A device was just removed, so let's gather
		  // information about it and alert the user
		  
		  // First, get the device header
		  dim dev_broadcast_header as MemoryBlock
		  dev_broadcast_header = LParamToMemoryBlock( lParam )
		  
		  // Now, we can figure out what the device type is
		  // from the header
		  dim deviceType as Integer
		  deviceType = dev_broadcast_header.Long( 4 )
		  
		  dim deviceTypeStr as String
		  deviceTypeStr = GetDeviceTypeString( deviceType )
		  
		  // Now that we know what device was removed, let's
		  // let the observers know.  First, let the specific
		  // observers know.  We can pass in the broadcast
		  // header because it's a generic structure that
		  // each specific type has extra data bolted on to
		  select case deviceTypeStr
		  case kDeviceTypePort
		    HandlePortRemoved( dev_broadcast_header )
		  case kDeviceTypeVolume
		    HandleVolumeRemoved( dev_broadcast_header )
		  case kDeviceTypeInterface
		    HandleDeviceInterfaceRemoved( dev_broadcast_header )
		  case kDeviceTypeOEM
		    HandleOEMRemoved( dev_broadcast_header )
		  end
		  
		  // Now let the generic observers know
		  dim observer as GenericDeviceStateObserver
		  for each observer in mObservers
		    observer.DeviceRemoved( deviceTypeStr )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleOEMAdded(dev_broadcast_oem as MemoryBlock)
		  // An OEM device was added, so let's gather
		  // the information from it
		  
		  // First, get the id
		  dim id as Integer
		  id = dev_broadcast_oem.Long( 12 )
		  
		  // Then get the device-specific support
		  // function
		  dim support as Integer
		  support = dev_broadcast_oem.Long( 16 )
		  
		  // And call the observers
		  dim observer as OemStateObserver
		  for each observer in mOemObservers
		    observer.OemAdded( id, support )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleOemRemoved(dev_broadcast_oem as MemoryBlock)
		  // An OEM device was removed, so let's gather
		  // the information from it
		  
		  // First, get the id
		  dim id as Integer
		  id = dev_broadcast_oem.Long( 12 )
		  
		  // Then get the device-specific support
		  // function
		  dim support as Integer
		  support = dev_broadcast_oem.Long( 16 )
		  
		  // And call the observers
		  dim observer as OemStateObserver
		  for each observer in mOemObservers
		    observer.OemRemoved( id, support )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandlePortAdded(dev_broadcast_port as MemoryBlock)
		  // A port was added.  The only extra information we have is
		  // about the friendly name of the port
		  dim friendlyName as String
		  if mUnicodeSavvy then
		    friendlyName = dev_broadcast_port.WString( 12 )
		  else
		    friendlyName = dev_broadcast_port.CString( 12 )
		  end if
		  
		  // Let the port observers know about it
		  dim observer as PortStateObserver
		  for each observer in mPortObservers
		    observer.PortAdded( friendlyName )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandlePortRemoved(dev_broadcast_port as MemoryBlock)
		  // A port was removed.  The only extra information we have is
		  // about the friendly name of the port
		  dim friendlyName as String
		  if mUnicodeSavvy then
		    friendlyName = dev_broadcast_port.WString( 12 )
		  else
		    friendlyName = dev_broadcast_port.CString( 12 )
		  end if
		  
		  // Let the port observers know about it
		  dim observer as PortStateObserver
		  for each observer in mPortObservers
		    observer.PortRemoved( friendlyName )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleQueryDeviceInterfaceRemoved(dev_broadcast_deviceinterface as MemoryBlock) As Boolean
		  // A device interface is about to be removed from the system, so
		  // we should report it.
		  
		  // A GUID has the following structure:
		  //typedef struct _GUID {
		  //   ULONG   Data1;
		  //   unsigned short Data2;
		  //   unsigned short Data3;
		  //   unsigned char Data4[8];
		  //} GUID;
		  
		  // First, read in the GUID
		  dim i, Data1, Data2, Data3 as Integer
		  Data1 = dev_broadcast_deviceinterface.Long( 12 )
		  Data2 = dev_broadcast_deviceinterface.UShort( 16 )
		  Data3 = dev_broadcast_deviceinterface.UShort( 18 )
		  dim Data4(8) as Integer
		  for i = 0 to 7
		    Data4( i ) = dev_broadcast_deviceinterface.Byte( 20 + i )
		  next
		  
		  // Now, define it as a string for the user
		  dim guid_string as String
		  guid_string = GuidToString( Data1, Data2, Data3, Data4 )
		  
		  // And grab the friendly name of the interface
		  dim friendlyName as String
		  friendlyName = dev_broadcast_deviceinterface.CString( 28 )
		  
		  // And alert the observers
		  dim observer as DeviceInterfaceStateObserver
		  dim ret as Boolean
		  for each observer in mDeviceInterfaceObservers
		    ret = observer.CancelDeviceInterfaceRemove( guid_string, friendlyName )
		    if ret then return true
		  next
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleQueryOemRemoved(dev_broadcast_oem as MemoryBlock) As Boolean
		  // The system wants to know if it's ok to
		  // remove this OEM device, so let's find out
		  
		  // First, get the id
		  dim id as Integer
		  id = dev_broadcast_oem.Long( 12 )
		  
		  // Then get the device-specific support
		  // function
		  dim support as Integer
		  support = dev_broadcast_oem.Long( 16 )
		  
		  // And call the observers
		  dim ret as Boolean
		  dim observer as OemStateObserver
		  for each observer in mOemObservers
		    ret = observer.CancelOemRemove( id, support )
		    if ret then return true
		  next
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleQueryPortRemoved(dev_broadcast_port as MemoryBlock) As Boolean
		  // A port is being removed.  The only extra information we have is
		  // about the friendly name of the port
		  dim friendlyName as String
		  if mUnicodeSavvy then
		    friendlyName = dev_broadcast_port.WString( 12 )
		  else
		    friendlyName = dev_broadcast_port.CString( 12 )
		  end if
		  
		  // Let the port observers know about it
		  dim observer as PortStateObserver
		  dim ret as Boolean
		  for each observer in mPortObservers
		    ret = observer.CancelPortRemove( friendlyName )
		    if ret then return true
		  next
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleQueryVolumeRemoved(dev_broadcast_volume as MemoryBlock) As Boolean
		  // The system wants to know whether it's ok
		  // to remove this volume
		  
		  Const DBTF_MEDIA = &h0001          // media comings and goings
		  Const DBTF_NET = &h0002          // network volume
		  
		  // First, get the drive bitmask
		  dim driveBitmask as Integer
		  driveBitmask = dev_broadcast_volume.Long( 12 )
		  dim driveLetters(-1) as String
		  driveLetters = GetDrivesFromMask( driveBitmask )
		  
		  // And then get the flags
		  dim flags as Integer
		  flags = dev_broadcast_volume.Short( 16 )
		  
		  // Now, figure out what the flags mean
		  dim isMedia as Boolean
		  isMedia = BitwiseAnd( flags, DBTF_MEDIA ) <> 0
		  dim isNetwork as Boolean
		  isNetwork = BitwiseAnd( flags, DBTF_NET ) <> 0
		  
		  // Let the port observers know about it
		  dim observer as VolumeStateObserver
		  dim ret as Boolean
		  for each observer in mVolumeObservers
		    ret = observer.CancelVolumeRemove( driveLetters, isMedia, isNetwork )
		    if ret then return true
		  next
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleVolumeAdded(dev_broadcast_volume as MemoryBlock)
		  // A volume has been added.  There are only
		  // two pieces of information to grab. What logical
		  // drives are affected, and whether it's media or network
		  
		  Const DBTF_MEDIA = &h0001          // media comings and goings
		  Const DBTF_NET = &h0002          // network volume
		  
		  // First, get the drive bitmask
		  dim driveBitmask as Integer
		  driveBitmask = dev_broadcast_volume.Long( 12 )
		  dim driveLetters(-1) as String
		  driveLetters = GetDrivesFromMask( driveBitmask )
		  
		  // And then get the flags
		  dim flags as Integer
		  flags = dev_broadcast_volume.Short( 16 )
		  
		  // Now, figure out what the flags mean
		  dim isMedia as Boolean
		  isMedia = BitwiseAnd( flags, DBTF_MEDIA ) <> 0
		  dim isNetwork as Boolean
		  isNetwork = BitwiseAnd( flags, DBTF_NET ) <> 0
		  
		  // Now, after all that, we can call the volume
		  // observers to let them know this was added
		  dim observer as VolumeStateObserver
		  for each observer in mVolumeObservers
		    observer.VolumeAdded( driveLetters, isMedia, isNetwork )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleVolumeRemoved(dev_broadcast_volume as MemoryBlock)
		  // A volume has been removed.  There are only
		  // two pieces of information to grab. What logical
		  // drives are affected, and whether it's media or network
		  
		  Const DBTF_MEDIA = &h0001          // media comings and goings
		  Const DBTF_NET = &h0002          // network volume
		  
		  // First, get the drive bitmask
		  dim driveBitmask as Integer
		  driveBitmask = dev_broadcast_volume.Long( 12 )
		  dim driveLetters(-1) as String
		  driveLetters = GetDrivesFromMask( driveBitmask )
		  
		  // And then get the flags
		  dim flags as Integer
		  flags = dev_broadcast_volume.Short( 16 )
		  
		  // Now, figure out what the flags mean
		  dim isMedia as Boolean
		  isMedia = BitwiseAnd( flags, DBTF_MEDIA ) <> 0
		  dim isNetwork as Boolean
		  isNetwork = BitwiseAnd( flags, DBTF_NET ) <> 0
		  
		  // Now, after all that, we can call the volume
		  // observers to let them know this was removed
		  dim observer as VolumeStateObserver
		  for each observer in mVolumeObservers
		    observer.VolumeRemoved( driveLetters, isMedia, isNetwork )
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function LParamToMemoryBlock(lParam as Integer) As MemoryBlock
		  // Make a new memory block to hold the pointer
		  dim block as new MemoryBlock( 4 )
		  // Set the pointer in the memory block
		  block.Long( 0 ) = lParam
		  // Retrieve (and return) the pointer we just stuffed
		  // in there as a MemoryBlock
		  return block.Ptr( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Pad(s as String, numDigits as Integer) As String
		  dim ret as String
		  dim start, numZeros, i as Integer
		  
		  numZeros = numDigits - Len( s )
		  for i = 0 to numZeros - 1
		    ret = ret + "0"
		  next
		  ret = ret + s
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PlugAndPlayWndProc(hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  #if TargetWin32
		    Const WM_DEVICECHANGE = &h0219
		    
		    // We will get a msg of WM_DEVICECHANGE
		    if msg <> WM_DEVICECHANGE then goto done
		    
		    // And it's wParam will have specify the event
		    Const DBT_DEVICEARRIVAL = &h8000
		    Const DBT_DEVICEQUERYREMOVE = &h8001
		    Const DBT_DEVICEREMOVECOMPLETE = &h8004
		    Const BROADCAST_QUERY_DENY = &h424D5144
		    
		    if wParam = DBT_DEVICEARRIVAL then
		      // A device has been added, so let's extract what type of
		      // device this really is
		      HandleDeviceAdded( lParam )
		    elseif wParam = DBT_DEVICEREMOVECOMPLETE then
		      // A device has been removed, so handle it
		      HandleDeviceRemoved( lParam )
		    elseif wParam = DBT_DEVICEQUERYREMOVE then
		      // The system wants to know whether we wish to
		      // allow the system to remove the device
		      if HandleDeviceQueryRemove( lParam ) then
		        // The user doesn't want this to happen, so
		        // we eat the message
		        return BROADCAST_QUERY_DENY
		      end
		    end
		    
		    done:
		    Soft Declare Function DefWindowProcA Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function DefWindowProcW Lib "User32" ( hwnd as Integer, msg as Integer, wParam as Integer, lParam as Integer ) as Integer
		    
		    if mUnicodeSavvy then
		      return DefWindowProcW( hwnd, msg, wParam, lParam )
		    else
		      return DefWindowProcA( hwnd, msg, wParam, lParam )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WatchForDeviceInterfaceEvents(observer as DeviceInterfaceStateObserver)
		  mDeviceInterfaceObservers.Append( observer )
		  
		  if mWnd = 0 then
		    CreateInvisibleWindow
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WatchForGenericDeviceEvents(observer as GenericDeviceStateObserver)
		  mObservers.Append( observer )
		  
		  if mWnd = 0 then
		    CreateInvisibleWindow
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WatchForOemDeviceEvents(observer as OemStateObserver)
		  mOemObservers.Append( observer )
		  
		  if mWnd = 0 then
		    CreateInvisibleWindow
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WatchForPortDeviceEvents(observer as PortStateObserver)
		  mPortObservers.Append( observer )
		  
		  if mWnd = 0 then
		    CreateInvisibleWindow
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WatchForVolumeDeviceEvents(observer as VolumeStateObserver)
		  mVolumeObservers.Append( observer )
		  
		  if mWnd = 0 then
		    CreateInvisibleWindow
		  end
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mDeviceInterfaceObservers(-1) As DeviceInterfaceStateObserver
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mObservers(-1) As GenericDeviceStateObserver
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mOemObservers(-1) As OemStateObserver
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPortObservers(-1) As PortStateObserver
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mUnicodeSavvy As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVolumeObservers(-1) As VolumeStateObserver
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWnd As Integer
	#tag EndProperty


	#tag Constant, Name = kDeviceTypeFileSystem, Type = String, Dynamic = False, Default = \"File System Handle", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypeInterface, Type = String, Dynamic = False, Default = \"Device Interface Class", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypeNetworkResource, Type = String, Dynamic = False, Default = \"Network Resource", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypeNode, Type = String, Dynamic = False, Default = \"Device Node", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypeOEM, Type = String, Dynamic = False, Default = \"OEM", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypePort, Type = String, Dynamic = False, Default = \"Port", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypeUnknown, Type = String, Dynamic = False, Default = \"Unknown", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kDeviceTypeVolume, Type = String, Dynamic = False, Default = \"Volume", Scope = Protected
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
