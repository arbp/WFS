#tag Module
Protected Module OSVersionInformationWFS
	#tag Method, Flags = &h1
		Protected Function Build() As Integer
		  Initialize
		  
		  return mBuild
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function BuildString() As String
		  Initialize
		  
		  return mBuildString
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ExtraInfo() As String
		  Initialize
		  
		  return mExtraInfo
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Initialize()
		  if mInitialized then return
		  mInitialized = true
		  
		  #if targetwin32 then
		    
		    Dim m As MemoryBlock
		    dim res as boolean
		    dim dwOSVersionInfoSize,wsuitemask,ret,i as integer
		    dim szCSDVersion, s as string
		    
		    Soft Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As ptr) As integer
		    Soft Declare Function GetVersionExW Lib "kernel32" (lpVersionInformation As ptr) As integer
		    
		    dim bUsingUnicode as Boolean = false
		    if System.IsFunctionAvailable( "GetVersionExW", "Kernel32" ) then
		      bUsingUnicode = true
		      m = new MemoryBlock(284) ''use this for osversioninfoex structure (2000+ only)
		      m.long(0) = m.size 'must set size before calling getversionex
		      ret = GetVersionExW(m) 'if not 2000+, will return 0
		      if ret = 0 then
		        m = new MemoryBlock(276)
		        m.long(0) = m.size 'must set size before calling getversionex
		        ret = GetVersionExW(m)
		        if ret = 0 then
		          // Something really strange has happened, so use the A version
		          // instead
		          bUsingUnicode = false
		          goto AVersion
		          return
		        end
		      end
		    else
		      AVersion:
		      m = new MemoryBlock(156) ''use this for osversioninfoex structure (2000+ only)
		      m.long(0) = m.size 'must set size before calling getversionex
		      ret = GetVersionExA(m) 'if not 2000+, will return 0
		      if ret = 0 then
		        m = new MemoryBlock(148) ' 148 sum of the bytes included in the structure (long = 4bytes, etc.)
		        m.long(0) = m.size 'must set size before calling getversionex
		        ret = GetVersionExA(m)
		        if ret = 0 then
		          return
		        end
		      end
		    end if
		    
		    mmajorVersion = m.long(4)
		    mminorVersion = m.long(8)
		    mbuild = m.long(12)
		    mplatformId = m.long(16)
		    
		    dim nextVal as Integer
		    if bUsingUnicode then
		      szCSDVersion = m.WString( 20 )
		      nextVal = 276
		    else
		      szCSDVersion = m.cstring(20)
		      nextVal = 148
		    end if
		    
		    try
		      mServicePackStr = Str( m.Short( nextVal ) ) + "." + Str( m.Short( nextVal + 2 ) )
		      if mServicePackStr = "0.0" then mServicePackStr = ""
		      
		      mServicePack = Val( mServicePackStr )
		      wsuitemask = m.short( nextVal + 4 )
		      msuitemask = hex( wSuitemask )
		      mproducttype = m.byte( nextVal + 6 )
		      
		    catch err as OutOfBoundsException
		      mServicePackStr = ""
		      mServicePack = 0
		      wsuitemask = 0
		      mSuiteMask = ""
		      mProductType = 0
		    end try
		    
		    if bitwiseAnd(val(msuitemask), &h1) > 0 then
		      mOSSuites  = "Small Business" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h2) > 0 then
		      mOSSuites  = mOSSuites + "Enterprise" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h4) > 0 then
		      mOSSuites  = mOSSuites + "BackOffice" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h8) > 0 then
		      mOSSuites  = mOSSuites + "Communications" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h10) > 0 then
		      mOSSuites  = mOSSuites + "Terminal Services" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h20) > 0 then
		      mOSSuites  = mOSSuites + "Small Business Restricted" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h40) > 0 then
		      mOSSuites  = mOSSuites + "Embedded NT" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h80) > 0 then
		      mOSSuites  = mOSSuites + "Data Center" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h100) > 0 then
		      mOSSuites  = mOSSuites + "Single User Terminal Services" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h200) > 0 then
		      mOSSuites  = mOSSuites + "Personal" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h400) > 0 then
		      mOSSuites  = mOSSuites + "Blade" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h800) > 0 then
		      mOSSuites  = mOSSuites + "Embedded Restricted" + chr(9)
		    end
		    if bitwiseAnd(wsuitemask, &h1000) > 0 then
		      mOSSuites  = mOSSuites + "Security Appliance" + chr(9)
		    end
		    
		    select case mproducttype
		    case 1
		      mmode = "Workstation"
		    case 2
		      mmode = "Domain Controller"
		    case 3
		      mmode = "Server"
		    end select
		    mextraInfo = szCSDVersion
		    if mplatformId = 2 then 'NT
		      mIsNx = true
		      if mMajorVersion = 6 then ' Vista
		        mIsVista = true
		      elseif mmajorversion = 4 then 'NT4
		        mOSName = "Windows NT 4.0"
		        mIsNT4 = true
		        mbuildstring = str(mbuild)
		      elseif mmajorVersion = 5 then
		        mIsNt5 = true
		        if mminorVersion = 0 then '2000
		          mOSName = "Windows 2000"
		          mIs2K = true
		          mbuildstring = str(build)
		        elseif mminorVersion = 1 then 'XP
		          mOSName = "Windows XP"
		          misXP = true
		          misXPPlus = true
		          mbuildstring = str(mbuild)
		        elseif mminorVersion = 2 then '2003
		          mOSName = "Windows 2003"
		          mis2003 = true
		          misXPPlus = true
		          mbuildstring = str(mbuild)
		        end if
		      end
		    else '9x
		      mis9x = true
		      if mminorVersion = 0 then ' 95
		        if instr(szCSDVersion,"C") > 0 or instr(szCSDVersion,"B") > 0 then
		          mOSName = "Windows 95 OSR2"
		        else
		          mOSName = "Windows 95"
		        end if
		        mis95 = true
		      elseif mminorVersion = 10 then '98
		        if instr(szCSDVersion,"A") > 0 then
		          mOSName = "Windows 98 SE"
		        else
		          mOSName = "Windows 98"
		        end if
		        mis98 = true
		        mbuildstring = "0"
		      elseif mminorVersion = 90 then 'ME
		        misME = true
		        mOSName = "Windows ME"
		        mbuildstring = "0"
		      end if
		    end if
		    s = str(majorVersion) + "." + str(mminorVersion) + "." + str(mbuild) + " " + mextraInfo
		    mversionstring = s
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Is2003() As Boolean
		  Initialize
		  
		  return mIs2003
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Is2k() As Boolean
		  Initialize
		  
		  return mIs2k
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Is95() As Boolean
		  Initialize
		  
		  return mIs95
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Is98() As Boolean
		  Initialize
		  
		  return mIs98
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Is9x() As Boolean
		  Initialize
		  
		  return mIs9x
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsME() As Boolean
		  Initialize
		  
		  return mIsME
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsNT4() As Boolean
		  Initialize
		  
		  return mIsNT4
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsNT5() As Boolean
		  Initialize
		  
		  return mIsNt5
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsNx() As Boolean
		  Initialize
		  
		  return mIsNx
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsTerminalSession() As Boolean
		  Declare Function GetSystemMetrics Lib "User32" Alias  "GetSystemMetrics" ( nIndex As integer ) As integer
		  
		  Const SM_REMOTESESSION = &h1000
		  
		  return GetSystemMetrics( SM_REMOTESESSION ) = 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsVista() As Boolean
		  Initialize
		  return mIsVista
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsXP() As Boolean
		  Initialize
		  
		  return mIsXP
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsXPPlus() As Boolean
		  Initialize
		  
		  return mIsXPPlus
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MajorVersion() As Integer
		  Initialize
		  
		  return mMajorVersion
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MinorVersion() As Integer
		  Initialize
		  
		  return mMinorVersion
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Mode() As String
		  Initialize
		  
		  return mMode
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function OSName() As String
		  Initialize
		  
		  return mOSName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function OSSuites() As String
		  Initialize
		  
		  return mOSSuites
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PlatformID() As Integer
		  Initialize
		  
		  return mPlatformID
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ProductType() As Integer
		  Initialize
		  
		  return mProductType
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ServicePack() As Double
		  return mServicePack
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ServicePack() As String
		  return mServicePackStr
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SuiteMask() As String
		  Initialize
		  
		  return mSuiteMask
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Version() As Integer
		  Initialize
		  
		  return mVersion
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Version() As String
		  Initialize
		  
		  return mVersionString
		End Function
	#tag EndMethod


	#tag Note, Name = Information
		
		Windows OS Version Detection Class
		
		GetVersionEx Win32 API called in class constructor
		
		Is95 True for WIndows 95 and Windows 95 OSR2
		Is98 True for Windows 98 and WIndows 98 SE
		IsME True for Windows ME
		IsNT4 True for WIndows NT 4.0
		Is2K True for Windows 2000
		IsXP True for Windows XP Home and Professional
		
		Is9x True for Windows 95,98,ME
		IsNx True for Windows NT,2000,XP.2003
		IsXPPLus True for Windows XP,2003
		
		For additional information:
		
		http://msdn.microsoft.com/library/en-us/ sysinfo/base/getversionex.asp 
		http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/osversioninfoex_str.asp
		http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/osversioninfo_str.asp
	#tag EndNote

	#tag Note, Name = VB Type info
		VB Type definitions
		Private Type OSVERSIONINFO
		OSVSize         As Long         'size, in bytes, of this data structure
		dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
		dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
		dwBuildNumber   As Long         'NT: build number of the OS
		'Win9x: build number of the OS in low-order word.
		'High-order word contains major & minor ver nos.
		PlatformID      As Long         'Identifies the operating system platform.
		szCSDVersion    As String * 128 'NT: string such as "Service Pack 3"
		'Win9x: arbitrary additional information
		End Type
		
		Private Type OSVERSIONINFOEX
		' I tacked the position in the memory block to the end of these for easier reference
		OSVSize            As Long ''0
		dwVerMajor        As Long ''4
		dwVerMinor         As Long ''8
		dwBuildNumber      As Long ''12
		PlatformID         As Long ''16
		szCSDVersion       As String * 128 ''20   ''' Fixed width string 128 bytes
		wServicePackMajor  As Integer ''148
		wServicePackMinor  As Integer ''150
		wSuiteMask         As Integer ''152
		wProductType       As Byte ''154  
		wReserved          As Byte ''155
		End Type
	#tag EndNote


	#tag Property, Flags = &h21
		Private mBuild As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mBuildString As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mExtraInfo As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialized As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mis2003 As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIs2K As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIs95 As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIs98 As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIs9x As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsME As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsNT4 As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsNt5 As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsNx As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsVista As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsXP As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsXPPlus As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMajorVersion As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMinorVersion As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMode As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mOSName As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mOSSuites As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPlatformID As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mProductType As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mServicePack As Double
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mServicePackStr As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSuiteMask As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVersion As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVersionString As string
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
