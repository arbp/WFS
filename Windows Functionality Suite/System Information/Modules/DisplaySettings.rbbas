#tag Module
Protected Module DisplaySettings
	#tag Method, Flags = &h1
		Protected Function BitsPerPel() As integer
		  Initialize
		  return mBitsPerPel
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DeviceName() As String
		  Initialize
		  return mDeviceName
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DisplayFlags() As integer
		  Initialize
		  return mDisplayFlags
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DisplayFrequency() As integer
		  Initialize
		  return mDisplayFrequency
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DriverExtra() As integer
		  Initialize
		  return mDriverExtra
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DriverVersion() As integer
		  Initialize
		  return mDriverVersion
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Initialize()
		  if mInitialized then return
		  mInitialized = true
		  
		  #if TargetWin32
		    Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As integer, iModeNum As integer, lpDevMode As ptr) As integer
		    Const ENUM_CURRENT_SETTINGS = -1
		    Const ENUM_REGISTRY_SETTINGS = -2
		    dim DevM as memoryBlock
		    dim ret as integer
		    dim i as integer
		    dim x,y as integer
		    DevM = new MemoryBlock(124) 'size as figured below
		    DevM.long(36) = devm.size 'set size member before call
		    ret = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, DevM)
		    
		    mDeviceName = DevM.cstring(0)
		    mSpecVersion = DevM.short(32)
		    mDriverVersion = DevM.short(34)
		    mDriverExtra = DevM.short(38)
		    mBitsPerPel = DevM.short(104)
		    mPelsWidth = DevM.long(108)
		    mPelsHeight = DevM.long(112)
		    mDisplayFlags = DevM.long(116)
		    mDisplayFrequency = DevM.long(120)
		    
		    'Const CCDEVICENAME = 32
		    'Const CCFORMNAME = 32
		    
		    'Type DEVMODE size = 124
		    'dmDeviceName As String * CCDEVICENAME 0
		    'dmSpecVersion As Integer 32
		    'dmDriverVersion As Integer 34
		    'dmSize As Integer 36
		    'dmDriverExtra As Integer 38
		    'dmFields As Long 40
		    'dmOrientation As Integer 44
		    'dmPaperSize As Integer 46
		    'dmPaperLength As Integer 48
		    'dmPaperWidth As Integer 50
		    'dmScale As Integer 52
		    'dmCopies As Integer  54
		    'dmDefaultSource As Integer  56
		    'dmPrintQuality As Integer  58
		    'dmColor As Integer  60
		    'dmDuplex As Integer  62
		    'dmYResolution As Integer 64
		    'dmTTOption As Integer 66
		    'dmCollate As Integer 68
		    'dmFormName As String * CCFORMNAME 70
		    'dmUnusedPadding As Integer 102
		    'dmBitsPerPel As Integer 104
		    'dmPelsWidth As Long 108
		    'dmPelsHeight As Long 112
		    'dmDisplayFlags As Long 116
		    'dmDisplayFrequency As Long 120
		    'End Type
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PelsHeight() As integer
		  Initialize
		  return mPelsHeight
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PelsWidth() As integer
		  Initialize
		  return mPelsWidth
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SpecVersion() As integer
		  Initialize
		  return mSpecVersion
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mBitsPerPel As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDeviceName As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDisplayFlags As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDisplayFrequency As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDriverExtra As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDriverVersion As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInitialized As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPelsHeight As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPelsWidth As integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mSpecVersion As integer
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
