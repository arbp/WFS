#tag Class
Protected Class ErrorReport
	#tag Method, Flags = &h0
		Sub AddDump(type as Integer, flags as Integer)
		  #if TargetWin32
		    Soft Declare Sub WerReportAddDump Lib "Wer" ( handle as Integer, processHandle as Integer, threadHandle as Integer, _
		    type as Integer, exceptionInfo as Integer, customOptions as Integer, flags as Integer )
		    
		    if System.IsFunctionAvailable( "WerReportAddDump", "Wer" ) and mHandle <> 0 then
		      // Get the process handle first
		      Declare Function GetModuleHandleW Lib "Kernel32" ( name as Integer ) as Integer
		      Dim processHandle as Integer = GetModuleHandleW( 0 )
		      
		      // Now, if the dump type is 'micro', we need to have a valid
		      // thread handle as well.
		      Dim threadHandle as Integer = 0
		      if type = kDumpMicro then
		        Declare Function GetCurrentThread Lib "Kernel32" () as Integer
		        threadHandle = GetCurrentThread
		      end if
		      
		      // Now we can add the dump
		      WerReportAddDump( mHandle, processHandle, threadHandle, type, 0, 0, flags )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddFile(f as FolderItem, type as Integer, flags as Integer)
		  #if TargetWin32
		    Soft Declare Sub WerReportAddFile Lib "Wer" ( handle as Integer, path as WString, type as Integer, flags as Integer )
		    
		    if System.IsFunctionAvailable( "WerReportAddFile", "Wer" ) and mHandle <> 0 then
		      WerReportAddFile( mHandle, f.AbsolutePath, type, flags )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(eventType as String, reportType as Integer)
		  #if TargetWin32
		    Soft Declare Sub WerReportCreate Lib "Wer" ( eventType as WString, reportType as Integer, reportInfo as Integer, ByRef handle as Integer )
		    
		    if System.IsFunctionAvailable( "WerReportCreate", "Wer" ) then
		      WerReportCreate( eventType, reportType, 0, mHandle )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  #if TargetWin32
		    Soft Declare Sub WerReportCloseHandle Lib "Wer" ( handle as Integer )
		    
		    if System.IsFunctionAvailable( "WerReportCloseHandle", "Wer" ) and mHandle <> 0 then
		      WerReportCloseHandle( mHandle )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Sub RegisterFile(f as FolderItem, type as Integer, flags as Integer)
		  #if TargetWin32
		    Soft Declare Sub WerRegisterFile Lib "Wer" ( path as WString, type as Integer, flags as Integer )
		    
		    if System.IsFunctionAvailable( "WerRegisterFile", "Wer" ) then
		      WerRegisterFile( f.AbsolutePath, type, flags )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Sub RegisterMemoryBlock(mb as MemoryBlock)
		  #if TargetWin32
		    Soft Declare Sub WerRegisterMemoryBlock Lib "Wer" ( block as Ptr, size as Integer )
		    
		    if System.IsFunctionAvailable( "WerRegisterMemoryBlock", "Wer" ) then
		      WerRegisterMemoryBlock( mb, mb.Size )
		    end if
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Send(consent as Integer, flags as Integer) As Integer
		  #if TargetWin32
		    Soft Declare Sub WerReportSubmit Lib "Wer" ( handle as Integer, consent as Integer, flags as Integer, ByRef result as Integer )
		    
		    if System.IsFunctionAvailable( "WerReportSubmit", "Wer" ) then
		      dim ret as Integer
		      if mHandle <> 0 then
		        WerReportSubmit( mHandle, consent, flags, ret )
		        return ret
		      end if
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub SetUIOptions(option as Integer, text as String)
		  #if TargetWin32
		    Soft Declare Sub WerReportSetUIOption Lib "Wer" ( handle as Integer, ui as Integer, customText as WString )
		    
		    if System.IsFunctionAvailable( "WerReportSetUIOption", "Wer" ) and mHandle <> 0 then
		      WerReportSetUIOption( mHandle, option, text )
		    end if
		  #endif
		End Sub
	#tag EndMethod


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mAdditionalDataHeader
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mAdditionalDataHeader = value
			  
			  Const WerUIAdditionalDataDlgHeader = 1
			  SetUIOptions( WerUIAdditionalDataDlgHeader, value )
			End Set
		#tag EndSetter
		AdditionalDataHeader As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mClosebody
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mClosebody = value
			  
			  Const WerUICloseDlgBody = 9
			  SetUIOptions( WerUICloseDlgBody, value )
			End Set
		#tag EndSetter
		CloseBody As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mClosebuttoncaption
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mClosebuttoncaption = value
			  
			  Const WerUICloseDlgButtonText = 10
			  SetUIOptions( WerUICloseDlgButtonText, value )
			End Set
		#tag EndSetter
		CloseButtonCaption As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mCloseheader
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mCloseheader = value
			  
			  Const WerUICloseDlgHeader = 8
			  SetUIOptions( WerUICloseDlgHeader, value )
			End Set
		#tag EndSetter
		CloseHeader As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mClosetext
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mClosetext = value
			  
			  Const WerUICloseText = 7
			  SetUIOptions( WerUICloseText, value )
			End Set
		#tag EndSetter
		CloseText As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mConsentbody
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mConsentbody = value
			  
			  Const WerUIConsentDlgBody = 4
			  SetUIOptions( WerUIConsentDlgBody, value )
			End Set
		#tag EndSetter
		ConsentBody As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mConsentheader
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mConsentheader = value
			  
			  Const WerUIConsentDlgHeader = 3
			  SetUIOptions( WerUIConsentDlgHeader, value )
			End Set
		#tag EndSetter
		ConsentHeader As String
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private mAdditionalDataHeader As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mClosebody As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mClosebuttoncaption As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCloseheader As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mClosetext As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mConsentbody As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mConsentHeader As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mOfflinesolutionchecktext As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mOnlinesolutionchecktext As String
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mOfflinesolutionchecktext
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mOfflinesolutionchecktext = value
			  
			  Const WerUIOnlineSolutionCheckText = 6
			  SetUIOptions( WerUIOnlineSolutionCheckText, value )
			End Set
		#tag EndSetter
		OfflineSolutionCheckText As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mOnlinesolutionchecktext
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mOnlinesolutionchecktext = value
			  
			  Const WerUIOnlineSolutionCheckText = 5
			  SetUIOptions( WerUIOnlineSolutionCheckText, value )
			End Set
		#tag EndSetter
		OnlineSolutionCheckText As String
	#tag EndComputedProperty


	#tag Constant, Name = kAddFileTypeHeapdump, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kAddFileTypeMicrodump, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kAddFileTypeMinidump, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kAddFileTypeOther, Type = Double, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kAddFileTypeUserDocument, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kConsentApproved, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kConsentDenied, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kConsentNotAsked, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kDumpHeap, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kDumpMicro, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kDumpMini, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kRegisterFileTypeOther, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kRegisterFileTypeUserDocument, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagHonorAppRecovery, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagHonorAppRestart, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagNoArchive, Type = Double, Dynamic = False, Default = \"256", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagNoCloseUI, Type = Double, Dynamic = False, Default = \"64", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagNoQueue, Type = Double, Dynamic = False, Default = \"128", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagOutOfProcess, Type = Double, Dynamic = False, Default = \"32", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagOutOfProcessAsync, Type = Double, Dynamic = False, Default = \"1024", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagQueue, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagShowDebug, Type = Double, Dynamic = False, Default = \"8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportFlagStartMinimized, Type = Double, Dynamic = False, Default = \"512", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportTypeAddRegisteredData, Type = Double, Dynamic = False, Default = \"16", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportTypeApplicationCrash, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportTypeApplicationHang, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportTypeCritical, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportTypeKernel, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kReportTypeNonCritical, Type = Double, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitReportAsync, Type = Double, Dynamic = False, Default = \"8", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitReportDisabledQueue, Type = Double, Dynamic = False, Default = \"7", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitResultDebug, Type = Double, Dynamic = False, Default = \"3", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitResultFailed, Type = Double, Dynamic = False, Default = \"4", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitResultQueued, Type = Double, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitResultReportCancelled, Type = Double, Dynamic = False, Default = \"6", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitResultUploaded, Type = Double, Dynamic = False, Default = \"2", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSubmitResultWERDisabled, Type = Double, Dynamic = False, Default = \"5", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="AdditionalDataHeader"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CloseBody"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CloseButtonCaption"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CloseHeader"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CloseText"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ConsentBody"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ConsentHeader"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
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
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="OfflineSolutionCheckText"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="OnlineSolutionCheckText"
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
	#tag EndViewBehavior
End Class
#tag EndClass
