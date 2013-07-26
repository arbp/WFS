#tag Window
Begin Window WFSWindow
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   3.0e+2
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   858513407
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   True
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   False
   Title           =   "WFS"
   Visible         =   True
   Width           =   4.67e+2
   Begin Label lblDate
      AutoDeactivate  =   True
      Bold            =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   30
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   238
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Multiline       =   ""
      Scope           =   0
      Selectable      =   False
      TabIndex        =   0
      TabPanelIndex   =   0
      Text            =   "Untitled"
      TextAlign       =   0
      TextColor       =   &h000000
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   10
      Transparent     =   False
      Underline       =   ""
      Visible         =   True
      Width           =   176
   End
   Begin Label lblTime
      AutoDeactivate  =   True
      Bold            =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   30
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   238
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Multiline       =   ""
      Scope           =   0
      Selectable      =   False
      TabIndex        =   1
      TabPanelIndex   =   0
      Text            =   "Untitled"
      TextAlign       =   0
      TextColor       =   &h000000
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   50
      Transparent     =   False
      Underline       =   ""
      Visible         =   True
      Width           =   176
   End
   Begin Timer tmrUpdateUI
      Height          =   32
      Index           =   -2147483648
      Left            =   568
      LockedInPosition=   False
      Mode            =   2
      Period          =   500
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   13
      Width           =   32
   End
   Begin PushButton btnSetSystemTime
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "SetSystemTime"
      Default         =   ""
      Enabled         =   True
      Height          =   25
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   20
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   92
      Underline       =   ""
      Visible         =   True
      Width           =   176
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  DatePicker = new DateTimePickerWFS
		  DatePicker.Create( self, 20, 10, 150, 30, true )
		  
		  TimePicker = new DateTimePickerWFS
		  TimePicker.Create( self, 20, 50, 150, 30, false )
		  
		  UpdateDateTimeLabels()
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub UpdateDateTimeLabels()
		  dim d as date = DatePicker.GetDateTime
		  dim t as date = TimePicker.GetDateTime
		  
		  lblDate.Text = d.SQLDateTime
		  lblTime.Text = t.SQLDateTime
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		DatePicker As DateTimePickerWFS
	#tag EndProperty

	#tag Property, Flags = &h0
		TimePicker As DateTimePickerWFS
	#tag EndProperty


#tag EndWindowCode

#tag Events tmrUpdateUI
	#tag Event
		Sub Action()
		  UpdateDateTimeLabels
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnSetSystemTime
	#tag Event
		Sub Action()
		  dim d as date = DatePicker.GetDateTime
		  dim t as date = TimePicker.GetDateTime
		  
		  d.Hour = t.Hour
		  d.Minute = t.Minute
		  d.Second = t.Second
		  
		  if SystemInformationWFS.SetSystemTime( d ) then
		    MsgBox "SetSystemTime succeeded."
		  else
		    MsgBox "SetSystemTime FAILED!"
		  end if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
