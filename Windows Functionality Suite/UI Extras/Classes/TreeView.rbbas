#tag Class
Protected Class TreeView
Implements WndProcSubclass
	#tag Method, Flags = &h21
		Private Function AddImageToImageList(p as Picture) As Integer
		  #if TargetWin32
		    Declare Function ImageList_Create Lib "comctl32" ( cx as Integer, cy as Integer, _
		    flags as Integer, initialSize as Integer, growBy as Integer ) as Integer
		    
		    Const ILC_COLOR4              = &h0004
		    Const ILC_COLOR8              = &h0008
		    Const ILC_COLOR16             = &h0010
		    Const ILC_COLOR24             = &h0018
		    Const ILC_COLOR32             = &h0020
		    
		    Dim TVM_SETIMAGELIST as Integer = TV_FIRST + 9
		    
		    // We have two lists, so which list we work on
		    // may change.  If the list hasn't been created
		    // yet, then we need to make it now as well.
		    dim flags as Integer
		    select case p.Depth
		    case 32
		      flags = ILC_COLOR32
		    case 24
		      flags = ILC_COLOR24
		    case 16
		      flags = ILC_COLOR16
		    case 8
		      flags = ILC_COLOR8
		    case 4
		      flags = ILC_COLOR4
		    end
		    if mImageList = 0 then
		      // Make the new image list
		      mImageList = ImageList_Create( p.Width, p.Height, flags, 1, 1 )
		      Call SendMessage( TVM_SETIMAGELIST, 0, mImageList )
		    end
		    
		    // Now we have an image list to work with, so let's
		    // add the picture.  This is rather much a pain in
		    // the arse, but that's ok, we can do it  using declares!
		    // We need to make an HBITMAP from the image that
		    // also uses the mask properly
		    dim hbitmap as Integer
		    hbitmap = HBITMAPFromPicture( p )
		    
		    // Now that we have the HBITMAP from the Picture
		    // object, we can add it to the image list
		    Declare Function ImageList_Add Lib "comctl32" ( list as Integer, _
		    image as Integer, mask as Integer ) as Integer
		    
		    dim index as Integer
		    index = ImageList_Add( mImageList, hbitmap, 0 )
		    
		    // Now that we're done with the bitmap, we should
		    // release it
		    Declare Sub DeleteObject Lib "Gdi32" ( obj as Integer )
		    DeleteObject( hbitmap )
		    
		    return index
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddItem(ByRef item as TreeViewItem, parent as TreeViewItem = nil)
		  #if TargetWin32
		    // Set up the proper structure we need to add a new item in
		    // the TreeView.
		    Dim TVM_INSERTITEMA as Integer = TV_FIRST + 0
		    Dim TVM_INSERTITEMW as Integer = TV_FIRST + 50
		    Dim tvinsertstruct as new MemoryBlock( 52 )
		    
		    // Set the parent handle properly.  This allows us to add
		    // sub items to anything passed into us.
		    if parent <> nil then
		      tvinsertstruct.Long( 0 ) = parent.Handle
		    else
		      tvinsertstruct.Long( 0 ) = 0
		    end
		    
		    // Determine where you would like this item to appear
		    // in relationship to the parent item.
		    Const TVI_ROOT                = &hFFFF0000
		    Const TVI_FIRST               = &hFFFF0001
		    Const TVI_LAST                = &hFFFF0002
		    Const TVI_SORT                = &hFFFF0003
		    
		    if parent = nil then
		      tvinsertstruct.Long( 4 ) = TVI_ROOT
		    else
		      tvinsertstruct.Long( 4 ) = TVI_LAST
		    end
		    
		    // Next, figure out what flags are valid with this
		    // object.
		    Const TVIF_TEXT               = &h0001
		    Const TVIF_IMAGE              = &h0002
		    Const TVIF_PARAM              = &h0004
		    Const TVIF_STATE              = &h0008
		    Const TVIF_HANDLE             = &h0010
		    Const TVIF_SELECTEDIMAGE      = &h0020
		    Const TVIF_CHILDREN           = &h0040
		    
		    dim flags as Integer
		    dim imageIndex, selectedImageIndex as Integer = -1
		    
		    // Every item has some text to it, so this flag is always valid
		    flags = TVIF_TEXT
		    
		    // If we have an image, then we need to add it to the image
		    // map
		    if item.Image <> nil then
		      flags = flags + TVIF_IMAGE
		      imageIndex = AddImageToImageList( item.Image )
		      // We will fall back on this image to be our drag image
		      item.DragImageIndex = imageIndex
		    end
		    
		    // If we have a different image to show when this item is
		    // selected, we should add that to the image list as well.
		    if item.SelectedImage <> nil then
		      flags = flags + TVIF_SELECTEDIMAGE
		      selectedImageIndex = AddImageToImageList( item.SelectedImage )
		      // We want to use this image when dragging the item since
		      // the item must be selected in order to start dragging it.
		      item.DragImageIndex = selectedImageIndex
		    end
		    
		    // Set the flags we're using
		    tvinsertstruct.Long( 8 ) = flags
		    
		    // Set the new item's text
		    dim text as MemoryBlock
		    if mUnicodeSavvy then
		      dim theText as String = ConvertEncoding( Item.Text, Encodings.UTF16 )
		      text = new MemoryBlock( LenB( theText ) + 2 )
		      text.WString( 0 ) = theText
		      
		      tvinsertstruct.Ptr( 24 ) = text
		      tvinsertstruct.Long( 28 ) = Len( theText ) + 1
		    else
		      text = new MemoryBlock( Len( item.Text ) + 1 )
		      text.CString( 0 ) = Item.Text
		      tvinsertstruct.Ptr( 24 ) = text
		      tvinsertstruct.Long( 28 ) = text.Size
		    end if
		    
		    // Then be sure to set the proper image indexes if we have them
		    if imageIndex >= 0 then
		      tvinsertstruct.Long( 32 ) = imageIndex
		    end
		    if selectedImageIndex >= 0 then
		      tvinsertstruct.Long( 36 ) = selectedImageIndex
		    end
		    
		    // After all that jazz, add the item itself
		    if mUnicodeSavvy then
		      item.Handle = SendMessage( TVM_INSERTITEMW, 0, MemoryBlockToInteger( tvinsertstruct ) )
		    else
		      item.Handle = SendMessage( TVM_INSERTITEMA, 0, MemoryBlockToInteger( tvinsertstruct ) )
		    end if
		    
		    // And make sure this item's handle maps to the actual item in
		    // our list
		    if mItems = nil then mItems = new Dictionary
		    mItems.Value( item.Handle ) = item
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub BeginDrag(lpnmtreeview as MemoryBlock)
		  #if TargetWin32
		    // First, we want to make the proper image list for the drag item
		    Dim TVM_CREATEDRAGIMAGE as Integer = TV_FIRST + 18
		    Dim oldItemOffset as Integer = 16
		    Dim newItemOffset as Integer = oldItemOffset + 40
		    
		    // Create the drag item from the tree view
		    mDragging = SendMessage( TVM_CREATEDRAGIMAGE, 0, lpnmtreeview.Long( newItemOffset + 4 ) )
		    
		    // Now tell the image list to start dragging this
		    Declare Sub ImageList_BeginDrag Lib "ComCtl32" ( hImg as Integer, index as Integer, _
		    dx as Integer, dy as Integer )
		    ImageList_BeginDrag( mDragging, 0, 0, 0 )
		    
		    // Set where to start the drag from
		    Dim x, y as Integer
		    Declare Sub ImageList_DragEnter Lib "ComCtl32" ( hOwner as Integer, x as Integer, y as Integer )
		    x = lpnmtreeview.Long( newItemOffset + 40 )
		    y = lpnmtreeview.Long( newItemOffset + 44 )
		    ImageList_DragEnter( mWnd, x, y )
		    
		    // Hide the mouse cursor
		    Declare Sub ShowCursor Lib "User32" ( show as Boolean )
		    ShowCursor( false )
		    
		    // And capture all mouse movements in this window
		    Declare Sub SetCapture Lib "User32" ( hwnd as Integer )
		    SetCapture( mWnd )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  #if TargetWin32
		    // We need to initialize the common controls library and tell it
		    // that we are going to be adding treeviews
		    Declare Sub InitCommonControlsEx Lib "comctl32" ( ex as Ptr )
		    
		    dim ex as new MemoryBlock( 8 )
		    ex.Long( 0 ) = ex.Size
		    Const ICC_TREEVIEW_CLASSES = &h00000002
		    ex.Long( 4 ) = ICC_TREEVIEW_CLASSES
		    
		    InitCommonControlsEx( ex )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Create(w as Window, x as Integer, y as Integer, width as Integer, height as Integer)
		  #if TargetWin32
		    Const WC_TREEVIEW = "SysTreeView32"
		    
		    Soft Declare Function CreateWindowExA Lib "User32" ( ex as Integer, className as CString, _
		    title as Integer, style as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, parent as Integer, menu as Integer, hInstance as Integer, _
		    lParam as Integer ) as Integer
		    Soft Declare Function CreateWindowExW Lib "User32" ( ex as Integer, className as WString, _
		    title as Integer, style as Integer, x as Integer, y as Integer, width as Integer, _
		    height as Integer, parent as Integer, menu as Integer, hInstance as Integer, _
		    lParam as Integer ) as Integer
		    
		    Declare Function GetModuleHandleA Lib "Kernel32" ( name as Integer ) as Integer
		    
		    // First, get the HINSTANCE of this application so we can
		    // use it to create a window
		    dim hInstance as Integer
		    hInstance = GetModuleHandleA( 0 )
		    
		    Const TVS_HASBUTTONS          = &h0001
		    Const TVS_HASLINES            = &h0002
		    Const TVS_LINESATROOT         = &h0004
		    Const TVS_EDITLABELS          = &h0008
		    Const TVS_DISABLEDRAGDROP     = &h0010
		    Const TVS_SHOWSELALWAYS       = &h0020
		    Const TVS_RTLREADING          = &h0040
		    Const TVS_NOTOOLTIPS          = &h0080
		    Const TVS_CHECKBOXES          = &h0100
		    Const TVS_TRACKSELECT         = &h0200
		    Const TVS_SINGLEEXPAND        = &h0400
		    Const TVS_INFOTIP             = &h0800
		    Const TVS_FULLROWSELECT       = &h1000
		    Const TVS_NOSCROLL            = &h2000
		    Const TVS_NONEVENHEIGHT       = &h4000
		    
		    Const WS_CHILD            = &h40000000
		    Const WS_EX_CLIENTEDGE        = &h00000200
		    
		    // Figure out what style we want.  Right now we just pick the style
		    // for the user, but this could be pretty easily user-configurable.
		    dim style as Integer
		    style = TVS_HASLINES + TVS_LINESATROOT + TVS_HASBUTTONS + TVS_EDITLABELS
		    
		    // For now, we're not going to allow dragging.  It seems to be
		    // messed up because the RB Framework captures the mouse out
		    // from underneath us.
		    style = style + TVS_DISABLEDRAGDROP
		    
		    Const SB_SETUNICODEFORMAT = &h2005
		    
		    mUnicodeSavvy = System.IsFunctionAvailable( "CreateWindowExW", "User32" )
		    
		    // Actually create the window that is our treeview
		    if mUnicodeSavvy then
		      mWnd = CreateWindowExW( WS_EX_CLIENTEDGE, WC_TREEVIEW, 0, WS_CHILD + style, _
		      x, y, width, height, w.WinHWND, 0, hInstance, 0 )
		      
		      Call SendMessage( SB_SETUNICODEFORMAT, 1, 0 )
		    else
		      mWnd = CreateWindowExA( WS_EX_CLIENTEDGE, WC_TREEVIEW, 0, WS_CHILD + style, _
		      x, y, width, height, w.WinHWND, 0, hInstance, 0 )
		      
		      Call SendMessage( SB_SETUNICODEFORMAT, 0, 0 )
		    end if
		    
		    // Make sure the window is shown properly
		    Declare Sub ShowWindow Lib "User32" ( hwnd as Integer, type as Integer )
		    Const SW_SHOW = 5
		    ShowWindow( mWnd, SW_SHOW )
		    
		    // Subclass the RB Window so that we can get messages
		    // for this window.  Be sure to keep the RB Window around tho
		    WndProcHelpers.Subclass( w, self )
		    mRBWindow = w
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Destructor()
		  // If we've been subclassed, we should unsubclass
		  // ourselves
		  if mRBWindow <> nil then
		    WndProcHelpers.Unsubclass( mRBWindow, self )
		  end
		  
		  #if TargetWin32
		    // And then destroy our window if we have one made
		    Declare Sub DestroyWindow Lib "User32" ( hwnd as Integer )
		    
		    if mWnd <> 0 then
		      DestroyWindow( mWnd )
		    end
		    mWnd = 0
		    
		    // Also be sure to destroy our image list when we're done
		    Declare Sub ImageList_Destroy Lib "ComCtl32" ( list as Integer )
		    if mImageList <> 0 then
		      ImageList_Destroy( mImageList )
		    end
		    mImageList = 0
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EnsureItemIsVisible(item as TreeViewItem)
		  #if TargetWin32
		    // We want to make sure this item is visible to the
		    // user, and that all parents are properly expanded
		    // to show the item
		    Dim TVM_ENSUREVISIBLE as Integer = TV_FIRST + 20
		    
		    Call SendMessage( TVM_ENSUREVISIBLE, 0, item.Handle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleWM_LButtonUp() As Integer
		  #if TargetWin32
		    // If we're not dragging, then bail out
		    if mDragging = 0 then return 0
		    
		    // Signal that we're leaving the drag
		    Declare Sub ImageList_DragLeave Lib "ComCtl32" ( htree as Integer )
		    ImageList_DragLeave( mWnd )
		    
		    // And end the drag
		    Declare Sub ImageList_EndDrag Lib "ComCtl32" ()
		    ImageList_EndDrag()
		    
		    // Get the selected item
		    DIm TVM_GETNEXTITEM as Integer = TV_FIRST + 10
		    Const TVGN_DROPHILITE = &h0008
		    Dim selected as Integer
		    selected = SendMessage( TVM_GETNEXTITEM, TVGN_DROPHILITE, 0 )
		    
		    // Now select it properly
		    Dim TVM_SELECTITEM as Integer = TV_FIRST + 11
		    Const TVGN_CARET = &h0009
		    
		    // First select the caret
		    Call SendMessage( TVM_SELECTITEM, TVGN_CARET, selected )
		    // Then the highlight
		    Call SendMessage( TVM_SELECTITEM, TVGN_DROPHILITE, selected )
		    
		    // Then release our capture on the mouse
		    Declare Sub ReleaseCapture Lib "User32" ()
		    ReleaseCapture()
		    
		    // And re-show the regular cursor
		    Declare Sub ShowCursor Lib "User32" ( show as Boolean )
		    ShowCursor( true )
		    
		    // And destroy our image list
		    Declare Sub ImageList_Destroy Lib "ComCtl32" ( list as Integer )
		    ImageList_Destroy( mDragging )
		    
		    mDragging = 0
		    
		    return 1
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleWM_MouseMove(lParam as Integer) As Integer
		  #if TargetWin32
		    // If we're not dragging, then we don't care
		    if mDragging = 0 then return 0
		    
		    // We're dragging, so we care.  Tell the image list that
		    // the drag has moved
		    Declare Sub ImageList_DragMove Lib "ComCtl32" ( x as Integer, y as Integer )
		    dim x, y as Integer
		    x = BitwiseAnd( lParam, &hFFFF )
		    y = BitwiseAnd( Bitwise.ShiftRight( lParam, 16 ), &hFFFF )
		    ImageList_DragMove( x - 32, y - 25 )
		    
		    // Lock the new image from showing yet
		    Declare Sub ImageList_DragShowNolock Lib "ComCtl32" ( show as Boolean )
		    ImageList_DragShowNolock( false )
		    
		    // We should hilight the item that we're dragging over so it's
		    // proper UI behavior
		    Dim TVM_HITTEST as Integer = TV_FIRST + 17
		    Dim TVM_SELECTITEM as Integer = TV_FIRST + 11
		    Const TVGN_DROPHILITE = &h0008
		    
		    Dim tvhittestinfo as new MemoryBlock( 16 )
		    tvhittestinfo.Long( 0 ) = x - 20
		    tvhittestinfo.Long( 4 ) = y - 20
		    
		    // Get the hit target
		    dim hitTarget as Integer
		    hitTarget = SendMessage( TVM_HITTEST, 0, MemoryBlockToInteger( tvhittestinfo ) )
		    if hitTarget <> 0 then
		      // There was a hit target, so select it for highlighting.
		      Call SendMessage( TVM_SELECTITEM, TVGN_DROPHILITE, hitTarget )
		    end
		    
		    // Now show the new image
		    ImageList_DragShowNolock( true )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HandleWM_Notify(hWnd as Integer, controlID as Integer, nmhdr as MemoryBlock) As Integer
		  // Check to see if the code for the nmhdr is one that
		  // we care about
		  dim code as Integer = nmhdr.Long( 8 )
		  Dim TVN_SELCHANGEDA as Integer = TVN_FIRST - 2
		  Dim TVN_SELCHANGEDW as Integer = TVN_FIRST - 51
		  DIm TVN_ENDLABELEDITA as Integer = TVN_FIRST - 11
		  Dim TVN_ENDLABELEDITW as Integer = TVN_FIRST - 60
		  Dim TVN_BEGINDRAGA as Integer = TVN_FIRST - 7
		  Dim TVN_BEGINDRAGW as Integer = TVN_FIRST - 56
		  
		  Dim oldItemOffset as Integer = 16
		  Dim newItemOffset as Integer = oldItemOffset + 40
		  dim oldItemHandle, newItemHandle as Integer
		  dim item as TreeViewItem
		  Dim textLen as Integer
		  
		  Const TVIF_TEXT               = &h0001
		  Const TVIF_IMAGE              = &h0002
		  Const TVIF_PARAM              = &h0004
		  Const TVIF_STATE              = &h0008
		  Const TVIF_HANDLE             = &h0010
		  Const TVIF_SELECTEDIMAGE      = &h0020
		  Const TVIF_CHILDREN           = &h0040
		  
		  if code = TVN_SELCHANGEDA or code = TVN_SELCHANGEDW then
		    // If this is the case, then the nmhdr is really a NMTREEVIEW
		    // We should figure out who's been unselected, and
		    // fire off an event for them.
		    
		    // Get a handle to the old item
		    if BitwiseAnd( nmhdr.Long( oldItemOffset + 0 ), TVIF_HANDLE ) = TVIF_HANDLE  then
		      oldItemHandle = nmhdr.Long( oldItemOffset + 4 )
		    end
		    
		    if oldItemHandle <> 0 then
		      ItemUnselected( mItems.Value( oldItemHandle ) )
		    end
		    
		    // Get a handle to the new item
		    if BitwiseAnd( nmhdr.Long( newItemOffset + 0 ), TVIF_HANDLE ) = TVIF_HANDLE  then
		      newItemHandle = nmhdr.Long( newItemOffset + 4 )
		    end
		    
		    if newItemHandle <> 0 then
		      ItemSelected( mItems.Value( newItemHandle ) )
		    end
		  elseif code = TVN_ENDLABELEDITA or code = TVN_ENDLABELEDITW then
		    // The user is done editing a label.  We really have a
		    // NMTVDISPINFO structure that points to the new
		    // information
		    if nmhdr.Long( 28 ) <> 0 then
		      // Get the new text and stuff it into the proper item
		      item = mItems.Value( nmhdr.Long( 16 ) )
		      if code = TVN_ENDLABELEDITA then
		        if mUnicodeSavvy then
		          item.Text = nmhdr.Ptr( 28 ).WString( 0 )
		        else
		          item.Text = nmhdr.Ptr( 28 ).CString( 0 )
		        end if
		      else
		        item.Text = nmhdr.Ptr( 28 ).WString( 0 )
		      end
		      mItems.Value( nmhdr.Long( 16 ) ) = item
		      
		      // We want to accept the label text change
		      return 1
		    end
		  elseif code = TVN_BEGINDRAGA or code = TVN_BEGINDRAGW then
		    // We're starting a drag operation, so get that ball rolling
		    BeginDrag( nmhdr )
		  end
		  
		  return 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HBITMAPFromPicture(p as Picture) As Integer
		  #if TargetWin32
		    #if RBVersion < 2010.05
		      // DISCLAIMER!!!  READ THIS!!!!
		      // The following declare uses a very unsupported feature in REALbasic.  This means a few
		      // things.  1) Don't rely on this call working forever -- it's entirely possible that upgrading to
		      // a new version of REALbasic will cause this declare to break.  2) Don't try to use declares
		      // into the plugins APIs yourself.  And 3) Since this is not a supported way to use declares, you
		      // get everything you deserve if you use a declare like this.  You will eventually come down
		      // with the plague and someone will drink all your beer.  Basically... don't use this sort of thing!
		      // The declare will be removed as soon as there is a sanctioned way to get this functionality
		      // from REALbasic.  This declare should (hypothetically) work in 5.5 only (tho it may work in
		      // version 5 as well).  Consider all other versions of REALbasic unsupported.
		      #if RBVersion < 2006.01 then
		        Declare Function REALGraphicsDC Lib "" ( g as Graphics ) as Integer
		      #endif
		      
		      // DISCLAIMER!!!  READ THIS!!!!
		      // The following declare uses a very unsupported feature in REALbasic.  This means a few
		      // things.  1) Don't rely on this call working forever -- it's entirely possible that upgrading to
		      // a new version of REALbasic will cause this declare to break.  2) Don't try to use declares
		      // into the plugins APIs yourself.  And 3) Since this is not a supported way to use declares, you
		      // get everything you deserve if you use a declare like this.  You will eventually come down
		      // with the plague and someone will drink all your beer.  Basically... don't use this sort of thing!
		      // The declare will be removed as soon as there is a sanctioned way to get this functionality
		      // from REALbasic.  This declare should (hypothetically) work in 5.5 only (tho it may work in
		      // version 5 as well).  Consider all other versions of REALbasic unsupported.
		      Declare Sub drawPicturePrimitive Lib "" ( hdc as Integer, p as Picture, drawRect as Ptr, transparent as Boolean )
		      Declare Function CreateCompatibleDC Lib "Gdi32" ( hdc as Integer ) as Integer
		      Declare Function CreateCompatibleBitmap Lib "Gdi32" ( hdc as Integer, width as Integer, height as Integer ) as Integer
		      Declare Function SelectObject Lib "Gdi32" ( hdc as Integer, hbitmap as Integer ) as Integer
		      Declare Sub DeleteDC Lib "Gdi32" ( hdc as Integer )
		      
		      // First, get the HDC from the graphics object for the picture
		      Dim hPictureDC as Integer
		      #if RBVersion < 2006.01 then
		        hPictureDC = REALGraphicsDC( p.Graphics )
		      #else
		        hPictureDC = p.Graphics.Handle( Graphics.HandleTypeHDC )
		      #endif
		      
		      // Then make a compatible device context
		      Dim hScratchDC as Integer
		      hScratchDC = CreateCompatibleDC( hPictureDC )
		      
		      // Then make the bitmap that we'll be using
		      Dim hBitmap as Integer
		      hBitmap = CreateCompatibleBitmap( hPictureDC, p.Width, p.Height )
		      
		      // Select the bitmap into the context and keep the old
		      // contents around
		      Dim hOld as Integer
		      hOld = SelectObject( hScratchDC, hBitmap )
		      
		      // Now draw the original image to the newly created
		      // bitmap
		      dim drawRect as new MemoryBlock( 8 )
		      drawRect.Short( 4 ) = p.Width
		      drawRect.Short( 6 ) = p.Height
		      drawPicturePrimitive( hScratchDC, p, drawRect, false )
		      
		      // Be sure to select the old bitmap back into place
		      Dim hRet as Integer
		      hRet = SelectObject( hScratchDC, hOld )
		      
		      // And free our memory
		      DeleteDC( hScratchDC )
		      
		      // And return our bitmap
		      return hRet
		    #endif
		  #else
		    // Since REAL Software decided to break Lib "" declares without giving
		    // any replacement whatsoever for this functionality, we have to go the
		    // slow route.  We will save the bitmaps out to disk, and then load them
		    // back up so we can get the handle.
		    Declare Function LoadImageW Lib "User32" ( hInst as Integer, name as WString, type as Integer, cx as Integer, cy as Integer, load as Integer ) as Integer
		    
		    dim main as FolderItem = SpecialFolder.Temporary.Child( "main" )
		    p.Save( main, Picture.SaveAsWindowsBMP )
		    Const IMAGE_BITMAP = 0
		    Const LR_LOADFROMFILE = &H10
		    return LoadImageW( 0, main.AbsolutePath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE )
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IntegerToMemoryBlock(i as Integer) As MemoryBlock
		  dim mb as new MemoryBlock( 4 )
		  mb.Long( 0 ) = i
		  return mb.Ptr( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MemoryBlockToInteger(mb as MemoryBlock) As Integer
		  dim ret as new MemoryBlock( 4 )
		  ret.Ptr( 0 ) = mb
		  return ret.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RemoveAllItems()
		  #if TargetWin32
		    Dim TVM_DELETEITEM as Integer = TV_FIRST + 1
		    Const TVI_ROOT                = &hFFFF0000
		    Call SendMessage( TVM_DELETEITEM, 0, TVI_ROOT )
		  #endif
		  mItems = nil
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RemoveItem(item as TreeViewItem)
		  #if TargetWin32
		    Dim TVM_DELETEITEM as Integer = TV_FIRST + 1
		    
		    Call SendMessage( TVM_DELETEITEM, 0, item.Handle )
		    
		    try
		      mItems.Remove( item.Handle )
		    catch
		    end
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SelectItem(item as TreeViewItem)
		  #if TargetWin32
		    // We want to change the selection to the
		    // proper item
		    Dim TVM_SELECTITEM as Integer = TV_FIRST + 11
		    Const TVGN_CARET = &h0009
		    Const TVGN_DROPHILITE = &h0008
		    
		    // First set the caret
		    Call SendMessage( TVM_SELECTITEM, TVGN_CARET, item.Handle )
		    
		    // Then set the highlight
		    Call SendMessage( TVM_SELECTITEM, TVGN_DROPHILITE, item.Handle )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SendMessage(msg as Integer, wParam as Integer, lParam as Integer) As Integer
		  #if TargetWin32
		    Soft Declare Function SendMessageA Lib "User32" ( wnd as Integer, msg as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    Soft Declare Function SendMessageW Lib "User32" ( wnd as Integer, msg as Integer, _
		    wParam as Integer, lParam as Integer ) as Integer
		    
		    if mUnicodeSavvy then
		      return SendMessageW( mWnd, msg, wParam, lParam )
		    else
		      return SendMessageA( mWnd, msg, wParam, lParam )
		    end if
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function WndProc(hWnd as Integer, msg as Integer, wParam as Integer, lParam as Integer, ByRef returnValue as Integer) As Boolean
		  Const WM_NOTIFY = &h004E
		  Const WM_MOUSEMOVE = &h0200
		  Const WM_LBUTTONUP = &h0202
		  
		  select case msg
		  case WM_NOTIFY
		    // We got a notification of some sort, so let's try to
		    // handle it
		    returnValue = HandleWM_Notify( hWnd, wParam, IntegerToMemoryBlock( lParam ) )
		    return true
		  case WM_MOUSEMOVE
		    // Handle the mouse moving
		    returnValue = HandleWM_MouseMove( lParam )
		    return true
		  case WM_LBUTTONUP
		    // Handle the mouse up
		    returnValue = HandleWM_LButtonUp
		    return true
		  end select
		  
		  return false
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event ItemSelected(item as TreeViewItem)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ItemUnselected(item as TreeViewItem)
	#tag EndHook


	#tag Property, Flags = &h21
		Private mDragging As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mImageList As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mItems As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mRBWindow As Window
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mUnicodeSavvy As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mWnd As Integer
	#tag EndProperty


	#tag Constant, Name = TVN_FIRST, Type = Integer, Dynamic = False, Default = \"-400", Scope = Private
	#tag EndConstant

	#tag Constant, Name = TV_FIRST, Type = Integer, Dynamic = False, Default = \"&h1100", Scope = Private
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
End Class
#tag EndClass
