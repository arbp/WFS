#tag Module
Protected Module Cryptography
	#tag Method, Flags = &h21
		Private Sub AppendDataChunkToBuffer(dataChunk as MemoryBlock, size as Integer)
		  if dataChunk = nil then return
		  
		  dim replacementBuffer as MemoryBlock
		  dim startLoc as Integer
		  
		  if mBuffer = nil then
		    // First, we may need to allocate the buffer
		    replacementBuffer = new MemoryBlock( size )
		    
		    startLoc = 0
		  else
		    // Otherwise we need to make a new buffer
		    // that is large enough to hold all the data
		    replacementBuffer = new MemoryBlock( size + mBuffer.Size )
		    
		    // And copy over the old data
		    replacementBuffer.StringValue( 0, mBuffer.Size ) = mBuffer
		    
		    startLoc = mBuffer.Size
		  end
		  
		  // And append on our new data
		  replacementBuffer.StringValue( startLoc, size ) = dataChunk.StringValue( 0, size )
		  
		  // And set our new buffer
		  mBuffer = replacementBuffer
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub BeginDecryption(password as String)
		  // This function does the same stuff as
		  // BeginEncryption does, so we'll just call
		  // thru to that
		  BeginEncryption( password )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub BeginEncryption(password as String)
		  #if TargetWin32
		    // We want to start encrypting data.  In order to
		    // do this, we need to hash out the password
		    
		    // Check to make sure we have a provider
		    if mProvider = 0 then
		      MsgBox "No provider"
		      return
		    end
		    
		    // Check to see if we already have a hash password,
		    // and if we do, it's an error
		    if mPasswordHashHandle <> 0 then
		      MsgBox "Already have a hash password!"
		      return
		    end
		    
		    // Hash out the new password
		    Call HashData( mProvider, password, mPasswordHashHandle )
		    
		    // Now that we have the password hashed, let's
		    // make our encryption key
		    Declare Function CryptDeriveKey Lib "AdvApi32" ( provider as Integer, _
		    cipher as Integer, hash as Integer, export as Integer, _
		    ByRef key as Integer ) as Integer
		    dim key, ret as Integer
		    Const CRYPT_EXPORTABLE = &h1
		    
		    ret = CryptDeriveKey( mProvider, mCryptoAlgorithm, mPasswordHashHandle, _
		    CRYPT_EXPORTABLE, mEncryptionKey )
		    
		    Declare Function GetLastError Lib "Kernel32" () as Integer
		    if ret = 0 then
		      MsgBox "Couldn't derive a key: 0x" + Hex( GetLastError )
		      return
		    end
		    
		    // Reset our encryption buffer
		    mBuffer = nil
		    
		    dim blockLen as Integer
		    // Now we want to get the hash value back
		    Declare Function CryptGetKeyParam Lib "AdvApi32" ( _
		    key as Integer, type as Integer, value as Ptr, _
		    ByRef length as Integer, flags as Integer ) as Integer
		    
		    dim size as Integer = 4
		    dim toss as new MemoryBlock( 4 )
		    Const KP_BLOCKLEN = 8
		    ret = CryptGetKeyParam( mEncryptionKey, KP_BLOCKLEN, toss, _
		    size, 0 )
		    
		    blockLen = toss.Long( 0 )
		    
		    mBlockLength = blockLen
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DecryptChunk(data as String)
		  #if TargetWin32
		    // Make sure that we have an decryption key
		    if mEncryptionKey = 0 then
		      MsgBox "No decryption key"
		      return
		    end
		    
		    // Now let's decrypt the data
		    Declare Function CryptDecrypt Lib "AdvApi32" ( key as Integer, _
		    hash as Integer, final as Boolean, flags as Integer, _
		    data as Integer, ByRef dataLength as Integer ) as Integer
		    
		    dim done as Boolean
		    done = (data = "")
		    dim ret as Integer
		    
		    // Stuff it into a memory block
		    dim dataLength as Integer = LenB( data )
		    
		    // Now we make a buffer large enough to hold the
		    // encrypted data
		    dim encryptedData as Integer
		    dim mb as MemoryBlock
		    if dataLength > 0 then
		      mb = new MemoryBlock( dataLength )
		      mb = data
		      encryptedData = MakeMemoryBlockAsInteger( mb )
		    end
		    
		    ret = CryptDecrypt( mEncryptionKey, 0, done, 0, encryptedData, dataLength )
		    
		    if ret = 0 and not done then
		      MsgBox "Could not decrypt data"
		      return
		    end
		    
		    // And add the new chunk of encrypted data to
		    // our ongoing data
		    AppendDataChunkToBuffer( MemoryBlockFromInteger( encryptedData ), dataLength )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DestroyHashAndKey()
		  #if TargetWin32
		    Dim ret as Integer
		    // Free our hash handle
		    Declare Function CryptDestroyHash Lib "AdvApi32" ( _
		    hashHandle as Integer ) as Integer
		    
		    ret = CryptDestroyHash( mPasswordHashHandle )
		    
		    if ret = 0 then
		      MsgBox "Couldn't destroy the password hash"
		      return
		    end
		    
		    // And free our key as well
		    Declare Function CryptDestroyKey Lib "AdvApi32" ( _
		    key as Integer ) as Integer
		    
		    ret = CryptDestroyKey( mEncryptionKey )
		    if ret = 0 then
		      MsgBox "Couldn't destroy the key"
		      return
		    end
		    
		    mPasswordHashHandle = 0
		    mEncryptionKey = 0
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub EncryptChunk(data as String)
		  #if TargetWin32
		    dim done as Boolean
		    done = (data = "")
		    
		    // Make sure that we have an encryption key
		    if mEncryptionKey = 0 then
		      MsgBox "No encryption key"
		      return
		    end
		    
		    // RB 5.5 and before are horrible about MemoryBlocks that
		    // are nil.  It sucks.  So we need to code for that.  So what we'll do is
		    // declare the function as an Integer and then use a helper function
		    // to convert a MemoryBlock into an integer
		    
		    // We want to encrypt this chunk of data
		    Declare Function CryptEncrypt Lib "AdvApi32" ( _
		    key as Integer, hash as Integer, final as Boolean, _
		    flags as Integer, data as Integer, ByRef dataLen as Integer, _
		    bufLen as Integer ) as Integer
		    
		    // Stuff it into a memory block
		    dim dataLength as Integer = LenB( data )
		    dim ret as Integer
		    
		    // Now we make a buffer large enough to hold the
		    // encrypted data
		    dim encryptedData as Integer
		    dim mb as MemoryBlock
		    if dataLength > 0 then
		      mb = new MemoryBlock( dataLength )
		      mb = data
		      encryptedData = MakeMemoryBlockAsInteger( mb )
		    end
		    
		    // And encrypt
		    ret = CryptEncrypt( mEncryptionKey, 0, done, 0, encryptedData, dataLength, dataLength )
		    
		    Declare Function GetLastError Lib "Kernel32" () as Integer
		    
		    if ret = 0 then
		      MsgBox "Couldn't encrypt the data: 0x " + Hex( GetLastError )
		      return
		    end
		    
		    // And add the new chunk of encrypted data to
		    // our ongoing data
		    AppendDataChunkToBuffer( MemoryBlockFromInteger( encryptedData ), dataLength )
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub EndDecryption()
		  // We want to finish the decryption process
		  DecryptChunk( "" )
		  
		  DestroyHashAndKey
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub EndEncryption()
		  // We need to do one last encryption where
		  // we pass in no data to signify that we're done
		  EncryptChunk( "" )
		  
		  DestroyHashAndKey
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Finalize()
		  #if TargetWin32
		    Declare Function CryptReleaseContext Lib "AdvApi32" ( _
		    provider as Integer, flags as Integer ) as Integer
		    
		    dim ret as Integer
		    ret = CryptReleaseContext( mProvider, 0 )
		    
		    if ret = 0 then
		      MsgBox "Could not release the provider context"
		    end
		    
		    mProvider = 0
		  #endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetDecryptedData() As String
		  return mBuffer
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetEncryptedData() As MemoryBlock
		  return mBuffer
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HashData(provider as Integer, data as String, ByRef handle as Integer) As MemoryBlock
		  #if TargetWin32
		    // Now we need a hash of the data
		    Declare Function CryptCreateHash Lib "AdvApi32" ( _
		    provider as Integer, algorithm as Integer, key as Integer, _
		    flags as Integer, ByRef hashHandle as Integer ) as Integer
		    
		    'Const ALG_CLASS_HASH = (4 << 13)
		    'Const ALG_TYPE_ANY = 0
		    'Const ALG_SID_MD5 = 3
		    Const CALG_MD5 = 32771'        (ALG_CLASS_HASH | ALG_TYPE_ANY | ALG_SID_MD5)
		    Const HP_HASHVAL = &h0002  // Hash value
		    Const HP_HASHSIZE = &h0004  // Hash value size
		    
		    // Get the hash handle
		    dim hashHandle, ret as Integer
		    ret = CryptCreateHash( provider, mHashAlgorithm, 0, 0, hashHandle )
		    if ret = 0 then
		      MsgBox "Couldn't create the hash"
		      return nil
		    end
		    
		    // And actually hash the data
		    Declare Function CryptHashData Lib "AdvApi32" ( _
		    hashHandle as Integer, data as Ptr, length as Integer, _
		    flags as Integer ) as Integer
		    
		    dim dataPtr as new MemoryBlock( Len( data ) )
		    dataPtr = data
		    
		    ret = CryptHashData( hashHandle, dataPtr, dataPtr.Size, 0 )
		    
		    if ret = 0 then
		      MsgBox "Couldn't hash data"
		      return nil
		    end
		    
		    // Now we want to get the hash value back
		    Declare Function CryptGetHashParam Lib "AdvApi32" ( _
		    hashHandle as Integer, type as Integer, value as Ptr, _
		    ByRef length as Integer, flags as Integer ) as Integer
		    
		    dim size as Integer = 4
		    dim toss as new MemoryBlock( 4 )
		    ret = CryptGetHashParam( hashHandle, HP_HASHSIZE, toss, _
		    size, 0 )
		    
		    if ret = 0 then
		      MsgBox "Couldn't get the hash value size"
		      return nil
		    end
		    
		    size = toss.Long( 0 )
		    
		    // Allocate a buffer to hold the data
		    dim hashValue as new MemoryBlock( size )
		    
		    // And get the actual hash value
		    ret = CryptGetHashParam( hashHandle, HP_HASHVAL, hashValue, _
		    size, 0 )
		    
		    if ret = 0 then
		      MsgBox "Couldn't get the hash value itself"
		      return nil
		    end
		    
		    handle = hashHandle
		    
		    return hashValue
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HashData(data as String) As MemoryBlock
		  dim toss as Integer
		  dim ret as MemoryBlock
		  
		  ret = HashData( mProvider, data, toss )
		  
		  // Free our hash handle
		  Declare Sub CryptDestroyHash Lib "AdvApi32" ( _
		  hashHandle as Integer )
		  
		  CryptDestroyHash( toss )
		  
		  return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Initialize() As Boolean
		  #if TargetWin32
		    Soft Declare Function CryptAcquireContextA Lib "AdvApi32" ( _
		    ByRef provider as Integer, container as Integer, _
		    providerName as CString, providerType as Integer, _
		    flags as Integer ) as Integer
		    Soft Declare Function CryptAcquireContextW Lib "AdvApi32" ( _
		    ByRef provider as Integer, container as Integer, _
		    providerName as WString, providerType as Integer, _
		    flags as Integer ) as Integer
		    
		    // First, get the provider
		    dim ret as Integer
		    Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
		    Const PROV_RSA_FULL = 1
		    Const CRYPT_NEWKEYSET = &h00000008
		    
		    if System.IsFunctionAvailable( "CryptAcquireContextW", "AdvApi32" ) then
		      ret = CryptAcquireContextW( mProvider, 0, MS_DEF_PROV, PROV_RSA_FULL, 0 )
		    else
		      ret = CryptAcquireContextA( mProvider, 0, MS_DEF_PROV, PROV_RSA_FULL, 0 )
		    end if
		    
		    if ret = 0 then
		      // Couldn't acquire the context, so try again with a new keyset
		      if System.IsFunctionAvailable( "CryptAcquireContextW", "AdvApi32" ) then
		        ret = CryptAcquireContextW( mProvider, 0, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_NEWKEYSET )
		      else
		        ret = CryptAcquireContextA( mProvider, 0, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_NEWKEYSET )
		      end if
		      
		      if ret = 0 then
		        MsgBox "0x" + Hex( Win32DeclareLibrary.GetLastError )
		        // We really can't do anything right, can we?
		        return false
		      end if
		    end
		    
		    mCryptoAlgorithm = kCryptoTypeRC4
		    mHashAlgorithm = kHashTypeMD5
		    
		    return true
		  #endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MakeMemoryBlockAsInteger(data as MemoryBlock) As Integer
		  dim ret as new MemoryBlock( 4 )
		  ret.Ptr( 0 ) = data
		  
		  return ret.Long( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MemoryBlockFromInteger(mb as Integer) As MemoryBlock
		  dim ret as new MemoryBlock( 4 )
		  ret.Long( 0 ) = mb
		  
		  return ret.Ptr( 0 )
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub UseCryptoAlgorithm(algo as Integer)
		  mCryptoAlgorithm = algo
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub UseHashAlgorithm(algo as Integer)
		  mHashAlgorithm = algo
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mBlockLength As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mBuffer As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCryptoAlgorithm As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mEncryptionKey As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHashAlgorithm As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPasswordHashHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mProvider As Integer
	#tag EndProperty


	#tag Constant, Name = kCryptoTypeRC4, Type = Integer, Dynamic = False, Default = \"26625", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeHMAC, Type = Integer, Dynamic = False, Default = \"32777", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeMAC, Type = Integer, Dynamic = False, Default = \"32773", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeMD2, Type = Integer, Dynamic = False, Default = \"32769", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeMD4, Type = Integer, Dynamic = False, Default = \"32770", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeMD5, Type = Integer, Dynamic = False, Default = \"32771", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeSHA, Type = Integer, Dynamic = False, Default = \"32772", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = kHashTypeSHAMD5, Type = Integer, Dynamic = False, Default = \"32776", Scope = Protected
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
