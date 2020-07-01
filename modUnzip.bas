Attribute VB_Name = "modUnzip"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' argv
Private Type UNZIPnames
  s(0 To 1023) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
  ch(0 To 32800) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
  ch(0 To 255) As Byte
End Type

' DCL structure
Public Type DCLIST
  ExtractOnlyNewer As Long      ' 1 to extract only newer
  SpaceToUnderScore As Long     ' 1 to convert spaces to underscore
  PromptToOverwrite As Long     ' 1 if overwriting prompts required
  fQuiet As Long                ' 0 = all messages, 1 = few messages, 2 = no messages
  ncflag As Long                ' write to stdout if 1
  ntflag As Long                ' test zip file
  nvflag As Long                ' verbose listing
  nUflag As Long                ' "update" (extract only newer/new files)
  nzflag As Long                ' display zip file comment
  ndflag As Long                ' all args are files/dir to be extracted
  noflag As Long                ' 1 if always overwrite files
  naflag As Long                ' 1 to do end-of-line translation
  nZIflag As Long               ' 1 to get zip info
  C_flag As Long                ' 1 to be case insensitive
  fPrivilege As Long            ' zip file name
  lpszZipFN As String           ' directory to extract to.
  lpszExtractDir As String
End Type

Private Type USERFUNCTION
  ' Callbacks:
  lptrPrnt As Long           ' Pointer to application's print routine
  lptrSound As Long          ' Pointer to application's sound routine.  NULL if app doesn't use sound
  lptrReplace As Long        ' Pointer to application's replace routine.
  lptrPassword As Long       ' Pointer to application's password routine.
  lptrMessage As Long        ' Pointer to application's routine for
                             ' displaying information about specific files in the archive
                             ' used for listing the contents of the archive.
  lptrService As Long        ' callback function designed to be used for allowing the
                             ' app to process Windows messages, or cancelling the operation
                             ' as well as giving option of progress.  If this function returns
                             ' non-zero, it will terminate what it is doing.  It provides the app
                             ' with the name of the archive member it has just processed, as well
                             ' as the original size.
                             
  ' Values filled in after processing:
  lTotalSizeComp As Long     ' Value to be filled in for the compressed total size, excluding
                             ' the archive header and central directory list.
  lTotalSize As Long         ' Total size of all files in the archive
  lCompFactor As Long        ' Overall archive compression factor
  lNumMembers As Long        ' Total number of files in the archive
  cchComment As Integer      ' Flag indicating whether comment in archive.
End Type

Public Type ZIPVERSIONTYPE
  major As Byte
  minor As Byte
  patchlevel As Byte
  not_used As Byte
End Type

Public Type UZPVER
  structlen As Long         ' Length of structure
  flag As Long              ' 0 is beta, 1 uses zlib
  betalevel As String * 10  ' e.g "g BETA"
  date As String * 20       ' e.g. "4 Sep 95" (beta) or "4 September 1995"
  zlib As String * 10       ' e.g. "1.0.5 or NULL"
  UNZIP As ZIPVERSIONTYPE
  zipinfo As ZIPVERSIONTYPE
  os2dll As ZIPVERSIONTYPE
  windll As ZIPVERSIONTYPE
End Type

Private Declare Function Wiz_SingleEntryUnzip Lib "vbuzip10.dll" (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, ByVal xfnc As Long, ByRef xfnv As UNZIPnames, dcll As DCLIST, Userf As USERFUNCTION) As Long
Public Declare Sub UzpVersion2 Lib "vbuzip10.dll" (uzpv As UZPVER)

' Object for callbacks:
Private m_UNZIP As UNZIP
Private m_bCancel As Boolean

Private Function plAddressOf(ByVal lPtr As Long) As Long
  ' VB Bug workaround fn
  plAddressOf = lPtr
End Function

Sub saveError(a, b, c)

End Sub

Sub FileZipCopy(Origen As String, Desti As String)
   Dim DestiZip As String, Zp As New UNZIP, P As Integer, Pp As Integer
   
   P = InStr(Desti, ".")
   Pp = Len(Desti)
   While P > 0
      Pp = P
      P = InStr(P + 1, Desti, ".")
   Wend
   
   DestiZip = Left(Desti, Pp) & "ZipTrans"
   P = InStr(DestiZip, "[Contingut#")
   If P > 0 Then DestiZip = Left(DestiZip, P - 1) & "[Zontingut#" & Right(DestiZip, Len(DestiZip) - P - 10)
   
   MyKill DestiZip
   
   Zp.ZipeaFile Origen, DestiZip
   
End Sub



Private Sub UnzipMessageCallBack(ByVal ucsize As Long, ByVal csiz As Long, ByVal cfactor As Integer, ByVal mo As Integer, ByVal dy As Integer, ByVal yr As Integer, ByVal hh As Integer, ByVal mm As Integer, ByVal c As Byte, ByRef fname As CBCh, ByRef meth As CBCh, ByVal crc As Long, ByVal fCrypt As Byte)
  
  Dim sFileName As String
  Dim sFolder As String
  Dim dDate As Date
  Dim sMethod As String
  Dim iPos As Long

  On Error GoTo fin
   
  ' Add to unzip class:
  With m_UNZIP
    ' Parse:
    sFileName = StrConv(fname.ch, vbUnicode)
    ParseFileFolder sFileName, sFolder
    dDate = DateSerial(yr, mo, hh)
    dDate = dDate + TimeSerial(hh, mm, 0)
    sMethod = StrConv(meth.ch, vbUnicode)
    iPos = InStr(sMethod, vbNullChar)
    If (iPos > 1) Then sMethod = Left$(sMethod, iPos - 1)
    Debug.Print fCrypt
    .DirectoryListAddFile sFileName, sFolder, dDate, csiz, crc, ((fCrypt And 64) = 64), cfactor, sMethod
  End With
  
  Exit Sub
  
fin:
  saveError "UNZIP", Err.Number, Err.Description
   
End Sub

Private Function UnzipPrintCallback(ByRef fname As CBChar, ByVal x As Long) As Long
  
  Dim iPos As Long
  Dim sFIle As String
  On Error GoTo fin
  
  ' Check we've got a message:
  If x > 1 And x < 1024 Then
    ' If so, then get the readable portion of it:
    ReDim b(0 To x) As Byte
    CopyMemory b(0), fname, x
    ' Convert to VB string:
    sFIle = StrConv(b, vbUnicode)
    
    ' Fix up backslashes:
    ReplaceSection sFIle, "/", "\"
    
    ' Tell the caller about it
    m_UNZIP.ProgressReport sFIle
  End If
  UnzipPrintCallback = 0
  
  Exit Function
  
fin:
  saveError "UNZIP", Err.Number, Err.Description


End Function

Private Function UnzipPasswordCallBack(ByRef pwd As CBCh, ByVal x As Long, ByRef s2 As CBCh, ByRef Name As CBCh) As Long

  Dim bCancel As Boolean
  Dim sPassword As String
  Dim b() As Byte
  Dim lSize As Long

  On Error GoTo fin
  
  ' The default:
  UnzipPasswordCallBack = 1
   
  If m_bCancel Then Exit Function
  
  ' Ask for password:
  m_UNZIP.PasswordRequest sPassword, bCancel
     
  sPassword = Trim$(sPassword)
  
  ' Cancel out if no useful password:
  If bCancel Or Len(sPassword) = 0 Then
    m_bCancel = True
    Exit Function
  End If
  
  ' Put password into return parameter:
  lSize = Len(sPassword)
  If lSize > 254 Then
    lSize = 254
  End If
  b = StrConv(sPassword, vbFromUnicode)
  CopyMemory pwd.ch(0), b(0), lSize
  
  ' Ask UnZip to process it:
  UnzipPasswordCallBack = 0
  
    Exit Function
  
fin:
  saveError "UNZIP", Err.Number, Err.Description
       
End Function

Private Function UnzipReplaceCallback(ByRef fname As CBChar) As Long
  
  Dim eResponse As EUZOverWriteResponse
  Dim iPos As Long
  Dim sFIle As String

  On Error GoTo fin
  eResponse = euzDoNotOverwrite
  
  ' Extract the filename:
  sFIle = StrConv(fname.ch, vbUnicode)
  iPos = InStr(sFIle, vbNullChar)
  If (iPos > 1) Then sFIle = Left$(sFIle, iPos - 1)

  ' No backslashes:
  ReplaceSection sFIle, "/", "\"
  
  ' Request the overwrite request:
  m_UNZIP.OverwriteRequest sFIle, eResponse
  
  ' Return it to the zipping lib
  UnzipReplaceCallback = eResponse
  
  Exit Function
  
fin:
  saveError "UNZIP", Err.Number, Err.Description

End Function

Private Function UnZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long
  
  Dim iPos As Long
  Dim sInfo As String
  Dim bCancel As Boolean
    
  '-- Always Put This In Callback Routines!
  On Error GoTo fin
    
  ' Check we've got a message:
  If x > 1 And x < 1024 Then
    ' If so, then get the readable portion of it:
    ReDim b(0 To x) As Byte
    CopyMemory b(0), mname, x
    ' Convert to VB string:
    sInfo = StrConv(b, vbUnicode)
    iPos = InStr(sInfo, vbNullChar)
    If iPos > 0 Then
      sInfo = Left$(sInfo, iPos - 1)
    End If
    ReplaceSection sInfo, "\", "/"
    m_UNZIP.Service sInfo, bCancel
    If bCancel Then
      UnZipServiceCallback = 1
    Else
      UnZipServiceCallback = 0
    End If
  End If

  Exit Function
  
fin:
  saveError "UNZIP", Err.Number, Err.Description

End Function

Private Sub ParseFileFolder(ByRef sFileName As String, ByRef sFolder As String)

  Dim iPos As Long
  Dim iLastPos As Long

  iPos = InStr(sFileName, vbNullChar)
  If (iPos <> 0) Then sFileName = Left$(sFileName, iPos - 1)
  
  iLastPos = ReplaceSection(sFileName, "/", "\")
  
  If (iLastPos > 1) Then
    sFolder = Left$(sFileName, iLastPos - 2)
    sFileName = Mid$(sFileName, iLastPos)
  End If
   
End Sub

Private Function ReplaceSection(ByRef sString As String, ByVal sToReplace As String, ByVal sReplaceWith As String) As Long

  Dim iPos As Long
  Dim iLastPos As Long
   
  iLastPos = 1
  Do
    iPos = InStr(iLastPos, sString, "/")
    If (iPos > 1) Then
      Mid$(sString, iPos, 1) = "\"
      iLastPos = iPos + 1
    End If
  Loop While Not (iPos = 0)
  ReplaceSection = iLastPos

End Function

' Main subroutine
Public Function VBUnzip(UNZIPObject As UNZIP, tDCL As DCLIST, iIncCount As Long, sInc() As String, iExCount As Long, sExc() As String) As Long
  
  Dim tUser As USERFUNCTION
  Dim lR As Long
  Dim tInc As UNZIPnames
  Dim tExc As UNZIPnames
  Dim i As Long

  On Error GoTo ErrorHandler

  Set m_UNZIP = UNZIPObject
  ' Set Callback addresses
  tUser.lptrPrnt = plAddressOf(AddressOf UnzipPrintCallback)
  tUser.lptrSound = 0& ' not supported
  tUser.lptrReplace = plAddressOf(AddressOf UnzipReplaceCallback)
  tUser.lptrPassword = plAddressOf(AddressOf UnzipPasswordCallBack)
  tUser.lptrMessage = plAddressOf(AddressOf UnzipMessageCallBack)
  tUser.lptrService = plAddressOf(AddressOf UnZipServiceCallback)
       
  ' Set files to include/exclude:
  If (iIncCount > 0) Then
    For i = 1 To iIncCount
      tInc.s(i - 1) = sInc(i)
    Next i
    tInc.s(iIncCount) = vbNullChar
  Else
    tInc.s(0) = vbNullChar
  End If
  If (iExCount > 0) Then
    For i = 1 To iExCount
      tExc.s(i - 1) = sExc(i)
    Next i
    tExc.s(iExCount) = vbNullChar
  Else
    tExc.s(0) = vbNullChar
  End If
  m_bCancel = False
  VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)
   
  Exit Function
  
ErrorHandler:
  Dim lErr As Long
  Dim sErr As String
  lErr = Err.Number
  sErr = Err.Description
  VBUnzip = -1
  Set m_UNZIP = Nothing
'  Err.Raise lErr, App.EXEName & ".VBUnzip", sErr
  saveError "UNZIP", Err.Number, Err.Description

End Function
