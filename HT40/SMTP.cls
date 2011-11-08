VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "SMTP"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

' maximums per RFC 821...
Private Const MAX_TEXTLINE_LEN = 1000               ' maximum total lenght of a text line
Private Const MAX_RECIPIENTS = 100                  ' maximum number of recipients that must be buffered

' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

' Class Enum for email encoding method
Public Enum ENCODE_METHOD
    MIME_ENCODE = 0
    UU_ENCODE = 1
End Enum

Private pbBase64Byt(0 To 63)    As Byte             ' base 64 encoder byte array
Private psUUEncodeChr(0 To 63)  As String           ' UU encoder string array

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Sub Class_Initialize()
  Dim iPtr
    ' initialize the arrays for base64 & uu encoders
    For iPtr = 0 To 63
        pbBase64Byt(iPtr) = Asc(Mid$(BASE64CHR, iPtr + 1, 1))
        psUUEncodeChr(iPtr) = Chr$(iPtr + &H20)
    Next
    psUUEncodeChr(0) = Chr$(&H60)

End Sub

Rem Body : String
Rem Charset : String
Rem ContentTransferEncoding : String
Rem ContentType : String
Rem DeferredDelivery : Date
Rem Encoding : String
Rem ErrorCode : Integer
Rem ErrorMessage : String
Rem ErrorSource : String
Rem ISOEncodeHeaders : Boolean
Rem Lazysend : Boolean
Rem Log : String
Rem Logging : Boolean
Rem MailDomain : String
Rem MimeVersion : String
Rem Priority : Integer
Rem Recipients : String
Rem ReplyTo : String
Rem ReturnReceipt : Boolean
Rem Sender : String
Rem SenderName : String
Rem ServerAddress : String
Rem Silent : Boolean
Rem SimpleLayout : Boolean
Rem Subject : String
Rem UsePipelining : Boolean
Rem AddAttachment( FileName, [ContentType] )
Rem AddCustomAttachment( FileName, Data )
Rem AddHeader( XHeader, Value )
Rem AddNativeHeader( Header, Value )
Rem AddRecipient( Email )
Rem AddRecipientBCC( Email )
Rem AddRecipientCC( Email )
Rem AddRecipientEx( Email, Name )
Rem AddURLAttachment( bstrURL, bstrAttachAs, [bstrAuth] )
Rem AppendBodyFromFile( FileName )
Rem AppendText( Text )
Rem ClearAttachments()
Rem ClearCustomHeaders()
Rem ClearRecipients()
Rem Close()
Rem Execute()
Rem ExtractEmailAddressesFromURL( bstrURL, [bstrAuth] )
Rem GetMessageBodyFromURL( bstrURL, [bstrAuth] )
Rem LogCustomMessage( Message )

Private Function EncodeBase64Byte(InArray() As Byte) As Byte()

'******************************************************************************
'
' Synopsis:     Base 64 encode a byte array
'
' Parameters:   InArray  - The input byte array
'
' Return:       encoded byte array
'
' Description:
'   Convert a byte array to a Base 64 encoded byte array. Coerce 3 bytes into
'   4 by converting 3, 8 bit bytes into 4, 6 bit values. Each 6 bit value
'   (0 to 63) is then used as a pointer into a base64 byte array to derive a
'   character.
'
'******************************************************************************

Dim lInPtr              As Long         ' pointer into input array
Dim lOutPtr             As Long         ' pointer into output array
Dim OutArray()          As Byte         ' output byte array buffer
Dim lLen                As Long         ' number of extra bytes past 3 byte boundry
Dim iNewLine            As Long         ' line counter

' if size of input array is not a multiple of 3,
' increase it to the next multiple of 3
lLen = (UBound(InArray) - LBound(InArray) + 1) Mod 3
If lLen Then
    lLen = 3 - lLen
    ReDim Preserve InArray(UBound(InArray) + lLen)
End If

' create an output buffer
ReDim OutArray(UBound(InArray) * 2 + 100)

' step through the input array, 3 bytes at a time
For lInPtr = 0 To UBound(InArray) Step 3
    
    ' add CrLf as required
    If iNewLine = 19 Then
        OutArray(lOutPtr) = 13
        OutArray(lOutPtr + 1) = 10
        lOutPtr = lOutPtr + 2
        iNewLine = 0
    End If
    
    ' convert 3 bytes into 4 base 64 encoded bytes
    OutArray(lOutPtr) = pbBase64Byt((InArray(lInPtr) And &HFC) \ 4)
    OutArray(lOutPtr + 1) = pbBase64Byt((InArray(lInPtr) And &H3) * &H10 + (InArray(lInPtr + 1) And &HF0) \ &H10)
    OutArray(lOutPtr + 2) = pbBase64Byt((InArray(lInPtr + 1) And &HF) * 4 + (InArray(lInPtr + 2) And &HC0) \ &H40)
    OutArray(lOutPtr + 3) = pbBase64Byt(InArray(lInPtr + 2) And &H3F)
    
    ' update pointers
    lOutPtr = lOutPtr + 4
    iNewLine = iNewLine + 1
Next

' add terminator '=' as required
Select Case lLen
    Case 1: OutArray(lOutPtr - 1) = 61
    Case 2: OutArray(lOutPtr - 1) = 61: OutArray(lOutPtr - 2) = 61
End Select

' add CrLf if not already there
If OutArray(lOutPtr - 2) <> 13 Then
    OutArray(lOutPtr) = 13
    OutArray(lOutPtr + 1) = 10
    lOutPtr = lOutPtr + 2
End If

' resize output buffer and return
ReDim Preserve OutArray(lOutPtr - 1)
EncodeBase64Byte = OutArray

End Function

Private Function EncodeBase64String(ByRef str2Encode As String) As String

' ******************************************************************************
'
' Synopsis:     Base 64 encode a string
'
' Parameters:   str2Encode  - The input string
'
' Return:       encoded string
'
' Description:
' Convert a string to a byte array and pass to EncodeBase64Byte function (above)
' for Base64 conversion. Convert byte array back to a string and return.
'
' ******************************************************************************

    Dim tmpByte()   As Byte

    If Len(str2Encode) Then
    
        ' convert string to byte array
        tmpByte = StrConv(str2Encode, vbFromUnicode)

        ' pass to the byte array encoder
        tmpByte = EncodeBase64Byte(tmpByte)

        ' convert back to string & return
        EncodeBase64String = StrConv(tmpByte, vbUnicode)

    End If
    
End Function

Private Function DecodeBase64String(ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim Bits(1 To 4)    As Byte
    Dim strDecode       As String

    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: Bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
        Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
        Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            strDecode = strDecode & Chr$(Bits(iCtr))
        Next

    Next

    DecodeBase64String = strDecode

End Function

Private Function FormatMail(ByVal strIn As String) As String

' ******************************************************************************
'
' Synopsis:     Re-format text lines per RFC 821
'
' Parameters:   strIn   - The input string to be formated
'
' Return:       re-formated string
'
' Description:
' RFC 821 places the following restrictions on user text:
' 1) Before sending a line of mail text begining with a '.
' the sender will add an additional '.
'
' 2) The receiver checks each line of mail text, if a line is single '.
' it is the end of the mail message. If the first character is
' a '.' and there are other characters on the line, the first '.
' is deleted.
'
' 3) The maximum line lenght will not exceed 1000 characters
'
' ******************************************************************************

    Dim sTextLine()     As String
    Dim sRemainder      As String
    Dim sNewLine        As String
    Dim sDelimiter      As String
    Dim lPtr            As Long
    Dim lSplit          As Long

    If Len(strIn) = 0 Then Exit Function

    ' Select the correct delimiter character
    If InStr(strIn, vbCrLf) Then
        sDelimiter = vbCrLf
    ElseIf InStr(strIn, vbCr) Then
        sDelimiter = vbCr
    Else
        sDelimiter = vbNullString
    End If

    ' split the text into seperate lines
    sTextLine() = Split(strIn, sDelimiter)

    ' process each line
    For lPtr = 0 To UBound(sTextLine)
        ' check for lines starting with a '.
        ' when found, add a second '.
        If Left$(sTextLine(lPtr), 1) = "." Then sTextLine(lPtr) = "." & sTextLine(lPtr)

        ' check that the line is not too long (account for 2 extra characters - vbCrLf)
        ' break into smaller elements as required
        If Len(sTextLine(lPtr)) > MAX_TEXTLINE_LEN - 2 Then
            sRemainder = sTextLine(lPtr)
            sNewLine = vbNullString
            If sDelimiter = vbNullString Then sDelimiter = vbCrLf
            Do While Len(sRemainder) > MAX_TEXTLINE_LEN - 2
                ' try to split at a space character, if not then split at MAX_TEXTLINE_LEN - 2
                lSplit = InStrRev(sRemainder, " ", MAX_TEXTLINE_LEN - 2)
                If lSplit = 0 Then lSplit = MAX_TEXTLINE_LEN - 2
                ' insert a vbCrLf at the split point
                sNewLine = sNewLine & Mid$(sRemainder, 1, lSplit) & sDelimiter
                sRemainder = Mid$(sRemainder, lSplit + 1)
            Loop
            sTextLine(lPtr) = sNewLine & sRemainder
        End If
    Next

    FormatMail = Join(sTextLine, sDelimiter)

End Function

Private Function IsDottedQuad(ByVal HostString As String) As Boolean

    ' verify that a string is 'xxx.xxx.xxx.xxx' format
    
    Dim sSplit()        As String
    Dim iCtr            As Integer

    ' split at the "."
    sSplit = Split(HostString, ".")

    ' should be 4 elements
    If UBound(sSplit) <> 3 Then Exit Function

    ' check each element
    For iCtr = 0 To 3
        ' should be numeric
        If Not IsNumeric(sSplit(iCtr)) Then Exit Function

        ' range check
        If iCtr = 0 Then
            If Val(sSplit(iCtr)) > 239 Then Exit Function
        Else
            If Val(sSplit(iCtr)) > 255 Then Exit Function
        End If
    Next
    
    IsDottedQuad = True

End Function

Private Function UUEncodeString(ByRef str2UUEncode As String) As String

' ******************************************************************************
'
' Synopsis:     UUEncode a string
'
' Parameters:   str2UUEncode  - The input string
'
' Return:       encoded string
'
' Description:
' UU Encode a string. Coerce 3 bytes into 4 by converting 3, 8 bit bytes into
' 4, 6 bit values. Each 6 bit value (0 to 63) is then used as a pointer into
' the UUEncode string array to derive the correct character. The string will
' be a multiple of 4 bytes in lenght after conversion, padded with '=' as
' required. The line length will be encoded as a leading character
' (same 0 to 63 encoding) in the return string.
'
' ******************************************************************************

    Dim sValue              As String
    Dim lPtr                As Long
    Dim lCtr                As Long
    Dim lLen                As Long
    Dim lLineLen            As Long
    Dim sEncoded            As String
    Dim Bits8(1 To 3)       As Byte
    Dim Bits6(1 To 4)       As Byte

    lLineLen = Len(str2UUEncode)
    ' lines are limited to 63
    If lLineLen > 63 Then Exit Function

    For lCtr = 1 To Len(str2UUEncode) Step 3

        ' Get 3 characters
        sValue = Mid$(str2UUEncode, lCtr, 3)
        lLen = Len(sValue)

        ' Move string data into a byte array, then
        ' swap bits to create 4, 6 bit values (0-63)
        If lLen < 3 Then Erase Bits8
        CopyMemory Bits8(1), ByVal sValue, lLen
        Bits6(1) = (Bits8(1) And &HFC) \ &H4
        Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
        Bits6(3) = (Bits8(2) And &HF) * &H4 + (Bits8(3) And &HC0) \ &H40
        Bits6(4) = Bits8(3) And &H3F

        ' Encode new 4 byte string by selecting a character from
        ' the array. Length is determined by 'lLen' to make sure
        ' the file attachment is the right length
        For lPtr = 1 To lLen + 1
            sEncoded = sEncoded & psUUEncodeChr(Bits6(lPtr))
        Next

    Next

    ' add the line length character
    sEncoded = psUUEncodeChr(lLineLen) & sEncoded

    ' The decoder expects the size to be a multiple of 4 bytes.
    ' Possible sizes for the last packet are: 2, 3 & 4.
    Select Case lLen + 1
        Case 2: sEncoded = sEncoded & "=="          ' send two pad characters
        Case 3: sEncoded = sEncoded & "="           ' send one pad character
    ' no pad characers needed
    End Select

    UUEncodeString = sEncoded

End Function

'--------------------------------------------------------------------------
'
' Author:       Ramon Bosch Smit <ramon@vincle.es>
' DateCreated:  10/09/1999
' Description:  ISO charset decoder function + support functions
' ModuleType:   BAS
'
'--------------------------------------------------------------------------


' MIME supports techiques to allow the encoding of non-ASCII text
' in various portions of a message header in a manner that is unlikely
' to confuse message-handling software that does not support these
' features.
'
' Certain sequences of ordinary printable ASCII characters (known as
' encoded-words= are reserved for use as encoded data. The syntax of
' encoded-words is such that they are unlikely to accidentally appear
' as normal text in message headers. Furthermore, the characters used
' in encoded-words are restricted to those that do not have special
' meaning in the context in which the encoded word appears.
'
' Generally, an encoded-word is a sequence of printable ASCII characters
' that begin with =?, end with ?=, and have two ?'s in between.  It
' specifies a character set and an encoding method and also includes
' the original text encoded as graphic ASCII characters, according to
' the rules for that encoding method
'
' Syntax:
'
' Encoded-word = "=?" charset "?" encoding "?" encoded-text "?="


Public Function DecodeWord _
                        ( _
                        ByVal p_strEncodedExpression As String, _
                        Optional ByRef p_strCharset As String, _
                        Optional ByRef p_strEncoding As String _
                        ) As String
    ' Description:
    '----------------------------------------------------------------------
    ' Decodes a MIME encoded string (following the RFC 2047 specification)
    
    ' Parameters:
    '----------------------------------------------------------------------
    ' In            p_strEncodedExpression  Encoded string
    ' Out           p_strCharset            Original Charset
    '               p_strEncoding           Encoding method used to
    '                                       encode the original text.
    '                                       Legal values for encoding are:
    '                                       "Q" Quoted Printable
    '                                       "B" Base64 (not supported)
    '                                       ""  Means that p_strEncodedExpression
    '                                       wasn't encoded
    ' Return value:
    '-----------------------------------------------------------------------
    ' Decoded string (the original text in the designated character set)
    
    Dim v_lngLength             As Long
    Dim v_lngIndex              As Long
    
    Dim v_strCurrentChar        As String
    Dim v_strCharset            As String
    Dim v_strEncodedChar        As String
    Dim v_strEncoding           As String
    Dim v_strDecoded            As String
    
    Dim v_blnDecoding           As Boolean
    Dim v_blnDecodingChar       As Boolean
    Dim v_blnFetchingCharset    As Boolean
    Dim v_blnFetchingEncoding   As Boolean
    Dim v_blnFetchingChar       As Boolean
    
    On Error GoTo ErrorHandler
    
    v_lngLength = Len(p_strEncodedExpression)
    
    For v_lngIndex = 1 To v_lngLength
        
        v_strCurrentChar = Mid$(p_strEncodedExpression, v_lngIndex, 1)
        
        Select Case v_strCurrentChar
        
        Case "="
                        
            If UCase$(Mid$(p_strEncodedExpression, v_lngIndex + 1, Len("?ISO"))) = "?ISO" Then
                
                v_blnDecoding = True
                v_blnFetchingCharset = True
                v_blnFetchingEncoding = False
                v_blnDecodingChar = False
                v_blnFetchingChar = False
                v_lngIndex = v_lngIndex + 1
            
            ElseIf v_blnDecodingChar Then
                
                v_blnFetchingChar = True
            
            End If
            
        Case "?"
            
            If v_blnFetchingCharset Then
                
                v_blnFetchingCharset = False
                v_blnFetchingEncoding = True
            
            ElseIf v_blnFetchingEncoding Then
                
                v_blnFetchingEncoding = False
                v_blnDecodingChar = True
            
            ElseIf Mid$(p_strEncodedExpression, v_lngIndex + 1, 1) = "=" Then
                
                v_blnDecoding = False
                v_blnDecodingChar = False
                v_blnFetchingCharset = False
                v_blnFetchingEncoding = False
            
            Else
                
                v_strDecoded = v_strDecoded & v_strCurrentChar
                
            End If
        
        Case "_"
            
            If v_blnDecoding Then
                
                v_strDecoded = v_strDecoded & " "
            
            End If
            
        Case Else
            
            If v_blnFetchingCharset Then
                
                v_strCharset = v_strCharset & v_strCurrentChar
            
            ElseIf v_blnFetchingEncoding Then
                
                v_strEncoding = v_strEncoding & v_strCurrentChar
            
            ElseIf v_blnFetchingChar Then
                
                v_strEncodedChar = v_strEncodedChar & v_strCurrentChar
                
                If Len(v_strEncodedChar) = 2 Then
                    
                    v_blnFetchingChar = False
                    v_strDecoded = v_strDecoded & DecodeChar(v_strEncodedChar, v_strCharset)
                    v_strEncodedChar = vbNullString
                
                End If
             
            Else
                
                v_strDecoded = v_strDecoded & v_strCurrentChar
            
            End If
                    
        End Select
    
    Next v_lngIndex
    
    DecodeWord = v_strDecoded
    p_strCharset = v_strCharset
    p_strEncoding = v_strEncoding
    
ExitHere:
    Exit Function
ErrorHandler:
    Debug.Assert 0  ' Soft stop
    ' TODO: Implement your error hanling/logging here
    Resume ExitHere
End Function

Private Function DecodeChar _
                        ( _
                        ByVal p_strEncodedChar As String, _
                        ByVal p_strCharset As String _
                        ) As String

    ' Description:
    '----------------------------------------------------------------------
    ' Decodes an expression that represents an encoded character to it's
    ' original charset (p_strCharset). In order to do so, p_strEncodedChar,
    ' which is an Hexadecimal number expression,  must be converted to it's decimal
    ' value and then converted to a character through a conversion table
    '
    ' Parameters:
    '----------------------------------------------------------------------
    ' In            p_strEncodedChar        Hexadecimal expression
    '               p_strCharset            Original Charset
    '
    ' Out           None
    '
    ' Return value:
    '-----------------------------------------------------------------------
    ' Decoded character
    
    Dim v_astrISOTable() As String
    Dim v_lngISOCode As Long

    On Error GoTo ErrorHandler
    
    ' Convert the encoded Character (2 digit Hexadecimal number) to
    ' it's decimal value
    v_lngISOCode = HexToDec(p_strEncodedChar)
    ' Load the corresponding ISO Conversion table for
    ' the specified Charset
    LoadISOTable p_strCharset, v_astrISOTable
    ' Lookup the corresponding character in the conversion table
    DecodeChar = v_astrISOTable(v_lngISOCode)

ExitHere:
    Exit Function
ErrorHandler:
    Debug.Assert 0  ' Soft stop
    ' TODO: Implement your error hanling/logging here
    Resume ExitHere
End Function

Private Function HexToDec _
                        ( _
                        ByVal p_strHexNumber As String _
                        ) As Double
    
    
    ' Description:
    '----------------------------------------------------------------------
    ' Converts an Hexadecimal number (expression) to it's decimal value
    '
    ' Parameters:
    '----------------------------------------------------------------------
    ' In            p_strHexNumber          Hexadecimal number
    '
    ' Out           None
    '
    ' Return value:
    '-----------------------------------------------------------------------
    ' Decimal number resulting from the conversion of p_strHexNumber
    
    Dim v_lngNumberLength As Long
    Dim v_lngIndex As Long
    Dim v_dblDecimalNumber As Double
    Dim v_strHexadecimalDigit As String
    Dim v_lngDecimalDigit As Long
    
    On Error GoTo ErrorHandler
    
    v_lngNumberLength = Len(p_strHexNumber)
    
    ' If the length of p_strHexNumber is = 0 then the
    ' function shall return 0 (Decimal)
    
    If v_lngNumberLength > 0 Then
    
        For v_lngIndex = 1 To v_lngNumberLength
                    
            v_strHexadecimalDigit = Mid$(p_strHexNumber, v_lngIndex, 1)
            
            If IsNumeric(v_strHexadecimalDigit) Then
                
                v_lngDecimalDigit = CLng(v_strHexadecimalDigit)
            
            Else
                
                If InStr(1, "ABCDEF", v_strHexadecimalDigit, vbTextCompare) > 0 Then
                
                    v_lngDecimalDigit = Asc(UCase$(v_strHexadecimalDigit)) - 55
            
                Else
                        
                    ' If p_strHexNumber is an invalid Hexadecimal number then we'll
                    ' raise an exception
                    
                    VBA.Err.Raise 5 ' Invalid procedure call
                        
                End If
            
            End If
            
            v_dblDecimalNumber = v_dblDecimalNumber + (v_lngDecimalDigit * (16 ^ (v_lngNumberLength - v_lngIndex)))
        
        Next v_lngIndex
    
    End If

    HexToDec = v_dblDecimalNumber

ExitHere:
    Exit Function
ErrorHandler:
    Debug.Assert 0  ' Soft here
    HexToDec = 0
    ' TODO: Implement your error hanling/logging here
    Resume ExitHere
End Function

Private Sub LoadISOTable _
                        ( _
                        ByVal p_strCharset As String, _
                        ByRef p_astrISOTable() As String _
                        )
    
    ' Description:
    '----------------------------------------------------------------------
    ' Loads a Conversion Table into the p_astrISOTable array.
    ' Each index of the array corresponds to the decimal value of an encoded
    ' Character. e.g: Hexadecimal "3D" = Decimal 61. Index 61 of the array
    ' will contain the decoded value of "3D". If "3D" was encoded using
    ' the ISO-8859-1 charset, then we shall load the ISO-8859-1 conversion table
    ' into the array, which'll give us "=" as the decoded value of "3D"
    '
    ' Parameters:
    '----------------------------------------------------------------------
    ' In            p_strCharset            Charset for which the convesrsion
    '                                       table must be loaded
    '               p_astrISOTable          Empty Array
    ' Out           p_astrISOTable          Loaded Array (contains the conversion
    '                                       table for p_strCharset)
    '
    ' Return value:
    '-----------------------------------------------------------------------
    ' None

    On Error GoTo ErrorHandler
    
    ReDim p_astrISOTable(0 To 255)
    
    Select Case UCase$(p_strCharset)
    
    Case "ISO-8859-1"
        
        ' Load the ISO-8859-1 Conversion table
        
        p_astrISOTable(32) = " "    ' white space
        p_astrISOTable(34) = """"   ' quotation mark
        p_astrISOTable(38) = "&"    ' ampersand
        p_astrISOTable(60) = "<"    ' less-than sign
        p_astrISOTable(61) = "="    ' equal sign
        p_astrISOTable(62) = ">"    ' greater-than sign
        p_astrISOTable(63) = "?"    ' question mark
        p_astrISOTable(160) = " "   ' non-breaking space
        p_astrISOTable(161) = "¡"   ' inverted exclamation
        p_astrISOTable(162) = "¢"   ' cent sign
        p_astrISOTable(163) = "£"   ' pound sterling
        p_astrISOTable(164) = "¤"   ' general currency sign
        p_astrISOTable(165) = "¥"   ' yen sign
        p_astrISOTable(166) = "¦"   ' broken vertical bar
        p_astrISOTable(167) = "§"   ' section sign
        p_astrISOTable(168) = "¨"   ' umlaut (dieresis)
        p_astrISOTable(169) = "©"   ' copyright
        p_astrISOTable(170) = "ª"   ' feminine ordinal
        p_astrISOTable(171) = "«"   ' left angle quote, guillemotleft
        p_astrISOTable(172) = "¬"   ' not sign
        p_astrISOTable(173) = "­"   ' soft hyphen
        p_astrISOTable(174) = "®"   ' registered trademark
        p_astrISOTable(175) = "¯"   ' macron accent
        p_astrISOTable(176) = "°"   ' degree sign
        p_astrISOTable(177) = "±"   ' plus or minus
        p_astrISOTable(178) = "²"   ' superscript two
        p_astrISOTable(179) = "³"   ' superscript three
        p_astrISOTable(180) = "´"   ' acute accent
        p_astrISOTable(181) = "µ"   ' micro sign
        p_astrISOTable(182) = "¶"   ' paragraph sign
        p_astrISOTable(183) = "·"   ' middle dot
        p_astrISOTable(184) = "¸"   ' cedilla
        p_astrISOTable(185) = "¹"   ' superscript one
        p_astrISOTable(186) = "º"   ' masculine ordinal
        p_astrISOTable(187) = "»"   ' right angle quote, guillemotright
        p_astrISOTable(188) = "¼"   ' fraction one-fourth
        p_astrISOTable(189) = "½"   ' fraction one-half
        p_astrISOTable(190) = "¾"   ' fraction three-fourths
        p_astrISOTable(191) = "¿"   ' inverted question mark
        p_astrISOTable(192) = "À"   ' capital A, grave accent
        p_astrISOTable(193) = "Á"   ' capital A, acute accent
        p_astrISOTable(194) = "Â"   ' capital A, circumflex accent
        p_astrISOTable(195) = "Ã"   ' capital A, tilde
        p_astrISOTable(196) = "Ä"   ' capital A, dieresis or umlaut mark
        p_astrISOTable(197) = "Å"   ' capital A, ring
        p_astrISOTable(198) = "Æ"   ' capital AE diphthong (ligature)
        p_astrISOTable(199) = "Ç"   ' capital C, cedilla
        p_astrISOTable(200) = "È"   ' capital E, grave accent
        p_astrISOTable(201) = "É"   ' capital E, acute accent
        p_astrISOTable(202) = "Ê"   ' capital E, circumflex accent
        p_astrISOTable(203) = "Ë"   ' capital E, dieresis or umlaut mark
        p_astrISOTable(204) = "Ì"   ' capital I, grave accent
        p_astrISOTable(205) = "Í"   ' capital I, acute accent
        p_astrISOTable(206) = "Î"   ' capital I, circumflex accent
        p_astrISOTable(207) = "Ï"   ' capital I, dieresis or umlaut mark
        p_astrISOTable(208) = "Ð"   ' capital Eth, Icelandic
        p_astrISOTable(209) = "Ñ"   ' capital N, tilde
        p_astrISOTable(210) = "Ò"   ' capital O, grave accent
        p_astrISOTable(211) = "Ó"   ' capital O, acute accent
        p_astrISOTable(212) = "Ô"   ' capital O, circumflex accent
        p_astrISOTable(213) = "Õ"   ' capital O, tilde
        p_astrISOTable(214) = "Ö"   ' capital O, dieresis or umlaut mark
        p_astrISOTable(215) = "×"   ' multiply sign
        p_astrISOTable(216) = "Ø"   ' capital O, slash
        p_astrISOTable(217) = "Ù"   ' capital U, grave accent
        p_astrISOTable(218) = "Ú"   ' capital U, acute accent
        p_astrISOTable(219) = "Û"   ' capital U, circumflex accent
        p_astrISOTable(220) = "Ü"   ' capital U, dieresis or umlaut mark
        p_astrISOTable(221) = "Ý"   ' capital Y, acute accent
        p_astrISOTable(222) = "Þ"   ' capital THORN, Icelandic
        p_astrISOTable(223) = "ß"   ' small sharp s, German (sz ligature)
        p_astrISOTable(224) = "à"   ' small a, grave accent
        p_astrISOTable(225) = "á"   ' small a, acute accent
        p_astrISOTable(226) = "â"   ' small a, circumflex accent
        p_astrISOTable(227) = "ã"   ' small a, tilde
        p_astrISOTable(228) = "ä"   ' small a, dieresis or umlaut mark
        p_astrISOTable(229) = "å"   ' small a, ring
        p_astrISOTable(230) = "æ"   ' small ae diphthong (ligature)
        p_astrISOTable(231) = "ç"   ' small c, cedilla
        p_astrISOTable(232) = "è"   ' small e, grave accent
        p_astrISOTable(233) = "é"   ' small e, acute accent
        p_astrISOTable(234) = "ê"   ' small e, circumflex accent
        p_astrISOTable(235) = "ë"   ' small e, dieresis or umlaut mark
        p_astrISOTable(236) = "ì"   ' small i, grave accent
        p_astrISOTable(237) = "í"   ' small i, acute accent
        p_astrISOTable(238) = "î"   ' small i, circumflex accent
        p_astrISOTable(239) = "ï"   ' small i, dieresis or umlaut mark
        p_astrISOTable(240) = "ð"   ' small eth, Icelandic
        p_astrISOTable(241) = "ñ"   ' small n, tilde
        p_astrISOTable(242) = "ò"   ' small o, grave accent
        p_astrISOTable(243) = "ó"   ' small o, acute accent
        p_astrISOTable(244) = "ô"   ' small o, circumflex accent
        p_astrISOTable(245) = "õ"   ' small o, tilde
        p_astrISOTable(246) = "ö"   ' small o, dieresis or umlaut mark
        p_astrISOTable(247) = "÷"   ' division sign
        p_astrISOTable(248) = "ø"   ' small o, slash
        p_astrISOTable(249) = "ù"   ' small u, grave accent
        p_astrISOTable(250) = "ú"   ' small u, acute accent
        p_astrISOTable(251) = "û"   ' small u, circumflex accent
        p_astrISOTable(252) = "ü"   ' small u, dieresis or umlaut mark
        p_astrISOTable(253) = "ý"   ' small y, acute accent
        p_astrISOTable(254) = "þ"   ' small thorn, Icelandic
        p_astrISOTable(255) = "ÿ"   ' small y, dieresis or umlaut mark
    
    Case Else
            
        ' TODO: Implement support for Charsets other than ISO-8859-1 (Latin)
            
    End Select

ExitHere:
    Exit Sub
ErrorHandler:
    Debug.Assert 0  ' Soft stop
    ' TODO: Implement your error hanling/logging here
    Resume ExitHere
End Sub



