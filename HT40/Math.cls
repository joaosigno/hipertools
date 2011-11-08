VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Type Result
    res As String
    rmm As String
End Type


Private Type Factor
    Base As Double
    Power As Double
End Type

Private Const LOCALE_SDECIMAL = &HE 'decimal separator

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Cancel As Boolean, Listing As Boolean

Private Const csValidChars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

' Secant
Function Sec(X As Double) As Double
   Sec = 1 / Cos(X)
End Function

' Cosecant
Function CoSec(X As Double) As Double
   CoSec = 1 / sIn(X)
End Function

' Cotangent
Function CoTan(X As Double) As Double
   CoTan = 1 / Tan(X)
End Function

' Inverse Sine
Function ArcSin(X As Double) As Double
   ArcSin = Atn(X / Sqr(-X * X + 1))
End Function

' Inverse Cosine
Function ArcCos(X As Double) As Double
   ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

' Inverse Secant
Function ArcSec(X As Double) As Double
   ArcSec = Atn(X / Sqr(X * X - 1)) + Sgn(X - 1) * (2 * Atn(1))
End Function

' Inverse Cosecant
Function ArcCoSec(X As Double) As Double
   ArcCoSec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
End Function

' Inverse Cotangent
Function ArcCoTan(X As Double) As Double
   ArcCoTan = Atn(X) + 2 * Atn(1)
End Function

' Hyperbolic Sine
Function HSin(X As Double) As Double
   HSin = (Exp(X) - Exp(-X)) / 2
End Function

' Hyperbolic Cosine
Function HCos(X As Double) As Double
   HCos = (Exp(X) + Exp(-X)) / 2
End Function

' Hyperbolic Tangent
Function HTan(X As Double) As Double
   HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function

' Hyperbolic Secant
Function HSec(X As Double) As Double
   HSec = 2 / (Exp(X) + Exp(-X))
End Function

' Hyperbolic Cosecant
Function HCoSec(X As Double) As Double
   HCoSec = 2 / (Exp(X) - Exp(-X))
End Function

' Hyperbolic Cotangent
Function HCotan(X As Double) As Double
   HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
End Function

' Inverse Hyperbolic Sine
Function HArcSin(X As Double) As Double
   HArcSin = Log(X + Sqr(X * X + 1))
End Function

' Inverse Hyperbolic Cosine
Function HArcCos(X As Double) As Double
   HArcCos = Log(X + Sqr(X * X - 1))
End Function

' Inverse Hyperbolic Tangent
Function HArcTan(X As Double) As Double
   HArcTan = Log((1 + X) / (1 - X)) / 2
End Function

' Inverse Hyperbolic Secant
Function HArcSec(X As Double) As Double
   HArcSec = Log((Sqr(-X * X + 1) + 1) / X)
End Function

' Inverse Hyperbolic Cosecant
Function HArcCoSec(X As Double) As Double
   HArcCoSec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function

' Inverse Hyperbolic Cotangent
Function HArcCoTan(X As Double) As Double
   HArcCoTan = Log((X + 1) / (X - 1)) / 2
End Function

' Logarithm to base N
Function LogN(X As Double, n As Double) As Double
   LogN = Log(X) / Log(n)
End Function

Function Dec2Fract(X As Single) As String
   ' Returns a string of a number rounded to a whole and
   'fraction in 16ths, 32nds, etc.
   ' Handles number > 1 and less than 0
   ' Use as follows:
   ' Label1.Caption = Dec2Fract(Val(Text1.Text))
   ' or Label1.Caption = Dec2Fract(Val(Label1.Caption)
   
   Dim f As String, Y As Single, Num As Integer, Den As Integer
   
   Den = 16  'Denominator: can be set to 8, 16, 32, 64 etc
   
   If X = 0 Then
      Dec2Fract = "0"
      Exit Function
   Else
      
      Y = Abs(X)
      If Y > 1 Then Y = Y - Int(Y)  ' get fractional part
      Num = CInt(Den * Y)
      
      If Num = Den Then
         f = "1"
      ElseIf Num = 0 Then
         If Abs(X) < 1 Then f = "0" Else f = ""
      Else
         Do Until Num Mod 2 <> 0
            Num = Num / 2
            Den = Den / 2
         Loop
         f = LTrim$(Str$(Num)) + "/" + LTrim$(Str$(Den))
      End If
      
      If Abs(X) > 1 Then
         If f <> "1" Then
            f = Trim$(Str$(Fix(X))) + " " + f
         Else
            f = Trim$(Str$(CInt(X)))
         End If
      End If
      
      If X < 0 And X > -1 Then f = "-" + f
      
      Dec2Fract = f
   End If
End Function

Private Function GetXfromAngle(Angle, Length, Buiging)
Select Case Angle
    Case 0 To 89, 360
        GetXfromAngle = Length - (90 - Angle Mod 90) ^ _
         Buiging * Length / 90 ^ Buiging
    Case 90 To 179
        GetXfromAngle = Length - (Angle Mod 90) ^ _
            Buiging * Length / 90 ^ Buiging
    Case 180 To 269
        GetXfromAngle = (90 - Angle Mod 90) ^ Buiging * _
            Length / 90 ^ Buiging - Length
    Case 270 To 359
        GetXfromAngle = (Angle Mod 90) ^ Buiging * _
            Length / 90 ^ Buiging - Length
End Select
End Function

Private Function GetYfromAngle(Angle, Length, Buiging)
Select Case Angle
    Case 0 To 89, 360
        GetYfromAngle = (Angle Mod 90) ^ Buiging * _
             Length / 90 ^ Buiging
    Case 90 To 179
        GetYfromAngle = Length - (90 - Angle Mod 90) ^ _
             Buiging * Length / 90 ^ Buiging + Length
    Case 180 To 269
        GetYfromAngle = Length - (Angle Mod 90) ^ _
             Buiging * Length / 90 ^ Buiging + Length
    Case 270 To 360
        GetYfromAngle = (90 - Angle Mod 90) ^ Buiging * _
             Length / 90 ^ Buiging
End Select
End Function

'Function takes an array with numeric elements as a parameter
'and calcuates the standard deviation

Public Function StandardDeviation(NumericArray As Variant) _
As Double

Dim dblSum As Double, dblSumSqdDevs As Double, dblMean As Double
Dim lngCount As Long, dblAnswer As Double
Dim vElement As Variant
Dim lngStartPoint As Long, lngEndPoint As Long, lngCtr As Long

On Error GoTo ErrorHandler
'if NumericArray is not an array, this statement will
'raise an error in the errorhandler

lngCount = UBound(NumericArray)

On Error Resume Next
lngCount = 0

'the check below will allow
'for 0 or 1 based arrays.

vElement = NumericArray(0)

lngStartPoint = IIf(Err.number = 0, 0, 1)
lngEndPoint = UBound(NumericArray)

'get sum and sample size
For lngCtr = lngStartPoint To lngEndPoint
  vElement = NumericArray(lngCtr)
    If IsNumeric(vElement) Then
      lngCount = lngCount + 1
      dblSum = dblSum + CDbl(vElement)
    End If
Next

'get mean
If lngCount > 1 Then
    dblMean = dblSum / lngCount

    'get sum of squared deviations
    For lngCtr = lngStartPoint To lngEndPoint
        vElement = NumericArray(lngCtr)

        If IsNumeric(vElement) Then
            dblSumSqdDevs = dblSumSqdDevs + _
            ((vElement - dblMean) ^ 2)
        End If
    Next

'divide result by sample size - 1 and get square root.
'this function calculates standard deviation of a sample.
'If your  set of values represents the population, use sample
'size not sample size - 1

    If lngCount > 1 Then
        lngCount = lngCount - 1 'eliminate for population values
        dblAnswer = Sqr(dblSumSqdDevs / lngCount)
    End If
    
End If

StandardDeviation = dblAnswer

Exit Function

ErrorHandler:
Err.Raise Err.number
Exit Function
End Function

Private Function CBin(number As Integer) As String
Rem Convert a Decimal Number to Binary
Dim Temp As Variant
Temp = 1 'Can't fouble nothing
Do Until Temp > number 'sets starting point for Len
    Temp = Temp * 2
Loop
Do Until Temp < 1
    If number >= Temp Then
       CBin = CBin + "1"
       number = number - Temp
    Else
        CBin = CBin + "0"
    End If
    Temp = Temp / 2
Loop 'Loop until string is complete
CBin = CStr(Val(CBin))
End Function

Function SetBytes(Bytes) As String
Rem Convert Bytes to the Appropriate Format (Bytes, KB, MB, GB)

On Error GoTo hell

If Bytes >= 1073741824 Then
    SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.00") _
         & " GB"
ElseIf Bytes >= 1048576 Then
    SetBytes = Format(Bytes / 1024 / 1024, "#0.00") & " MB"
ElseIf Bytes >= 1024 Then
    SetBytes = Format(Bytes / 1024, "#0.00") & " KB"
ElseIf Bytes < 1024 Then
    SetBytes = Fix(Bytes) & " Bytes"
End If

Exit Function
hell:
SetBytes = "0 Bytes"
End Function

Public Function Dec2Frac(ByVal f As Double) As String
Rem Convert a Decimal to a Fraction

   Dim df As Double
   Dim lUpperPart As Long
   Dim lLowerPart As Long
   
   lUpperPart = 1
   lLowerPart = 1
   
   df = lUpperPart / lLowerPart
   While (df <> f)
      If (df < f) Then
         lUpperPart = lUpperPart + 1
      Else
         lLowerPart = lLowerPart + 1
         lUpperPart = f * lLowerPart
      End If
      df = lUpperPart / lLowerPart
   Wend
Dec2Frac = CStr(lUpperPart) & "/" & CStr(lLowerPart)
End Function

Public Function BCon(ByVal InputValue As String, ByVal _
  InputBase As Long, ByVal OutputBase As Long) As String

Rem Convert Values Between Different Bases

'// Declare following variables
Dim cuOutputCharCount As Currency, lInputCharCount As Long
Dim lCounter As Long, sCompareWith As String, sChar As String
Dim dDecimalChars As Double, lPos As Long, cuBitValue As Currency
Dim cuDecimalValue As Currency, sOutput As String

'// Turn on error trapping
On Error Resume Next

'// Check to see that both InputBase and OutputBase are
'between 2 and 36 inclusively
If (InputBase < 2) Or (InputBase > 36) Then
   BCon = ""
   Exit Function
End If

If (OutputBase < 2) Or (OutputBase > 36) Then
   BCon = ""
   Exit Function
End If

'// Now initiate translation into 10-based value
InputValue = Trim(InputValue)
lInputCharCount = Len(InputValue)
sCompareWith = Left$(csValidChars, InputBase)

'// Check to see that input string is at least one character
If lInputCharCount < 1 Then
   BCon = ""
   Exit Function
End If

'// Loop through each character in input string. Check for
   'invalid characters according to input base
For lCounter = 1 To lInputCharCount
   sChar = Mid$(InputValue, lCounter, 1)
   If InStr(1, sCompareWith, sChar, vbTextCompare) < 1 Then
      BCon = ""
      Exit Function
   End If
Next

If InputBase = OutputBase Then
   BCon = InputValue
   Exit Function
End If

'// Now get the number of decimal characters needed to hold
   'input value
dDecimalChars = lInputCharCount * Log(InputBase) / Log(10)

'// If decimal string is longer than 14 characters then
   'exit function
If dDecimalChars > 14 Then
   BCon = ""
   Exit Function
End If

If InputBase = 10 Then
   cuDecimalValue = CCur(InputValue)
Else
   '// Initiate calculation into decimal value
   cuBitValue = 1
   cuDecimalValue = 0

   '// As long there are characters left in input string, add
   'their value into the decimal hold variable
   While Len(InputValue) > 0
      sChar = Right(InputValue, 1)
      lPos = InStr(1, sCompareWith, sChar, vbTextCompare) - 1
      cuDecimalValue = cuDecimalValue + lPos * cuBitValue
      cuBitValue = cuBitValue * InputBase
      InputValue = Left(InputValue, Len(InputValue) - 1)
   Wend

   '// Now we can do a trick. If output base is 10, we already
    'have the decimal value...
   If OutputBase = 10 Then
      BCon = CStr(cuDecimalValue)
      Exit Function
   End If
End If

'// Calculate number of characters in output string. If not an
   'integer round up to nearest integer
cuOutputCharCount = 1 + Log(cuDecimalValue) / Log(OutputBase)

'// Now we must convert the decimal value back into a string in
   'the output base format
cuBitValue = 1
For lCounter = 2 To cuOutputCharCount
   cuBitValue = cuBitValue * OutputBase
Next

'// As long as we have a decimal value greater than 0, add a
   'character to the output string
sOutput = ""
sCompareWith = Left$(csValidChars, OutputBase)
While cuBitValue > 0
   lPos = cuDecimalValue \ cuBitValue
   sChar = Mid$(sCompareWith, 1 + lPos, 1)
   sOutput = sOutput & sChar
   cuDecimalValue = cuDecimalValue - lPos * cuBitValue
   cuBitValue = cuBitValue \ OutputBase
Wend

'// Strip leading zeros
While Left$(sOutput, 1) < "1"
   sOutput = Right$(sOutput, Len(sOutput) - 1)
Wend

'// Let BCon function value be the output string in output base
'format
BCon = sOutput

End Function

Rem Demo of Various Mathematical Functions and Algorithms
Rem DIVS is algorithm for optimized stack assisted routine for integer division
Rem based on additions and subtractions only. It calculates result and remainder
Rem simultaneously.
Rem FACTS is an algorithm for factorial calculations. It has extra-reduced number
Rem of operations that is based on decomposition of factorial to the product of the
Rem prime number powers, e.g. 10! = 2^8 � 3^4 � 5^2 � 7.
Rem EVAL is recursive routine for evaluation of algebraic expressions with
Rem brackets and operation priority.
Sub SWAP(a As Variant, b As Variant)
    Dim c As Variant
    c = a
    a = b
    b = c
End Sub

Function NoLead0(ByVal a As String, ByVal D As String) As String
    Dim s As String
    a = Trim(a)
    s = Left$(a, 1)
    If s = "-" Then
        a = Mid$(a, 2)
    Else
        s = ""
    End If
    a = Replace$(LTrim$(Replace$(a, Left$(D, 1), " ")), " ", Left$(D, 1))
    NoLead0 = IIf(a = "", Left$(D, 1), s & a)
End Function

Function ABSS(ByVal a As String, D As String) As String
    a = NoLead0(a, D)
    ABSS = IIf(Left$(a, 1) = "-", Mid$(a, 2), a)
End Function

Function COMPS(ByVal b As String, ByVal c As String, D As String) As Integer
    Dim i As Long, j As Long, k As Long
    If Len(c) = Len(b) Then
        For i = 1 To Len(c)
            j = InStr(D, Mid$(b, i, 1))
            k = InStr(D, Mid$(c, i, 1))
            If j <> k Then
                If j < k Then
                    COMPS = -1
                    Exit Function
                Else
                    COMPS = 1
                    Exit Function
                End If
            End If
        Next
        COMPS = 0
    Else
        If Len(c) > Len(b) Then
            COMPS = -1
        Else
            COMPS = 1
        End If
    End If
End Function

Function ADDS(ByVal b As String, ByVal c As String, D As String) As String
    Dim n As Long, of As Long, i As Long, f As Long, f1 As Long, a As String, ppp As String
    n = Len(D)
    of = Len(c) - Len(b)
    If of < 0 Then
        SWAP b, c
        of = -of
    End If
    a = ""
    f = 0
    For i = Len(b) To 1 Step -1
        f1 = f + InStr(D, Mid$(b, i, 1)) + InStr(D, Mid$(c, of + i, 1)) - 2
        If f1 >= n Then
            f = 1
            f1 = f1 - n
        Else
            f = 0
        End If
        a = Mid$(D, 1 + f1, 1) & a
    Next
    If of Then
        For i = of To 1 Step -1
            If f Then
                f1 = f + InStr(D, Mid$(c, i, 1)) - 1
                If f1 >= n Then
                    f = 1
                    f1 = f1 - n
                Else
                    f = 0
                End If
                a = Mid$(D, 1 + f1, 1) & a
            Else
                a = Mid$(c, 1, i) & a
                Exit For
            End If
        Next
    End If
    If f Then ADDS = Mid$(D, 2, 1) & a Else ADDS = a
End Function

Function SUBS(ByVal b As String, ByVal c As String, D As String) As String
    Dim s As String, g As String, i As Long
    If COMPS(b, c, D) < 0 Then
        SWAP c, b
        g = "-"
    Else
        g = ""
    End If
    s = String$(Len(b) - Len(c), Right$(D, 1))
    For i = 1 To Len(c)
        s = s & Mid$(D, 1 + Len(D) - InStr(D, Mid$(c, i, 1)), 1)
    Next
    SUBS = g & NoLead0(Right$(ADDS(b, ADDS(s, Mid$(D, 2, 1), D), D), Len(b)), D)
End Function

Function MULS(ByVal b As String, ByVal c As String, D As String) As String
    Dim i As Long, j As Long, n As Long, f As Long, f1 As Long
    Dim m As Long, a As String, p As String, nul As String
    n = Len(D)
    If Len(b) > Len(c) Then SWAP b, c
    a = ""
    nul = ""
    For i = Len(b) To 1 Step -1
        m = InStr(D, Mid$(b, i, 1)) - 1
        p = ""
        f = 0
        For j = Len(c) To 1 Step -1
            f1 = f + m * (InStr(D, Mid$(c, j, 1)) - 1)
            If f1 >= n Then
                f = f1 \ n
                f1 = f1 Mod n
            Else
                f = 0
            End If
            p = Mid$(D, 1 + f1, 1) + p
        Next
        If f Then p = Mid$(D, 1 + f, 1) + p
        p = p + nul
        nul = nul + Left$(D, 1)
        a = ADDS(a, p, D)
    Next
    MULS = a
End Function

' The DIVS subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function DIVS(ByVal a As String, ByVal b As String, dg As String) As Result
    Dim D As String, c As String, p As String, Stack As New Collection
    If b = Left$(dg, 1) Then
        DIVS.res = 0
        DIVS.rmm = 0
    Else
        D = Mid$(dg, 2, 1)
        Do
            Stack.Add D
            Stack.Add b
            D = ADDS(D, D, dg)
            b = ADDS(b, b, dg)
        Loop While COMPS(b, a, dg) <= 0
        b = Stack.Item(Stack.Count): Stack.Remove Stack.Count
        b = SUBS(a, b, dg)
        p = Stack.Item(Stack.Count): Stack.Remove Stack.Count
        If p = Mid$(dg, 2, 1) Then
            If Left$(b, 1) = "-" Then
                p = Left$(dg, 1)
                b = a
            End If
        Else
            Do
                D = Stack.Item(Stack.Count): Stack.Remove Stack.Count
                c = Stack.Item(Stack.Count): Stack.Remove Stack.Count
                D = SUBS(D, b, dg)
                If D = Left$(dg, 1) Or Left$(D, 1) = "-" Then
                    b = IIf(Left$(D, 1) = "-", Mid$(D, 2), D)
                    p = ADDS(p, c, dg)
                End If
            Loop Until c = Mid$(dg, 2, 1)
        End If
        DIVS.res = p
        DIVS.rmm = b
    End If
End Function

Function POWS(ByVal c As String, ByVal b As String, D As String) As String
    Dim s As String, ppp As String
    Static ff As Object
    If TypeName(ff) <> "Dictionary" Then Set ff = CreateObject("Scripting.Dictionary")
    ppp = c & "|" & b & "|" & D
    If ff.Exists(ppp) Then
        POWS = ff(ppp)
    Else
        b = A2B(b, D, "01")
        s = Mid$(D, 2, 1)
        Do
            If Right$(b, 1) = "1" Then s = MULS(s, c, D)
            b = Left$(b, Len(b) - 1)
            If b <> "" Then c = MULS(c, c, D) Else Exit Do
        Loop
        ff.Add ppp, s
        POWS = s
    End If
End Function

Function XORS(ByVal b As String, ByVal c As String, D As String) As String
    Dim i As Long, s As String
    If Len(b) > Len(c) Then SWAP b, c
    If InStr("|1|2|4|8", "|" & Trim$(Replace$(Hex$(Len(D)), "0", " "))) Then
        s = ""
        b = StrReverse(b)
        c = StrReverse(c)
        For i = 1 To Len(b)
            s = s & Mid$(D, ((InStr(D, Mid$(b, i, 1)) - 1) Xor (InStr(D, Mid$(c, i, 1)) - 1)) + 1, 1)
        Next
        XORS = NoLead0(StrReverse(s & Mid$(c, Len(b) + 1)), D)
    Else
        b = StrReverse(A2B(b, D, "01234567"))
        c = StrReverse(A2B(c, D, "01234567"))
        s = ""
        For i = 1 To Len(b)
            s = s & Oct$(Val(Mid$(b, i, 1)) Xor Val(Mid$(c, i, 1)))
        Next
        XORS = A2B(StrReverse(s & Mid$(c, Len(b) + 1)), "01234567", D)
    End If
End Function

Function ORS(ByVal b As String, ByVal c As String, D As String) As String
    Dim i As Long, s As String
    If Len(b) > Len(c) Then SWAP b, c
    If InStr("|1|2|4|8", "|" & Trim$(Replace$(Hex$(Len(D)), "0", " "))) Then
        s = ""
        b = StrReverse(b)
        c = StrReverse(c)
        For i = 1 To Len(b)
            s = s & Mid$(D, ((InStr(D, Mid$(b, i, 1)) - 1) Or (InStr(D, Mid$(c, i, 1)) - 1)) + 1, 1)
        Next
        ORS = NoLead0(StrReverse(s & Mid$(c, Len(b) + 1)), D)
    Else
        b = StrReverse(A2B(b, D, "01234567"))
        c = StrReverse(A2B(c, D, "01234567"))
        s = ""
        For i = 1 To Len(b)
            s = s & Oct$(Val(Mid$(b, i, 1)) Or Val(Mid$(c, i, 1)))
        Next
        ORS = A2B(StrReverse(s & Mid$(c, Len(b) + 1)), "01234567", D)
    End If
End Function

Function ANDS(ByVal b As String, ByVal c As String, D As String) As String
    Dim i As Long, s As String
    If Len(b) > Len(c) Then SWAP b, c
    If InStr("|1|2|4|8", "|" & Trim$(Replace$(Hex$(Len(D)), "0", " "))) Then
        s = ""
        b = StrReverse(b)
        c = StrReverse(c)
        For i = 1 To Len(b)
            s = s & Mid$(D, ((InStr(D, Mid$(b, i, 1)) - 1) And (InStr(D, Mid$(c, i, 1)) - 1)) + 1, 1)
        Next
        ANDS = NoLead0(StrReverse(s), D)
    Else
        b = StrReverse(A2B(b, D, "01234567"))
        c = StrReverse(A2B(c, D, "01234567"))
        s = ""
        For i = 1 To Len(b)
            s = s & Oct$(Val(Mid$(b, i, 1)) And Val(Mid$(c, i, 1)))
        Next
        ANDS = A2B(StrReverse(s), "01234567", D)
    End If
End Function

Function NEGS(b As String, D As String) As String
    NEGS = A2B(Replace(Replace(Replace(A2B(b, D, "01"), "0", "X"), "1", "0"), "X", "1"), "01", D)
End Function

Function ChgSgn(a As String) As String
    ChgSgn = IIf(Left$(a, 1) = "-", Mid$(a, 2), "-" & a)
End Function

Function SUBA(ByVal b As String, ByVal c As String, D As String) As String
    Dim s As Integer
    b = NoLead0(b, D)
    c = NoLead0(c, D)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0
        SUBA = SUBS(b, c, D)
    Case 3
        SUBA = ChgSgn(SUBS(b, c, D))
    Case 1
        SUBA = "-" & ADDS(b, c, D)
    Case 2
        SUBA = ADDS(b, c, D)
    End Select
End Function

Function ORA(b As String, c As String, D As String) As String
    ORA = ORS(ABSS(b, D), ABSS(c, D), D)
End Function

Function ANDA(b As String, c As String, D As String) As String
    ANDA = ANDS(ABSS(b, D), ABSS(c, D), D)
End Function

Function XORA(ByVal b As String, ByVal c As String, D As String) As String
    XORA = XORS(ABSS(b, D), ABSS(c, D), D)
End Function

Function NEGA(b As String, D As String) As String
    NEGA = NEGS(ABSS(b, D), D)
End Function

Function ADDA(ByVal b As String, ByVal c As String, D As String) As String
    Dim s As Integer
    b = NoLead0(b, D)
    c = NoLead0(c, D)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0
        ADDA = ADDS(b, c, D)
    Case 3
        ADDA = "-" & ADDS(b, c, D)
    Case 1
        ADDA = SUBS(c, b, D)
    Case 2
        ADDA = SUBS(b, c, D)
    End Select
End Function

Function MULA(ByVal b As String, ByVal c As String, D As String) As String
    Dim s As Integer
    b = NoLead0(b, D)
    c = NoLead0(c, D)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0, 3
        MULA = MULS(b, c, D)
    Case 1, 2
        MULA = "-" & MULS(b, c, D)
    End Select
End Function

Function DIVA(ByVal b As String, ByVal c As String, D As String) As String
    Dim s As Integer
    b = NoLead0(b, D)
    c = NoLead0(c, D)
    If Left$(b, 1) = "-" Then
        s = 1
        b = Mid$(b, 2)
    Else
        s = 0
    End If
    If Left$(c, 1) = "-" Then
        s = s Or 2
        c = Mid$(c, 2)
    End If
    Select Case s
    Case 0, 3
        DIVA = DIVS(b, c, D).res
    Case 1, 2
        DIVA = "-" & DIVS(b, c, D).res
    End Select
End Function

' The POWA subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function POWA(ByVal b As String, ByVal c As String, D As String) As String
    Dim s As Integer
    b = NoLead0(b, D)
    c = NoLead0(c, D)
    c = IIf(Left$(c, 1) = "-", Mid$(c, 2), c)
    If Left$(b, 1) = "-" Then
        b = Mid$(b, 2)
        If InStr(D, Right$(c, 1)) And 1 Then
            POWA = POWS(b, c, D)
        Else
            POWA = "-" & POWS(b, c, D)
        End If
    Else
        POWA = POWS(b, c, D)
    End If
End Function

Function ChooseOp(a As String, b As String, c As String, D As String) As String
    Select Case a
    Case "+"
        ChooseOp = ADDA(b, c, D)
    Case "%"
        ChooseOp = ORA(b, c, D)
    Case "&"
        ChooseOp = ANDA(b, c, D)
    Case "@"
        ChooseOp = XORA(b, c, D)
    Case "-"
        ChooseOp = SUBA(b, c, D)
    Case "*"
        ChooseOp = MULA(b, c, D)
    Case "/"
        ChooseOp = DIVA(b, c, D)
    Case "^"
        ChooseOp = POWA(b, c, D)
    Case "~"
        ChooseOp = NEGA(c, D)
    Case "�"
        ChooseOp = ChgSgn(c)
    Case "!"
        ChooseOp = FACTA(b, D)
    End Select
End Function

' The EVAL subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function EVAL(ByVal a As String, l As String, D As String) As String
    Dim i As Long, j As Variant, b As String
    a = NoLead0(a, D)
    i = UBound(Split(a, "(")) - UBound(Split(a, ")"))
    If i > 0 Then a = a & String(i, ")")
    a = Replace(a, "#", l)
    i = InStr(a, "-")
    Do While i
        If InStr("(+-%&@/*^�~", Mid$(a, IIf(i = 1, 1, i - 1), 1)) Then Mid$(a, i, 1) = "�"
        i = InStr(i + 1, a, "-")
    Loop
    Do
        b = XB(a)
        If b = String$(Len(a), "X") Then
            a = Mid$(a, 2, Len(a) - 2)
        Else
            Exit Do
        End If
    Loop
    If b <> "Y" Then
        For Each j In Array("+", "-", "%", "&", "@", "/", "*", "^", "�", "~", "!")
            i = InStrRev(b, j)
            If i Then
                EVAL = ChooseOp(CStr(j), EVAL(Left$(a, i - 1), l, D), EVAL(Mid$(a, i + 1), l, D), D)
                Exit Function
            End If
        Next
        EVAL = NoLead0(a, D)
    Else
        EVAL = Left$(D, 1)
    End If
End Function

Function XB(a As String) As String
    Dim i As Long, j As Long
    XB = a
    If InStr(a, "(") Then
        j = 0
        For i = 1 To Len(a)
            Select Case Mid$(a, i, 1)
            Case "("
                j = j + 1
                If j > 0 Then Mid(XB, i, 1) = "X"
            Case ")"
                If j > 0 Then Mid(XB, i, 1) = "X"
                j = j - 1
            Case Else
                If j > 0 Then Mid(XB, i, 1) = "X"
            End Select
        Next
    End If
    If j Then XB = "Y"
End Function

Function A2BS(ByVal a As String, da As String, db As String) As String
    Dim sg As String, s As String, m As String, i As Long, ppp As String
    ReDim n(0 To Len(da)) As String
    Static ff As Object
    If TypeName(ff) <> "Dictionary" Then Set ff = CreateObject("Scripting.Dictionary")
    ppp = a & "|" & da & "|" & db
    If ff.Exists(ppp) Then
        A2BS = ff(ppp)
    Else
        A2BS = A2B(a, da, db)
        ff.Add ppp, A2BS
    End If
End Function

Function A2B(ByVal a As String, da As String, db As String) As String
    Dim sg As String, s As String, m As String, i As Long, ppp As String
    ReDim n(0 To Len(da)) As String
    If Left$(a, 1) = "-" Then
        sg = "-"
        a = Mid$(a, 2)
    Else
        sg = ""
    End If
    a = NoLead0(UCase$(a), da)
    m = Mid$(db, 2, 1)
    s = Left$(db, 1)
    n(0) = s
    For i = 1 To Len(da)
        n(i) = ADDS(m, n(i - 1), db)
    Next
    For i = Len(a) To 2 Step -1
        s = ADDS(s, MULS(n(InStr(da, Mid$(a, i, 1)) - 1), m, db), db)
        m = MULS(n(Len(da)), m, db)
    Next
    A2B = sg & ADDS(s, MULS(n(InStr(da, Mid$(a, i, 1)) - 1), m, db), db)
End Function

Function ValS(c As String, D As String) As Long
    Dim b As Long, l As Long, i As Long
    Select Case D
    Case "0123456789"
        ValS = CLng(c)
    Case "01234567"
        ValS = Val("&O" & c)
    Case "0123456789ABCDEF"
        ValS = Val("&H" & c)
    Case Else
        ValS = 0
        b = 1
        l = Len(D)
        For i = Len(c) To 1 Step -1
            ValS = ValS + b * (InStr(D, Mid$(c, i, 1)) - 1)
            b = b * l
        Next
    End Select
End Function

Function StrS(ByVal a As Long, D As String) As String
    Dim l As Long
    Select Case D
    Case "0123456789"
        StrS = CStr(a)
    Case "01234567"
        StrS = Oct$(a)
    Case "0123456789ABCDEF"
        StrS = Hex$(a)
    Case Else
        StrS = ""
        l = Len(D)
        Do While a
            StrS = Mid$(D, 1 + (a Mod l), 1) & StrS
            a = a \ l
        Loop
    End Select
End Function

' The FACTS subroutine is based on the original algoritham that is derived
' by the program's author: Dipl. Ing. Andrija Radovic
Function FACTS(aa As String, D As String) As String
    Static ff As Object
    Dim fa As Variant, a As Long, b As Long, c As Long, i As Long, dd As Long
    Dim f As String, cc As String
    fa = Array(3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, 103, _
    107, 109, 113, 127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, 199, _
    211, 223, 227, 229, 233, 239, 241, 251, 257, 263, 269, 271, 277, 281, 283, 293, 307, 311, 313, _
    317, 331, 337, 347, 349, 353, 359, 367, 373, 379, 383, 389, 397, 401, 409, 419, 421, 431, 433, _
    439, 443, 449, 457, 461, 463, 467, 479, 487, 491, 499, 503, 509, 521, 523, 541, 547, 557, 563, _
    569, 571, 577, 587, 593, 599, 601, 607, 613, 617, 619, 631, 641, 643, 647, 653, 659, 661, 673, _
    677, 683, 691, 701, 709, 719, 727, 733, 739, 743, 751, 757, 761, 769, 773, 787, 797, 809, 811, _
    821, 823, 827, 829, 839, 853, 857, 859, 863, 877, 881, 883, 887, 907, 911, 919, 929, 937, 941, _
    947, 953, 967, 971, 977, 983, 991, 997, 1009, 1013, 1019, 1021, 1031, 1033, 1039, 1049, 1051, _
    1061, 1063, 1069, 1087, 1091, 1093, 1097, 1103, 1109, 1117, 1123, 1129, 1151, 1153, 1163, _
    1171, 1181, 1187, 1193, 1201, 1213, 1217, 1223, 1229, 1231, 1237, 1249, 1259, 1277, 1279, _
    1283, 1289, 1291, 1297, 1301, 1303, 1307, 1319, 1321, 1327, 1361, 1367, 1373, 1381, 1399, _
    1409, 1423, 1427, 1429, 1433, 1439, 1447, 1451, 1453, 1459, 1471, 1481, 1483, 1487, 1489, _
    1493, 1499, 1511, 1523, 1531, 1543, 1549, 1553, 1559, 1567, 1571, 1579, 1583, 1597, 1601, _
    1607, 1609, 1613, 1619, 1621, 1627, 1637, 1657, 1663, 1667, 1669, 1693, 1697, 1699, 1709, _
    1721, 1723, 1733, 1741, 1747, 1753, 1759, 1777, 1783, 1787, 1789, 1801, 1811, 1823, 1831, _
    1847, 1861, 1867, 1871, 1873, 1877, 1879, 1889, 1901, 1907, 1913, 1931, 1933, 1949, 1951, _
    1973, 1979, 1987, 1993, 1997, 1999, 2003, 2011, 2017, 2027, 2029, 2039, 2053, 2063, 2069)
    If TypeName(ff) <> "Dictionary" Then Set ff = CreateObject("Scripting.Dictionary")
    If ff.Exists(aa & "|" & D) Then
        FACTS = ff(aa & "|" & D)
    Else
        a = ValS(aa, D)
        f = "1"
        cc = "2"
        c = 2
        i = 0
        Do
            b = a
            dd = 0
            Do
                b = b \ c
                dd = dd + b
            Loop While b > 1
            f = MULS$(f, POWS$(cc, StrS$(dd, D), D), D)
            c = fa(i)
            cc = StrS$(c, D)
            i = i + 1
        Loop Until c > a
        ff.Add aa & "|" & D, f
        FACTS = f
    End If
End Function

Function FACTA(ByVal a As String, D As String) As String
    a = NoLead0(a, D)
    FACTA = FACTS(IIf(Left$(a, 1) = "-", Mid$(a, 2), a), D)
End Function

Function Sintax(a As String, b As String, c As String) As Boolean
    Select Case a
    Case "~", "-"
        Select Case b
        Case "+", "%", "&", "@", "/", "*", "^", "!", "~", ")", "-"
            Sintax = False
        Case Else
            Sintax = True
        End Select
    Case "("
        Select Case b
        Case "+", "%", "&", "@", "/", "*", "^", "!", ")"
            Sintax = False
        Case Else
            Sintax = True
        End Select
    Case "+", "%", "&", "@", "/", "*", "^"
        Select Case b
        Case "+", "%", "&", "@", "/", "*", "^", "!", ")"
            Sintax = False
        Case Else
            Sintax = True
        End Select
    Case "!"
        Sintax = IIf(b = "!", False, True)
    Case Else
        Select Case b
        Case "~"
            Sintax = False
        Case ")"
            If UBound(Split(c, "(")) > UBound(Split(c, ")")) Then
                Sintax = True
            Else
                Sintax = False
            End If
        Case Else
            Sintax = True
        End Select
    End Select
End Function

Rem Determine if a Number is Even or Odd
Public Function IsEven(ByVal number As Long) As Boolean
    IsEven = (number Mod 2 = 0)
End Function

Rem Determine if an Integer is a Prime Number
Public Function IsPrime(ByVal lValue As Long) As Boolean

Dim lCtr As Long
Dim lRes As Long
Dim lCount As Long

'eliminate even numbers right away
lValue = Abs(lValue)
If lValue Mod 2 = 0 Then Exit Function

For lCtr = 1 To lValue
    lRes = lValue Mod lCtr
    If lRes <> 0 Then
        lCount = lCount + 1
        End If
Next

IsPrime = (lCount = lValue - 2)

End Function

Rem Enhancement of VBs Val function

Public Function PsyVal(sdata As String, DefaultValue As Long) _
  As Long

'sData: = string data to convert
'DefaultValue = Number to return if conversion fails

'e.g., PsyVal("100", 150) = 100
'      PsyVal("100a", 150) = 150

Dim nSize As Long, nCount As Long
Dim multip As Integer
Dim nResult As Long, nTemp As Long

    PsyVal = DefaultValue
    If Len(sdata) = 0 Then Exit Function
    nResult = 0
    nCount = 1
    If Mid(sdata, nCount, 1) = "-" Then
        multip = -1
        sdata = Right(sdata, Len(sdata) - 1)
    Else
        If Mid(sdata, nCount, 1) = "+" Then
            sdata = Right(sdata, Len(sdata) - 1)
        End If
        multip = 1
    End If
    nSize = Len(sdata)
    Do
        On Error GoTo MayOverFlow
        If IsNumeric(Mid(sdata, 1, 1)) Then
            If Mid(sdata, 1, 1) <> 0 Then
                nTemp = Val(Mid(sdata, 1, 1))
                While nSize > 1
                    nTemp = nTemp * 10
                    nSize = nSize - 1
                Wend
                nResult = nResult + nTemp
            End If
            sdata = Right(sdata, Len(sdata) - 1)
            nSize = Len(sdata)
        Else
            Exit Function
        End If
    Loop Until nSize = 0
    PsyVal = nResult * multip
    Exit Function

MayOverFlow:
    'donothing. Value already set

End Function

Public Function Factorial(ByVal Factor As Byte) As Variant
'******************************************************
'PURPOSE:
'SOLVE FACTORIALS: N!, or N*(N-1)*(N-2)*...*2*1
'Parameter: N
'Returns: Factorial

'EXAMPLE:
'MsgBox Factorial(6)
'Will display 720 because 6 * 5 * 4 * 3 * 2 * 1 = 720

'NOTE: Overflow will occur with any parameter over 170
'******************************************************

On Error GoTo ErrorHandler

If Factor = 0 Then
    Factorial = 1
Else
    Factorial = Factor * Factorial(Factor - 1)
End If
Exit Function

ErrorHandler:
If Err.number = 6 Then 'oveflow
  Err.Raise Err.number, , _
 "Overflow: Number passed to function was too large"

Else 'unknown reason for error, shouldn't occur
    Err.Raise Err.number, , Err.Description
End If
End Function

Rem Finds the simplified nth root of a number

Public Function SimplifyNthRoot(intRoot As Integer, _
   intNumber As Integer) As String
   Dim X As Integer
   For X = Int(intNumber ^ (1 / intRoot)) + 1 To 1 Step -1
      If intNumber Mod (X ^ intRoot) = 0 Then
         SimplifyNthRoot = X & " * " & _
           intNumber / (X ^ intRoot) & "^(1/" & intRoot & ")"
         Exit Function
      End If
   Next X
End Function

Rem HiWord/LoWord Extraction and Combination for Long Integers

Public Function GetLoWord(l As Long) As Integer
    GetLoWord = (l / &H10000) And &HFFFF&
End Function

Public Function GetHiWord(l As Long) As Integer
    GetHiWord = l And &HFFFF&
End Function

Public Function IntegersToLong(LoWord As Integer, _
   HiWord As Integer) As Long

IntegersToLong = (HiWord * &H10000) + (LoWord And &HFFFF&)

End Function

Rem Output Prime Numbers and Prime Factors
Private Sub OutputPrime(iPrime As Double, iNumber As Double, _
                        FileNum As Integer)
                        
'Output the prime number to the debug window, the file, and the textbox

On Error Resume Next
Rem    Debug.Print iPrime & " is prime #" & iNumber
Rem    Print #FileNum, iPrime & " is prime #" & iNumber
    
Rem    txtPrime.text = iPrime
End Sub

Private Sub OutputNonPrime(iNonPrime As Double, _
  arrFactors() As Factor, FileNum As Integer)

'Output the non-prime to the debug window and the file

Dim i As Long       'Current factor
Dim sText As String 'Temp. string to hold text to output

    sText = iNonPrime & " = "
    For i = LBound(arrFactors()) To UBound(arrFactors())
    
        If arrFactors(i).Power = 1 Then
            sText = sText & arrFactors(i).Base
            
        Else
           sText = sText & arrFactors(i).Base & "^" & _
              arrFactors(i).Power
            
        End If '(-> If arrFactors(i).Power = 1
        
        sText = sText & " * "
        
    Next i
    
    sText = Left$(sText, Len(sText) - 3)
    
    Debug.Print sText
    Print #FileNum, sText
End Sub

Private Function IsPrimeNew(iNumber As Double) As Boolean

'Check for a prime number
'
'First, check for even number.
'Since an even number will have a 0 in the
'first binary bit, a simple way to test for
'an even number is to do a bitwise And with
'1, and check for a 0 result.
'
'If the number is odd, we only need to check
'up to Sqr(iNumber) - if iNumber is divisble
'by X, and X > Sqr(iNumber), since
'   X * Sqr(iNumber) > iNumber
'   iNumber / X < Sqr(iNumber)
'Therefore, iNumber = X * Y, Y < Sqr(iNumber),
'so we will already have found Y by the time
'we reach Sqr(iNumber)!
'
'Also, we don't need to check even numbers,
'since if iNumber is not divisible by 2, it will
'not be divisible by any multiple of 2.

Dim ID As Double        'Number to check
Dim iSqr As Double      'Square-root of iNumber
Dim iND As Double       'iNumber / iD
Dim found As Boolean    'iD is a factor?

On Error GoTo ErrHandler
    If (iNumber And 1) = 0 Then
        IsPrimeNew = False
        
    Else
        iSqr = Int(Sqr(iNumber))
        ID = 3
        Do Until ID > iSqr Or found
            iND = iNumber / ID
            If iND = Int(iND) Then
                found = True
                
            Else
                ID = ID + 2
                
            End If '(-> If iND = Int(iND)
            
        Loop '(-> Do Until iD > Sqr Or found
        
        IsPrimeNew = Not (found)
        
    End If '(-> If (iNumber And 1) = 0
    Exit Function
ErrHandler:
    'Stop the program
    Cancel = True
End Function

Private Function Factorize(iNumber As Double) As Factor()

'Factorize iNumber into an array of Factor UDTs
'Note: This will only work in VB6, since it returns an
'array of UDTs. For VB5, modify the code to return the
'array as a parameter:
'
'Private Sub Factorize(iNumber As Double, arrFactors() As Factor)

Dim iCount As Double    'Count of unique factors
Dim iTemp As Double     'iNumber / all factors
Dim ID As Double        'Number to check
Dim iND As Double       'iTemp / iD
Dim arrRes() As Factor  'Temp. array of Factors
Dim iLastFac As Double  'Last factor

    'Allocate space for all factors
    ReDim arrRes(1 To iNumber) As Factor
    
    iTemp = iNumber
    ID = 2
    
    'When iTemp is prime, there are no more factors
    Do Until IsPrime(iTemp)
    
        Do '(-> Loop Until iND = Int(iND)
            iND = iTemp / ID
            If iND = Int(iND) Then
            'iTemp is divisble by iD
                If ID = iLastFac Then
                'Repeated factor - just increase the power
                  arrRes(iCount).Power = arrRes(iCount).Power + 1
                    
                Else
                'New factor
                    iLastFac = ID
                    iCount = iCount + 1
                    arrRes(iCount).Base = ID
                    arrRes(iCount).Power = 1
                    
                End If '(-> If iD = iLastFac
                
                iTemp = iTemp / ID
                
                'NB - Don't increase iD, since it could be a
                'multiple factor of iNumber
                
            Else
                'Increment iD - note that, other than 2, we
                'don't need to check any even factors
                If ID = 2 Then
                    ID = ID + 1
                Else
                    ID = ID + 2
                End If
                
            End If '(-> If iND = Int(iND)
            
        Loop Until iND = Int(iND)
        
    'NB - No need to reset iD, since as iNumber is not divisible
        'by anything less than iD, iTemp won't be divisible by
        'anything less than iD either.
        
    Loop '(-> Do Until IsPrime(iTemp)
    
    If iTemp > 1 Then
    'One last prime factor
        If iTemp = iLastFac Then
            arrRes(iCount).Power = arrRes(iCount).Power + 1
            
        Else
            iCount = iCount + 1
            arrRes(iCount).Base = iTemp
            arrRes(iCount).Power = 1
            
        End If '(-> If iTemp = iLastFac
        
    End If '(-> If iTemp > 1
    
    'Resize array to only include valid factors
    ReDim Preserve arrRes(1 To iCount) As Factor
    
    Factorize = arrRes()
End Function

Private Sub FindPrimes(Optional ShowNonPrimes As Boolean = _
   False, Optional StartAt As Double = 2)

'Find all primes, starting at StartAt. If ShowNonPrimes is True,
'will include a factorization of all non-primes.

Dim theNumber As Double         'The number to check
Dim NumberofPrimes As Double    'The count of primes
Dim FileNum As Integer

On Error GoTo ErrHandler
    Listing = True
    
    FileNum = FreeFile
    Open "Primes.txt" For Append As FileNum
    
    Cancel = False
    
    If StartAt <= 2 Then
        '2 is a prime number
        NumberofPrimes = 1
        OutputPrime 2, 1, FileNum
        
        theNumber = 3
        
    ElseIf Int(StartAt) >= 3 Then
        theNumber = Int(StartAt)
        
    Else
        theNumber = 3
        
    End If '(-> If StartAt <= 2
    
    
    Do Until Cancel
        If IsPrime(theNumber) Then
            NumberofPrimes = NumberofPrimes + 1
            OutputPrime theNumber, NumberofPrimes, FileNum
            
        ElseIf ShowNonPrimes Then
            OutputNonPrime theNumber, Factorize(theNumber), _
              FileNum
            
        End If '(-> If IsPrime(theNumber)
        
        If ShowNonPrimes Then
            theNumber = theNumber + 1
            
        Else
            'Only need to check odd numbers
            theNumber = theNumber + 2
            
        End If '(-> If ShowNonPrimes
        
        'Call DoEvents to allow the user to
        'click the cance button
        DoEvents
        
    Loop '(-> Do Until Cancel
    
ErrHandler:
    'Clean up
    Close #FileNum
    
    Listing = False
    
    Rem cmdCancel.Caption = "Start"
End Sub

Rem Pounds to Kilograms Conversion and Vice Versa
Public Function KilogramsToPounds(Kilograms As Single) As Single

'divide number of kilograms by .4536 to obtain number of pounds

    KilogramsToPounds = Round(Kilograms / 0.4536, 4)
    
    
End Function

Public Function PoundsToKilograms(Pounds As Single) As Single

'multiply the number of pounds times .4536 to obtain
'number of Kilograms.
    PoundsToKilograms = Pounds * 0.4536
    
End Function

Rem Return a Triangle's Angle Given the Length of the Sides
Public Property Get AngleAFromThreeSides(a As Double, _
 b As Double, c As Double) As Variant
 
   'returns angle A, the angle opposite from side a, in radians
    AngleAFromThreeSides = Acos((b ^ 2 + c ^ 2 - a ^ 2) _
        / (2 * b * c))
    
End Property

Rem Falta definir o valor do PI
Public Function Acos(Xcos As Variant) As Double
Dim PI
  'pass in the cosine and receive the angle
  'Xcos of 1, or greater, will cause an error in the Acos
  'function.
  
  If Xcos > 1 Or Xcos < -1 Then
    Acos = 0: Exit Function 'This could bite you in the arse!
  ElseIf Xcos = 1 Then
    Acos = 0: Exit Function
  ElseIf Xcos = -1 Then
    Acos = PI / 2: Exit Function
  End If
  
  
  'Function to do arccosine
  Acos = Atn(-Xcos / Sqr(-Xcos * Xcos + 1)) + 2 * Atn(1)

  'alternate definition
  'Acos = Pi / 2 - 2 * Atn(x / (1 + Sqr(Abs(1 - x * x))))

End Function

Rem Round a Number to Any Precision
Public Function RoundTo(ByVal dblNumber As Double, ByVal _
   dblRoundTo As Double) As Double
   
'Examples:
'MsgBox RoundTo(2.231312312312, 0.1) = 2.2
'MsgBox RoundTo(2.261312312312, 0.1) = 2.3

    On Error Resume Next
    RoundTo = Round(dblNumber / dblRoundTo) * dblRoundTo
    
End Function

Rem Round Up or Down to the Nearest Specified Value
Public Function RoundToValue(ByVal nValue, _
    nCeiling As Double, Optional RoundUp As Boolean = True) _
    As Double
    
    Dim Tmp As Integer
    Dim tmpVal
    If Not IsNumeric(nValue) Then Exit Function
    nValue = CDbl(nValue)
    
'Round up to a whole integer -
'Any decimal value will force a round to the next integer.
'i.e. 0.01 = 1 or 0.8 = 1

    tmpVal = ((nValue / nCeiling) + (-0.5 + (RoundUp And 1)))
    Tmp = Fix(tmpVal)
    tmpVal = CInt((tmpVal - Tmp) * 10 ^ 0)
    nValue = Tmp + tmpVal / 10 ^ 0

'Multiply by ceiling value to set RoundtoValue
    RoundToValue = nValue * nCeiling
       
End Function

Rem Return the Decimal Portion of a Number

Public Function DecimalPortion(number As Double) As Double
  Dim lPos As Long

  lPos = InStr(1, number, ".")
  If lPos > 0 Then
    DecimalPortion = Val(Mid(number, lPos))
  Else
    DecimalPortion = 0
  End If

End Function

Rem Solve Pythagorum Theorum Equations
Function PythagorumTheorum(a As String, b As String, c As _
  String) As Variant

   'Make sure to enter "?" for the value that you want to find
    
On Error GoTo ErrorHandler
   If a = "?" Then
      PythagorumTheorum = Sqr(c ^ 2 - b ^ 2)
   ElseIf b = "?" Then
      PythagorumTheorum = Sqr(c ^ 2 + a ^ 2)
   ElseIf c = "?" Then
      PythagorumTheorum = Sqr(a ^ 2 + b ^ 2)
   Else: PythagorumTheorum = "?"
   ' returns a "?" if no "?" was entered for one of the values
   End If
ErrorHandler:

End Function

'***************************************************************
'Here you can find a function that converts integer into Roman
'numerals (I, II, III, IV ...)

'Call the function like this:
'   strRoman = RomanNumerals(intNumber)

'intNumber can be any integer, and strRoman will contain the
'resulting string.

Public Function RomanNumerals(n As Integer) As String
    'Converts n to a roman numeral

    Dim arabic(12) As Integer, roman(12) As String
    Dim i As Integer, Out As String

    arabic(0) = 1000
    arabic(1) = 900
    arabic(2) = 500
    arabic(3) = 400
    arabic(4) = 100
    arabic(5) = 90
    arabic(6) = 50
    arabic(7) = 40
    arabic(8) = 10
    arabic(9) = 9
    arabic(10) = 5
    arabic(11) = 4
    arabic(12) = 1
    roman(0) = "M"
    roman(1) = "CM"
    roman(2) = "D"
    roman(3) = "CD"
    roman(4) = "C"
    roman(5) = "XC"
    roman(6) = "L"
    roman(7) = "XL"
    roman(8) = "X"
    roman(9) = "IX"
    roman(10) = "V"
    roman(11) = "IV"
    roman(12) = "I"

    i = 0

    While n
        While n >= arabic(i)
            n = n - arabic(i)
            Out = Out + roman(i)
        Wend
        i = i + 1
    Wend

    RomanNumerals = Out

End Function

'***************************************************************
'The function Int2Bin converts an integer to binary format
'and returns a string with the binary digits.

'Call it like this:
'   strBin = Int2Bin(1234, True)

'You can replace 1234 with any number that fits in an integer.
'The second parameter specifies whether you want to remove the
'leading zeros or not. Set to true to remove them or false to
'keep them.

Public Function Int2Bin(Num As Integer, blnRemZeros As Boolean) As String
    Dim lNum As Long, sTemp As String, D As Integer

    lNum = Num And &HFFFF&
    sTemp = String$(16, 48)

    For D = 16 To 1 Step -1
      If lNum And 1 Then Mid$(sTemp, D, 1) = "1"
      lNum = lNum \ 2
      If lNum = 0 Then Exit For
    Next D

    If blnRemZeros Then
        Int2Bin = Mid$(sTemp, D)
    Else
        Int2Bin = sTemp
    End If

End Function

'***************************************************************
'Here 's a function that will take a decimal value and convert it
'to a fraction.

'Usage:
'    MsgBox Num2Frac(0.25)
    
'This will display a message box saying "1/4"

'For further information see comments at start of routine.

Public Function GetDecimalSeparator() As Long
    
    Dim strBuffer As String
    
    'Init buffer
    strBuffer = String$(2, 0)
    
    'Get decimal separator
    strBuffer = Left$(strBuffer, GetLocaleInfo(GetUserDefaultLCID, LOCALE_SDECIMAL, strBuffer, Len(strBuffer)) - 1)
    
    GetDecimalSeparator = Asc(strBuffer)
    
    'Assume dot (.) if function failed
    If GetDecimalSeparator = 0 Then GetDecimalSeparator = Asc(".")
    
End Function



Function Num2Frac(dblValue As Double, Optional Repdigit As Integer = 0) As String
    On Error GoTo FracErr
    
    'Converts a decimal number to a fraction.
    'If you provide a repeating digit decimal, such as 1.3333,
    'also specify the optional Repdigit argument to indicate
    'how many digits repeat. For example, Num2Frac(1.3333)
    'returns 1 3333/10000, while Num2Frac(1.3333, 4) returns
    '1 1/3.

    Dim RepValue As Double, Temp As String
    Dim intValue As String, decValue As String
    Dim X As Long, strNumerator As String, strDenominator As String
    Dim lngDecSep As Long
    
    'Retrieve system decimal separator
    lngDecSep = GetDecimalSeparator
    
    intValue = CStr(dblValue)
    X = InStr(1, intValue, Chr$(lngDecSep))
    If X = 0 Then X = InStr(1, intValue, Asc("."))
    
    If X Then
        Temp = Left(intValue, X - 1)
        decValue = Right(intValue, Len(intValue) - X)
        intValue = Temp

        If Repdigit > 0 Then
        
            If Len(decValue) < Repdigit Then
                decValue = decValue & String(Repdigit - Len(decValue), "0")
            End If
            
            decValue = "." & decValue
            RepValue = 10 ^ Repdigit
            Temp = RepValue * CDbl(decValue & Right(decValue, Repdigit))
            Temp = Temp - CDbl(decValue)
            RepValue = RepValue - 1
            X = GCF(CLng(RepValue * 2), CLng(Temp * 2))
            strNumerator = CLng(Temp * 2) / X
            strDenominator = RepValue * 2 / X
            
        Else
            X = GCF(CLng("1" & String(Len(decValue), "0")), CLng(decValue))
            strNumerator = CLng(decValue) / X
            strDenominator = CLng("1" & String(Len(decValue), "0")) / X
        End If
        
        Num2Frac = IIf(intValue = "0", "", intValue) & " " & strNumerator & "/" & strDenominator
        
    Else
        Num2Frac = intValue
    End If
    
    Exit Function

FracErr:
    Num2Frac = "#Error " & Err.number
    Exit Function
    
End Function

Function GCF(ByVal X As Long, ByVal Y As Long) As Long
    
    ' Returns the Greatest Common Factor
    ' i.e. the largest number which will evenly divide into both X and Y
    
    Dim Temp As Long
    
    X = Abs(X) 'Make both numbers positive
    Y = Abs(Y)
    Temp = X Mod Y
    
    Do While Temp > 0
        X = Y
        Y = Temp
        Temp = X Mod Y
    Loop
    
    GCF = Y
    
End Function

'***************************************************************
'There are two functions below. The first one, FeetInches, converts inches
'to feet and inches, with fractions.
'Example:
'    MsgBox FeetInches(38.5)
    
'This will show a message box saying 3' 2 1/2" (3 feet, two-and-a-half inches)
'The function rounds fractions to the nearest 64th of an inch.

'The other function does the opposite, it takes the output of FeetInches and
'converts it back into inches only.
'Example:
'    MsgBox TotalInches(FeetInches(38.5))

'This will produce a message box saying 38.5.

Public Function FeetInches(ByVal Num) As String
    Dim Recip As Single, Denom As Long
    Dim Out As String
    Const Tick = 64   '1/Tick = Smallest allowable increment
    
    Num = CSng(Num)
    If Num > 12 Then
        Denom = Num \ 12
        Num = Num - (Denom * 12)
        Out = CStr(Denom) & "'"
    End If
    
    If Num > 0 Then
        If Num >= 1 Then
            Out = Out & Space(1) & CStr(Int(Num))
            Num = Num - Int(Num)
        End If
        If Num > 0 Then
            Recip = Int(Num * Tick)
            Denom = Tick
            Do While (Recip And 1!) = 0
                Recip = Recip \ 2
                Denom = Denom \ 2
            Loop
            Out = Out & Space(1) & CStr(Recip) & "/" & CStr(Denom)
        End If
        Out = Out & """"
    End If
    
    FeetInches = Out
    
End Function


Public Function TotalInches(NewValue) As Single
    Dim Tmp As Long, Total As Single
    Dim Pos As Long, Measured As String

    Measured = Trim$(CStr(NewValue))
    Pos = InStr(Measured, "'")
    
    If Pos Then
        Tmp = Val(Measured)
        Total = Tmp * 12
        Measured = Trim(Mid$(Measured, Pos + 2))
    End If
    
    Pos = InStr(Measured, " ")
    
    If Pos Then
        Tmp = Val(Left$(Measured, Pos))
        Total = Total + Tmp
        Measured = Mid$(Measured, Pos + 1)
    End If
    
    Tmp = Val(Measured)
    Pos = InStr(Measured, "/")
    
    If Pos Then
        Total = Total + CSng(Tmp / Val(Mid$(Measured, Pos + 1)))
    Else
        Total = Total + Tmp
    End If
    TotalInches = Total
    
End Function


