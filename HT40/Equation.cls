VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Equation"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//
'// An equation evaluator class
'//
'// Description:
'// Evaluates mathematical expressions. All the standard mathematical
'// functions are included (sin, cos etc.). You can also add your own
'// functions.
'//
'// ***************************************************************
'// *  Go to Dragon's VB Code Corner for more useful sourcecode:  *
'// *  http://personal.inet.fi/cool/dragon/vb/                    *
'// ***************************************************************
'//
'// Author of this module: Unknown. Does anyone know who created this
'// great class?
'//
'
' It does a significant amount of work in the
' parsing of an equation, so it's more efficient
' when solving the same equation several times.
'
' The equation is not case sensitive.
'
'
' 1-1-96:  A Bug related to determining the difference between
'          a negative sign and negation was fixed. (And a priority
'          level PRI_NEG was added.)  - TPA
'
' 5-19-97: Updated to VB5.0.

'          Added user defined equations. (VB4.0 didn't allow calls
'          from within a function to the same function in another
'          instance of the class (Each call apparently used the
'          same local variable space.))
'
'          Minor bug fixes.
'
' 7-25-97: Increased the precision of the two conversion constants
'          DEG_TO_RAD and RAD_TO_DEG.  Also increased the precision
'          of the constant e.
'
'          Added the EQ_CLOSE_PAREN flag to the parse routine so that
'          parts in parenthesis would be treated as a number for
'          calculation, but so they aren't mistaken as a number. ie.
'          Fixes a bug when calculating things like (4+3)+23.
'          Also removed the isNeg flag since isNeg was always true
'          when t <> EQ_NONE after the EQ_CLOSE_PAREN flag was added.
'
           
Public Enum EquationErrors
   EquError_UnbalancedParen = 1100 ' Unbalanced parenthesis
   EquError_UnknownFunction        ' Unknown function
   EquError_UnknownVariable        ' Unknown variable
   EquError_InvalidEqu             ' Invalid Equation
   EquError_InvalidArg             ' Invalid argument to function
End Enum


Private Dirty As Boolean
Private Parsed As Boolean

Private Vars As Collection
Private Equs As Collection
Private Equ As String
Private Deg As Boolean

Private dAnswer As Double
Private EquParsed As Collection   'The parsed equation
Private EquOrder  As Collection   'Order in which to solve the equation


' Constants used in parsing
' Priority levels
Private Const PRI_ADD = 1
Private Const PRI_MOD = 2
Private Const PRI_MUL = 3
Private Const PRI_NEG = 4
Private Const PRI_EXP = 5
Private Const PRI_VAR = 6
Private Const PRI_PAR = 7
Private Const PRI_LEVEL = 7

Private Const EQ_NONE = 0
Private Const EQ_STRING = 1
Private Const EQ_NUMBER = 2
Private Const EQ_CLOSE_PAREN = 3

Private Const ER_NONE = 0
Private Const ER_VAR = 1

Private Const PI = 3.14159265358979
Private Const DEG_TO_RAD = 1.74532925199433E-02
Private Const RAD_TO_DEG = 57.2957795130824

Public Property Let Degrees(b As Boolean)
   If b <> Deg Then
      Deg = b
      Dirty = True
   End If
End Property

Public Property Get Degrees() As Boolean
   Degrees = Deg
End Property

Public Property Let Equation(e As String)
   Parsed = False
   Dirty = True
   Equ = LCase(e)
End Property

Public Property Get Equation() As String
   Equation = Equ
End Property

Public Property Set UserEqu(Name As String, eq As Equation)
   On Error Resume Next
   Dirty = True
   Equs.Remove Name & "("
   Equs.Add eq, Name & "("
End Property

Public Property Get UserEqu(Name As String) As Equation
   On Error GoTo GetError
   
   Set UserEqu = Equs(Name)
   Exit Property
   
GetError:
   Set UserEqu = Nothing
End Property

Public Sub UserEquClear()
   Set Equs = New Collection
   Dirty = True
End Sub

Public Sub UserEquRemove(Name As String)
   On Error Resume Next
   Equs.Remove Name & "("
   Dirty = True
End Sub

Public Sub VarClear()
   Set Vars = New Collection
   Dirty = True
End Sub

Public Sub VarRemove(Name As String)
   On Error Resume Next
   Vars.Remove Name
   Dirty = True
End Sub

Public Property Get Var(Name As String) As Double
   On Error GoTo GetError
   
   Var = CDbl(Vars(Name))
   Exit Property
   
GetError:
   Var = 0#
End Property

Public Property Let Var(Name As String, ByVal Num As Double)
   On Error Resume Next
   Dirty = True
   Vars.Remove Name
   Vars.Add Num, Name
End Property

' Internal search function...
Private Function GetRight(ByVal j As Long, v() As Variant) As Long
   Dim i As Long
   
   For i = j + 1 To UBound(v)
      If Not IsNull(v(i)) Then
         GetRight = i
         Exit Function
      End If
   Next i
   GetRight = 0
End Function
' Internal search function
Private Function GetLeft(ByVal j As Long, v() As Variant) As Long
   Dim i As Long
   
   For i = j - 1 To 1 Step -1
      If Not IsNull(v(i)) Then
         GetLeft = i
         Exit Function
      End If
   Next i
   GetLeft = 0
End Function

Public Function Value() As Double
   If Dirty Then
      Solve
   End If
   
   Value = dAnswer
End Function

Private Sub Class_Initialize()
   Dirty = False
   Parsed = True
   Degrees = False
   Set Vars = New Collection
   Set Equs = New Collection
End Sub

Private Sub Class_Terminate()
   Set Vars = Nothing
   Set Equs = Nothing
   Set EquParsed = Nothing
   Set EquOrder = Nothing
End Sub

Private Sub Parse()
   Dim i As Integer
   Dim s As String
   Dim t As Integer
   Dim j As Integer
   Dim sTmp As String
   Dim p As Integer
   Dim EquPriority As New Collection
   Dim maxPriority
   
   s = ""
   t = EQ_NONE
   j = 1
   p = 0
   Set EquParsed = New Collection
   
   EquParsed.Add ""
   EquPriority.Add ""
   maxPriority = PRI_LEVEL
   
   For i = 1 To Len(Equ)
      sTmp = Mid$(Equ, i, 1)
      
      Select Case sTmp
      Case "A" To "Z", "a" To "z", "_"
         If t = EQ_NONE Or t = EQ_CLOSE_PAREN Then
            t = EQ_STRING
            s = sTmp
         ElseIf t = EQ_NUMBER Then
            t = EQ_STRING
            EquParsed.Add s, , j
            EquPriority.Add 0, , j
            j = j + 1
            EquParsed.Add "*", , j
            EquPriority.Add PRI_MUL + p, , j
            j = j + 1
            s = sTmp
         Else
            s = s + sTmp
         End If
         
      Case "1" To "9", "0", "."
         If t = EQ_NONE Or t = EQ_CLOSE_PAREN Then
            t = EQ_NUMBER
            s = sTmp
         Else
            s = s + sTmp
         End If
      
      Case "(":
         If t = EQ_STRING Then
            EquParsed.Add s + sTmp, , j
            EquPriority.Add p + PRI_PAR, , j
            j = j + 1
            s = ""
         ElseIf t = EQ_NUMBER Then
            EquParsed.Add s, , j
            EquPriority.Add 0, , j
            j = j + 1
            EquParsed.Add "*", , j
            EquPriority.Add p + PRI_MUL, , j
            j = j + 1
            EquParsed.Add sTmp, , j
            EquPriority.Add p + PRI_PAR, , j
            j = j + 1
            s = ""
         Else
            EquParsed.Add sTmp, , j
            EquPriority.Add p + PRI_PAR, , j
            j = j + 1
         End If
         
         p = p + PRI_LEVEL
         t = EQ_NONE
         
         If maxPriority < p + PRI_LEVEL Then
            maxPriority = p + PRI_LEVEL
         End If
      
      Case "*", "/":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         EquParsed.Add sTmp, , j
         EquPriority.Add p + PRI_MUL, , j
         j = j + 1
         t = EQ_NONE
      
      Case "\":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         EquParsed.Add sTmp, , j
         EquPriority.Add p + PRI_MUL, , j
         j = j + 1
         t = EQ_NONE
      
      Case "+":
         If t <> EQ_NONE Then
            If t <> EQ_CLOSE_PAREN Then
                EquParsed.Add s, , j
                EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                j = j + 1
                s = ""
            End If
            EquParsed.Add sTmp, , j
            EquPriority.Add p + PRI_ADD, , j
            j = j + 1
            t = EQ_NONE
         Else
            'Ignore things like "(+1)"
            t = EQ_NONE
         End If
      
      Case "-":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         'If we are preceded by a number, variable, or a closed
         'paren then we are a minus sign.
         If t <> EQ_NONE Then
            EquParsed.Add sTmp, , j
            EquPriority.Add p + PRI_ADD, , j
            j = j + 1
            t = EQ_NONE
         Else ' we are a negation sign
            EquParsed.Add "~", , j
            EquPriority.Add p + PRI_NEG, , j
            j = j + 1
            t = EQ_NONE
         End If
         
      Case "^":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         EquParsed.Add sTmp, , j
         EquPriority.Add p + PRI_EXP, , j
         j = j + 1
         t = EQ_NONE
         
      Case "%":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         EquParsed.Add sTmp, , j
         EquPriority.Add p + PRI_MOD, , j
         j = j + 1
         t = EQ_NONE
         
      Case ",":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         EquParsed.Add Null, , j
         EquPriority.Add 0, , j
         j = j + 1
         t = EQ_NONE

      Case ")":
         If t <> EQ_NONE And t <> EQ_CLOSE_PAREN Then
            EquParsed.Add s, , j
            EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
            j = j + 1
            s = ""
         End If
         
         EquParsed.Add sTmp, , j
         EquPriority.Add p - (PRI_LEVEL - PRI_PAR), , j
         p = p - PRI_LEVEL
         j = j + 1
         t = EQ_CLOSE_PAREN
      End Select
   Next i
   
   If s <> "" Then
      EquParsed.Add s, , j
      EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
      j = j + 1
   End If
   
   EquParsed.Remove j
   EquPriority.Remove j
   
   If p <> 0 Then
      Err.Raise EquError_UnbalancedParen, "Equation", "Unbalanced parenthesis"
      Exit Sub
   End If
   
       'Debugging section...
      'For i = 1 To EquParsed.Count
      '   Debug.Print EquParsed(i) & ";";
      'Next i
      'Debug.Print
      '   For i = 1 To EquPriority.Count
      '   Debug.Print EquPriority(i) & ";";
      'Next i
      'Debug.Print
      'Debug.Print "MaxPriority = " & maxPriority
      ' End Debugging section....
   
   Set EquOrder = New Collection
   EquOrder.Add ""
   
   For j = 1 To maxPriority
      For i = EquPriority.Count To 1 Step -1
         If EquPriority(i) = j Then
            EquOrder.Add i, , , 1
         End If
      Next i
   Next j
   
   EquOrder.Remove 1
   
   'For i = 1 To EquOrder.Count
   '   Debug.Print EquOrder(i) & ";";
   'Next i
   'Debug.Print
   
   Parsed = True
End Sub

Public Sub Solve()
   Dim i As Long
   Dim j As Long
   Dim l As Long
   Dim r As Long
   Dim m As Long
   Dim n As Long
   Dim X As Double
   Dim Y As Double
   Dim v As Variant
   Dim eSpace As Integer
   Dim Temp() As Variant
   Dim f As Equation
   Dim j2 As Long  ' debug variable
   
   On Error GoTo SolveError
   
   If Not Parsed Then
      Parse
   End If
      
   ' Copy the equation to a working array
   ReDim Temp(1 To EquParsed.Count)
   
   For i = 1 To EquParsed.Count
      Temp(i) = EquParsed(i)
   Next
   
   eSpace = ER_NONE
   
   ' Solve the equation
   For i = 1 To EquOrder.Count
      'Debug.Print "Pro -> " & EquOrder(i) & " = ";
      'For j2 = 1 To UBound(Temp)
      '   Debug.Print Temp(j2) & ";";
      'Next j2
      'Debug.Print
      
      m = EquOrder(i)
      v = Temp(m)
      
      Select Case v
      ' Standard operators
      Case "~"  'Negative operator (inserted by the parser)
         r = GetRight(m, Temp)
         Temp(m) = -CDbl(Temp(r))
         Temp(r) = Null
         
      Case "*"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) * CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
         
      Case "/"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) / CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
         
      Case "\"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) \ CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
         
      Case "+"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) + CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
    
      Case "-"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) - CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
    
      Case "^"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) ^ CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
         
      Case "%"
         l = GetLeft(m, Temp)
         r = GetRight(m, Temp)
         Temp(l) = CDbl(Temp(l)) Mod CDbl(Temp(r))
         Temp(r) = Null
         Temp(m) = Null
         
      Case "("
         i = i + 1
         n = EquOrder(i)
         r = GetRight(m, Temp)
         If r >= n Then
            Temp(m) = 0#
            Temp(n) = Null
         Else
            Temp(m) = Temp(r)
            Temp(r) = Null
            Temp(n) = Null
         End If
     
      Case Else
         If Right$(Temp(m), 1) = "(" Then
            'Must be a function
            i = i + 1
            n = EquOrder(i)
            
            l = GetRight(m, Temp)
            r = GetLeft(n, Temp)
            
            If l >= n Then
               Err.Raise EquError_InvalidArg, "clsEquation", "Invalid arguments to function: " & v & ")"
               Exit Sub
            Else
               X = CDbl(Temp(l))
            End If
            
            If r <= m Then
               Err.Raise EquError_InvalidArg, "clsEquation", "Invalid arguments to function: " & v & ")"
               Exit Sub
            Else
               Y = CDbl(Temp(r))
            End If
            
            Temp(r) = Null
            Temp(l) = Null
            Temp(m) = Null
            Temp(n) = Null
            
            Select Case v
               ' Standard functions
               Case "abs("
                  Temp(m) = Abs(X)
                  
               Case "atn("
                  If Degrees Then
                     Temp(m) = Atn(X) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(X)
                  End If
                  
               Case "arctan("
                  If Degrees Then
                     Temp(m) = Atn(X) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(X)
                  End If
                  
               Case "cos("
                  If Degrees Then
                     Temp(m) = Cos(X * DEG_TO_RAD)
                  Else
                     Temp(m) = Cos(X)
                  End If
                  
               Case "exp("
                  Temp(m) = Exp(X)
                  
               Case "fix("
                  Temp(m) = Fix(X)
                  
               Case "int("
                  Temp(m) = Int(X)
                  
               Case "log("
                  Temp(m) = Log(X)
                  
               Case "rnd("
                  Temp(m) = Rnd(X)
                  
               Case "sgn("
                  Temp(m) = Sgn(X)
                  
               Case "sin("
                  If Degrees Then
                     Temp(m) = sIn(X * DEG_TO_RAD)
                  Else
                     Temp(m) = sIn(X)
                  End If
                  
               Case "sqr("
                  Temp(m) = Sqr(X)
                  
               Case "tan("
                  If Degrees Then
                     Temp(m) = Tan(X * DEG_TO_RAD)
                  Else
                     Temp(m) = Tan(X)
                  End If
                  
               ' 2 variable functions
               Case "min("
                  Temp(m) = IIf(X < Y, X, Y)
                  
               Case "max("
                  Temp(m) = IIf(X > Y, X, Y)
                  
               Case "random("
                  Temp(m) = (Rnd * (Y - X)) + X
                  
               Case "mod("
                  Temp(m) = X Mod Y
                  
               Case "logn("
                  Temp(m) = Log(X) / Log(Y)
               
               ' Misc equations
               Case "rand("
                  Temp(m) = Int(Rnd * X)
               
               ' Derived functions
               Case "sec("
                  If Degrees Then
                     Temp(m) = (1 / Cos(X * DEG_TO_RAD))
                  Else
                     Temp(m) = 1 / Cos(X)
                  End If
                  
               Case "cosec("
                  If Degrees Then
                     Temp(m) = (1 / sIn(X * DEG_TO_RAD))
                  Else
                     Temp(m) = 1 / sIn(X)
                  End If
                  
               Case "cotan("
                  If Degrees Then
                     Temp(m) = (1 / Tan(X * DEG_TO_RAD))
                  Else
                     Temp(m) = 1 / Tan(X)
                  End If
                  
               Case "arcsin("
                  If Degrees Then
                     Temp(m) = (Atn(X / Sqr(-X * X + 1))) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(X / Sqr(-X * X + 1))
                  End If
                  
               Case "arccos("
                  If Degrees Then
                     Temp(m) = (Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
                  End If
                  
               Case "arcsec("
                  If Degrees Then
                     Temp(m) = (Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
                  End If
                  
               Case "arccosec("
                  If Degrees Then
                     Temp(m) = (Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
                  End If
                  
               Case "arccotan("
                  If Degrees Then
                     Temp(m) = (Atn(X * DEG_TO_RAD) + 2 * Atn(1)) * RAD_TO_DEG
                  Else
                     Temp(m) = Atn(X) + 2 * Atn(1)
                  End If
                  
               Case "sinh("
                  Temp(m) = (Exp(X) - Exp(-X)) / 2
                  
               Case "cosh("
                  Temp(m) = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
                  
               Case "tanh("
                  Temp(m) = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
                  
               Case "sech("
                  Temp(m) = 2 / (Exp(X) + Exp(-X))
                  
               Case "cosech("
                  Temp(m) = 2 / (Exp(X) - Exp(-X))
                  
               Case "cotanh("
                  Temp(m) = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
                  
               Case "arcsinh("
                  Temp(m) = Log(X + Sqr(X * X + 1))
                  
               Case "arccosh("
                  Temp(m) = Log(X + Sqr(X * X - 1))
                  
               Case "arctanh("
                  Temp(m) = Log((1 + X) / (1 - X)) / 2
                  
               Case "arcsech("
                  Temp(m) = Log((Sqr(-X * X + 1) + 1) / X)
                  
               Case "arccosech("
                  Temp(m) = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
                  
               Case "arccotanh("
                  Temp(m) = Log((X + 1) / (X - 1)) / 2
                  
               Case "log10("
                  Temp(m) = Log(X) / Log(10)
                  
               Case "log2("
                  Temp(m) = Log(X) / Log(2)
                  
               Case "ln("    'A macro to Log
                  Temp(m) = Log(X)
                  
               ' conversion functions
               Case "deg("   ' Radians to degrees
                  Temp(m) = X * RAD_TO_DEG
                  
               Case "rad("   ' Degrees to radians
                  Temp(m) = X * DEG_TO_RAD
                  
               Case Else
                  'check for user defined equ's
                  On Error Resume Next
                  Set f = Equs(v)
                  If Err = 0 Then
                     On Error GoTo SolveError
                     f.Var("x") = X
                     f.Var("y") = Y
                     Temp(m) = f.Value
                  Else
                     Err.Raise EquError_InvalidEqu, "clsEquation", "Undefined Function: " & v & ")"
                     Exit Sub
                  End If
            End Select
         Else
            'Must be a variable
            Select Case v
            Case "pi":
               Temp(m) = PI
               
            Case "e":
               Temp(m) = Exp(1)
                           
            Case "rnd":
               Temp(m) = Rnd
               
            Case Else
               eSpace = ER_VAR
               Temp(m) = CDbl(Vars(Temp(m)))
               eSpace = ER_NONE
            End Select
         End If
      End Select
   Next i
   
   dAnswer = CDbl(Temp(GetRight(0, Temp)))
   Dirty = False
   Exit Sub
   
SolveError:
   Select Case Err
   'Overflow, division by 0, internal errors...
   Case 6, 11, EquError_UnbalancedParen To EquError_InvalidArg
      Err.Raise Err, "clsEquation", Err.Description
   Case 5:
      Select Case eSpace
         Case ER_VAR
            Err.Raise EquError_UnknownVariable, "clsEquation", "Undefined Variable: " & v
         Case Else
            Err.Raise Err, "clsEquation", Err.Description
      End Select
   Case Else
      Err.Raise EquError_InvalidEqu, "clsEquation", "Invalid Equation"
   End Select
End Sub



