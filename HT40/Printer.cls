VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private psLastError As String

Public Sub PrintText(sText As String, Optional NextOnSameLine As Boolean = False)
    On Error GoTo ErrorHandler
    If NextOnSameLine Then
        Printer.Print sText;
    Else
        Printer.Print sText
    End If

ErrorHandler:
    psLastError = Err.Description
    
End Sub
Public Sub DrawCircle(cX As Single, cY As Single, cRadius As Single, _
Optional cStart As Single, Optional cEnd As Single, Optional cAspect As _
Single = 1)
   On Error GoTo ErrorHandler
    Printer.Circle (cX, cY), (cRadius), , cStart, cEnd, cAspect

ErrorHandler:
    psLastError = Err.Description
End Sub
Public Sub DrawLine(lX1 As Single, lY1 As Single, Optional lX2 As Single, _
Optional lY2 As Single, Optional lBox As Boolean = False, Optional lFill As _
Boolean = False)
   On Error GoTo ErrorHandler
    If IsNull(lX2) Then
        If lBox Then
            If lFill Then
                Printer.Line -(lX1, lY1), , BF
            Else
                Printer.Line -(lX1, lY1), , B
            End If
        Else
                    Printer.Line -(lX1, lY1)
        End If
    Else
        If lBox Then
            If lFill Then
                Printer.Line (lX1, lY1)-(lX2, lY2), , BF
            Else
                Printer.Line (lX1, lY1)-(lX2, lY2), , B
            End If
        Else
            Printer.Line (lX1, lY1)-(lX2, lY2)
        End If
    End If

ErrorHandler:
    psLastError = Err.Description

End Sub
Public Sub OutputText()
   On Error GoTo ErrorHandler
    Printer.EndDoc

ErrorHandler:
    psLastError = Err.Description

End Sub
Public Sub NewPage()
   On Error GoTo ErrorHandler
    Printer.NewPage

ErrorHandler:
    psLastError = Err.Description

End Sub
Public Property Let Default(ByVal NewValue As Integer)
   On Error GoTo ErrorHandler
    Set Printer = Printers(NewValue)

ErrorHandler:
    psLastError = Err.Description

End Property
Public Property Let FontBold(ByVal NewValue As Boolean)
   On Error GoTo ErrorHandler
    Printer.FontBold = NewValue

ErrorHandler:
    psLastError = Err.Description

End Property
Public Property Let FontUnderline(ByVal NewValue As Boolean)
   On Error GoTo ErrorHandler
    Printer.FontUnderline = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let FontItalic(ByVal NewValue As Boolean)
   On Error GoTo ErrorHandler
    Printer.FontItalic = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let FontName(ByVal NewValue As String)
   On Error GoTo ErrorHandler
    Printer.FontName = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let FontSize(ByVal NewValue As Single)
   On Error GoTo ErrorHandler
    Printer.FontSize = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let Zoom(ByVal NewValue As Long)
   On Error GoTo ErrorHandler
    Printer.Zoom = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let CurrentX(ByVal NewValue As Integer)
   On Error GoTo ErrorHandler
    Printer.CurrentX = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let CurrentY(ByVal NewValue As Integer)
   On Error GoTo ErrorHandler
    Printer.CurrentY = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Let Orientation(ByVal NewValue As Integer)
   On Error GoTo ErrorHandler
    Printer.Orientation = NewValue

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get Orientation() As Integer
   On Error GoTo ErrorHandler
    Orientation = Printer.Orientation

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get FontBold() As Boolean
   On Error GoTo ErrorHandler
    FontBold = Printer.FontBold

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get FontUnderline() As Boolean
   On Error GoTo ErrorHandler
    FontUnderline = Printer.FontUnderline

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get FontItalic() As Boolean
   On Error GoTo ErrorHandler
    FontItalic = Printer.FontItalic

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get Zoom() As Long
   On Error GoTo ErrorHandler
    Zoom = Printer.Zoom

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get Fonts(Num As Integer) As String
   On Error GoTo ErrorHandler
    Fonts = Printer.Fonts(Num)

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get DeviceName() As String
   On Error GoTo ErrorHandler
    DeviceName = Printer.DeviceName

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get hDC() As Long
   On Error GoTo ErrorHandler
    hDC = Printer.hDC

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get Page() As Integer
   On Error GoTo ErrorHandler
    Page = Printer.Page

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get CurrentX() As Integer
   On Error GoTo ErrorHandler
    CurrentX = Printer.CurrentX

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get CurrentY() As Integer
   On Error GoTo ErrorHandler
    CurrentY = Printer.CurrentY

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get PrinterCount() As Integer
   On Error GoTo ErrorHandler
    PrinterCount = Printers.Count

ErrorHandler:
    psLastError = Err.Description
End Property
Public Property Get PrinterName(Num As Integer) As String
   On Error GoTo ErrorHandler
    PrinterName = Printers(Num).DeviceName

ErrorHandler:
    psLastError = Err.Description
End Property


Public Property Get LastError()
    LastError = psLastError


End Property


