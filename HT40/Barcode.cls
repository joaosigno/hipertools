VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Barcode"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' BarCode.cls
'-------------------------------------------------------------------------
' Descricao   : Monta Codigo de Barra
' Criaçao     : 08/03/1998 6:39:00PM
' Local       : Brasilia/DF
' Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
' Versao      : 1.0.0
' Copyright   : 1997, 1998, 1999 by Zevallos(r) Tecnologia em Informacao
'-------------------------------------------------------------------------

Option Explicit

Const conBB0 = "/img/barcode/barcode0.gif"
Const conBB1 = "/img/barcode/barcode1.gif"
Const conBW0 = "/img/barcode/barcode2.gif"
Const conBW1 = "/img/barcode/barcode3.gif"

Const conStart = "/img/barcode/barcode8.gif"
Const conStop = "/img/barcode/barcode9.gif"

Private sstrComments As String, _
    sstrTitle As String, _
    sstrVersion As String, _
    sstrCompanyName As String, _
    sstrProductName As String, _
    sstrLegalCopyright As String

Private Request As Request, _
        Response As Response, _
        Server As Server, _
        Session As Session, _
        Application As Application, _
        Strings As HiperTools30.String
    
Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  TimeBombX xtbBarcode, 15656, 12934, 3080, 14624
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  Set Strings = New HiperTools30.String
  
End Sub

Public Sub OnEndPage()

  Set Request = Nothing
  Set Response = Nothing
  Set Server = Nothing
  Set Session = Nothing
  Set Application = Nothing
  Set Strings = Nothing
End Sub

Private Sub Class_Initialize()

    sstrComments = App.Comments
    sstrTitle = App.Title
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

End Sub

Rem -------------------------------------------------------------------------
Rem Mostra o Codigo 2 of 5
Rem -------------------------------------------------------------------------
Public Sub ShowI2o5(ByVal strValue As String)
  
  Dim intValueLen As Long, intCount As Long
  Dim strResult As String
  Dim strFirst As String, strSecond As String
  
  strResult = ""
  
  strValue = Strings.NormalizeString(strValue)
  
  intValueLen = Len(strValue)
  
  intCount = 1
  
  
  Do
    strFirst = Mid(strValue, intCount, 1)

    intCount = intCount + 1

    If intCount <= intValueLen Then
      strSecond = Mid(strValue, intCount, 1)
      
      intCount = intCount + 1
      
    Else
      strSecond = "0"
      
    End If

    strResult = strResult & MountBarCodeI2o5(strFirst, strSecond)
        
  Loop While intCount < intValueLen
  
  Response.Write MountImageI5o5("Start", "") & strResult & MountImageI5o5("Stop", "")
  
End Sub

Rem -------------------------------------------------------------------------
Rem Mostra o Codigo 2 of 5
Rem -------------------------------------------------------------------------
Private Function MountBarCodeI2o5(ByVal strFirst As String, _
                                  ByVal strSecond As String) As String
  Dim strResult
  Dim i
  
  strFirst = GetPacketI5o5(strFirst)
  strSecond = GetPacketI5o5(strSecond)
  
  For i = 1 To 5
    strResult = strResult & MountImageI5o5(Mid(strFirst, i, 1), "B")
    strResult = strResult & MountImageI5o5(Mid(strSecond, i, 1), "W")
    
  Next

  MountBarCodeI2o5 = strResult

End Function

Rem -------------------------------------------------------------------------
Rem Monta a imagem
Rem -------------------------------------------------------------------------
Private Function MountImageI5o5(ByVal strBit As String, _
                                ByVal strType As String) As String
  
  Dim strResult
  Dim strTest
  
  strTest = strType & strBit
  
  Select Case strTest
    Case "B0"
      strResult = Session("htDefaultHiperToolsWorkPath") & conBB0
      
    Case "B1"
      strResult = Session("htDefaultHiperToolsWorkPath") & conBB1

    Case "W0"
      strResult = Session("htDefaultHiperToolsWorkPath") & conBW0
      
    Case "W1"
      strResult = Session("htDefaultHiperToolsWorkPath") & conBW1
      
    Case "Start"
      strResult = Session("htDefaultHiperToolsWorkPath") & conStart

    Case "Stop"
      strResult = Session("htDefaultHiperToolsWorkPath") & conStop
      
    Case Else
      
   End Select

  MountImageI5o5 = "<IMG SRC=" & strResult & ">"
  
  Rem " ALT=" & strBit &

End Function

Rem -------------------------------------------------------------------------
Rem Monta a imagem
Rem -------------------------------------------------------------------------
Private Function GetPacketI5o5(ByVal strValue As String) As String
    
  Select Case strValue
    Case "1"
      GetPacketI5o5 = "10001"

    Case "2"
      GetPacketI5o5 = "01001"
     
    Case "3"
      GetPacketI5o5 = "11000"
     
    Case "4"
      GetPacketI5o5 = "00101"
     
    Case "5"
      GetPacketI5o5 = "10100"
     
    Case "6"
      GetPacketI5o5 = "01100"
     
    Case "7"
      GetPacketI5o5 = "00011"
     
    Case "8"
      GetPacketI5o5 = "10010"
     
    Case "9"
      GetPacketI5o5 = "01010"
     
    Case "0"
      GetPacketI5o5 = "00110"

    Case "Start"
      GetPacketI5o5 = "0000"

    Case "Stop"
      GetPacketI5o5 = "100"
     
    Case Else
      
   End Select
   
End Function
