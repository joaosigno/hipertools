<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<%
REM =========================================================================
REM /ZNovidades.asp
REM -------------------------------------------------------------------------
REM Nome     : Novidades
REM Descricao: Sistema de Novidades do Site Zevallos
REM Home     : www.zevallos2.com.br/Novidades
REM Criacao  : 3/4/0 7:21PM
REM          : Fernando Aquino (Desenvolvimento)
REM Versao   : 1
REM Local    :  - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

  Const conScriptTimeout  = 15
  Const conSessionTimeout = 300

  Const conPOption = "O"
  Const conPTarget = "T"

  Const conOptionNewsList  = "1"
  Const conOptionNewsUnit  = "2"

  Const conTableSize = "760"

  Dim sparOption
  Dim sparTarget

  Dim sobjRS
  Dim sobjRSAux
  Dim sobjConn

  sparOption = Request.QueryString(conPOption)
  sparTarget = CInt(Request.QueryString(conPTarget))

  If sparTarget <= 0 Then
     sparTarget = 1
  End If

  Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout

  Set sobjRS = Server.CreateObject("ADODB.RecordSet")
  Set sobjRSAux = Server.CreateObject("ADODB.RecordSet")
  Set sobjConn = Server.CreateObject("ADODB.Connection")

  sobjConn.ConnectionTimeout = 300
  sobjConn.CommandTimeout = 300
  sobjConn.Open Session("ConnectionString")

  MainBody
  Server.ScriptTimeOut = Session("ScriptTimeOut")

  Set sobjRS    = nothing
  Set sobjRSAux = nothing
  Set sobjConn  = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Procedimento que monta a lista de tipos de notícias
REM -------------------------------------------------------------------------
Private Sub ShowNewsTypes

  Dim sql

  sql = "SELECT tnvCodigo,tnvNome,tnvDescricao FROM zsnTipoNovidades"
  sobjRS.Open sql, sobjConn, adOpenKeySet, adLockReadOnly

  Table.Padding = 1
  Table.Spacing = 1
  Table.BeginTable "100%"
  Table.CellAlign  = "top"
    Do While Not sobjRS.EOF
      Table.Row Strings.FormatText( "<a href=""$s?O=$s&T=$s"" title=""$s"">$s</a>",  Initializer.ScriptName, _
                                    conOptionNewsList, sobjRS("tnvCodigo"), sobjRS("tnvDescricao"), sobjRS("tnvNome") )
      sobjRS.MoveNext
    Loop
  Table.CellAlign  = ""
  Table.EndTable

  sobjRS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowNewsType
REM =========================================================================

REM =========================================================================
REM Procedimento que monta a lista de notícias
REM -------------------------------------------------------------------------
Private Sub ShowNewsList
  Dim sql

  sql = "SELECT novCodigo,novDataCriacao,novTitulo,novResumo FROM zsnNovidades WHERE novTipoNovidade = " & sparTarget
  sobjRS.Open sql, sobjConn, adOpenStatic, adLockReadOnly

  Table.BeginTable conTableSize, "Novidades", 2, False
  Table.CellVAlign = "top"

    Do While Not sobjRS.EOF

      Table.BeginRow 2
        Table.BeginCell

          If sobjRS("novDataCriacao") > "" Then
             Show.HTML Strings.LongDate(sobjRS("novDataCriacao")) & " - "
          End If

          URL.BeginURL Initializer.ScriptName
            URL.Add conPOption, conOptionNewsUnit
            URL.Add conPTarget, sobjRS("novCodigo")
            URL.Show sobjRS("novTitulo")
          URL.EndURL

          If sobjRS("novResumo") > "" Then
             Show.Message sobjRS("novResumo")
          End If

        Table.EndCell
      Table.EndRow

      sobjRS.MoveNext
    Loop

  Table.CellAlign  = ""
  Table.EndTable

  sobjRS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowNewsList
REM =========================================================================

REM =========================================================================
REM Procedimento que monta o Unit de notícias
REM -------------------------------------------------------------------------
Private Sub ShowNewsUnit

  Dim sql

  sql = "SELECT * FROM zsnNovidades WHERE novCodigo = " & sparTarget
  sobjRS.Open sql, sobjConn, adOpenStatic, adLockReadOnly

  If Not sobjRS.EOF Then

    Table.BeginTable conTableSize, sobjRS("novTitulo"), 3, False
    Table.CellVAlign  = "top"

    Table.BeginRow 2

      Table.BeginCell

        REM -------------------------------------------------------------------------
        REM Imagem
        If sobjRS("novImagem") > "" Then
           Show.Image "img/" & sobjRS("novImagem")
        Else
           Show.Nbsp
        End If

      Table.EndCell

      Table.BeginCell

        REM -------------------------------------------------------------------------
        REM SubTítulo
        If sobjRS("novSubTitulo") > "" Then
           Show.HTML Strings.BoldText(sobjRS("novSubTitulo"))
           Show.Br
           Show.Br
        End If

        REM -------------------------------------------------------------------------
        REM Texto da Notícia
        If sobjRS("novTexto") > "" Then
           Show.HTML "<P ALIGN=""justify"">" & Replace(sobjRS("novTexto"), vbcrlf, "<br>")
           Show.Br
           Show.Br
        Else
           Show.Nbsp
           Show.Br
           Show.Br
        End If

        REM -------------------------------------------------------------------------
        REM Saiba mais
        If sobjRS("novSaibamaisUrl") > "" Then
           Show.HTML "Saiba mais em "
           URL.BeginURL "http://" & sobjRS("novSaibamaisUrl")
             URL.Show sobjRS("novSaibamaisUrl")
           URL.EndURL
           Show.HTML "."
           Show.Br
        End If

      Table.EndCell

      Table.CellWidth = "150"
      Table.BeginCell

      Show.HTML "<FONT SIZE=1>"

        REM -------------------------------------------------------------------------
        REM Autor
        If sobjRS("novAutor") > "" Then

           Show.HTML Strings.BoldText(sobjRS("novAutor"))
           Show.Br

           REM -------------------------------------------------------------------------
           REM Referência
           If sobjRS("novReferencia") > "" Then
              Show.HTML Strings.ItalicText(sobjRS("novReferencia"))
              Show.Br
           End If

           REM -------------------------------------------------------------------------
           REM E-mail
           If sobjRS("novMailAutor") > "" Then
              Show.HTML "E-mail: " & "<A HREF=""mailto:" & Trim(sobjRS("novMailAutor")) & """>" & sobjRS("novMailAutor") & "</A>"
              Show.Br
           End If

           REM -------------------------------------------------------------------------
           REM Referência
           If sobjRS("novUrlReferencia") > "" Then
              Show.HTML "Web: "
              URL.BeginURL "http://" & sobjRS("novUrlReferencia")
                URL.Show sobjRS("novUrlReferencia")
              URL.EndURL
              Show.Br
           End If

        Else
           Show.Nbsp
        End If

        Show.HTML "</FONT>"

      Table.EndCell

      Table.CellWidth = ""

    Table.EndRow

    Table.CellVAlign  = ""
    Table.EndTable

    REM -------------------------------------------------------------------------
    REM Realiza as marcações referentes à acesso do sistema.
    sql = "SELECT * FROM zsnNovidades WHERE novCodigo = " & sparTarget
    sobjRSAux.Open sql, sobjConn, adOpenDynamic, adLockPessimistic

    If Not sobjRSAux.EOF Then
       If sobjRSAux("novAcessos") > "" Then
          sobjRSAux("novAcessos") = sobjRSAux("novAcessos") + 1
       Else
          sobjRSAux("novAcessos") = 1
       End If
       sobjRSAux("novDatUltmAcess") = Now
    End If

    sobjRSAux.Update
    sobjRSAux.Close

    REM Fim do Realiza as marcações referentes à acesso do sistema.
    REM -------------------------------------------------------------------------

  End If

  sobjRS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowNewsUnit
REM =========================================================================

REM =========================================================================
REM Procedimento que retorna o caminho da URL completo sem o nome do arquivo
REM apenas
REM -------------------------------------------------------------------------
Private Function TranslateSiteRoot

  Dim strReverse
  strReverse = Strings.Reverse( Initializer.ScriptURL )
  TranslateSiteRoot = Strings.Reverse( Mid( strReverse, InStr( strReverse, "/" ) ) )

End Function
REM -------------------------------------------------------------------------
REM Fim do TranslateSiteRoot
REM =========================================================================


REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  Default.BodyBGColor         = "white"'"#0079BD"
  Default.BodyBackground      = TranslateSiteRoot & "/img/bgnovidades.gif"
  Default.BodyText            = "#000000"
  Default.BodyLink            = "#0000FF"
  Default.BodyVLink           = "gray" '"#FFFF00"
  Default.BodyALink           = "red"
  Default.LinkStyleSheetHRef  = "/default.css"
  Default.BodyTopMargin       = 0

  Table.Style.BackgroundFormat = tbStFormatNothing
  Table.Style.BorderFormat     = tbBdFormatInvisible
  Table.Style.ColorFormat      = tbStFormatNothing
  Table.Style.BaseColor        = ""

  Default.HTMLBegin
  Default.HeadAll "Novidades"
  Default.BodyBegin

    Table.BeginTable "100%"
    Table.BeginRow

      Table.CellWidth  = "150"
      Table.CellAlign  = "center"
      Table.CellVAlign = "top"
      Table.BeginCell
         ShowNewsTypes
      Table.EndCell

      Table.CellWidth = ""
      Table.BeginCell

          Table.CellAlign = ""
          Table.Padding   = 5
          Table.Spacing   = 5
          If ( sparOption = conOptionNewsUnit ) Then
             ShowNewsUnit
          Else
             ShowNewsList
          End If

      Table.EndCell

    Table.EndRow
    Table.EndTable

  Default.BodyEnd
  Default.HTMLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do ZNovidades.asp
REM =========================================================================
%>
