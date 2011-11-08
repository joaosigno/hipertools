<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<!--#INCLUDE FILE="Sugestions.inc"-->
<!--#INCLUDE FILE="Sugestions_Data.inc"-->
<!--#INCLUDE FILE="Sugestions_Form.inc"-->
<%
REM =========================================================================
REM /Sugest.asp
REM -------------------------------------------------------------------------
REM Nome     : Sugestões HiperTools
REM Descricao:
REM Home     : www.zevallos.com.br/sugest
REM Criacao  : 2000/04/27 12:18AM
REM Autor    : Ridai Govinda <ridai@zevallos.com.br>
REM Versao   : 1.0.0
REM Local    :  - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

  Main

REM =========================================================================
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main

  Server.ScriptTimeOut = conScriptTimeout
  Session.TimeOut      = conSessionTimeout

  MainBody

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main
REM =========================================================================

REM =========================================================================
REM Procedimento que mostra a primeira página
REM -------------------------------------------------------------------------
Private Sub ShowFirstPage

  Table.Padding = 0
  Table.Spacing = 0

  Table.BeginTable "45%", "Sugestões para o HiperTools 30", 2, True

     Table.ColumnColSpan = 0

     Table.ColumnColSpan = 2

    Table.BeginRow 2, True
      Table.ColumnAlign = "CENTER"
      Table.BeginColumn

        Show.HTMLCR "Painel de Controle"
        Show.BR
        Show.HTMLCR "Tabelas: "

      Table.EndColumn
    Table.EndRow
    Table.BeginRow 2
      Table.BeginColumn
        URL.BeginURL
          URL.Target = "Body"
          URL.Add efQueryStrAction, efQSActionEditor
          URL.Add efQueryStrWhat, conWhatBugs

          URL.Show "Bugs", efQueryStrEditableStr
          Show.Nbsp 4

          URL(efQueryStrWhat) = conWhatIdeias
          URL.Show "Idéias", efQueryStrEditableStr
          Show.Nbsp 4

          URL(efQueryStrWhat) = conWhatMelhorias
          URL.Show "Melhorias", efQueryStrEditableStr
          Show.BR

          URL(efQueryStrWhat) = conWhatClasses
          URL.Show "Classes", efQueryStrEditableStr
          Show.Nbsp 4

          URL(efQueryStrWhat) = conWhatPosition
          URL.Show "Posição das Classes", efQueryStrEditableStr
          Show.BR

          URL(efQueryStrWhat) = conWhatPessoas
          URL.Show "Pessoas", efQueryStrEditableStr

        URL.EndURL

      Table.EndColumn
      Table.ColumnAlign = ""
    Table.EndRow
    Table.BeginRow 2, True
    Table.Column "&nbsp;"
    Table.EndRow

  Table.EndTable

End Sub
REM -------------------------------------------------------------------------
REM Fim do ShowFirstPage
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  Edit.QueryString
  Edit.DebugMode = True
  REM Edit.ShowSQLQuery = True

  FormatTable

  'Descomente para apagar e recriar as tabelas:
  REM CriaTodasTabelas

  Select Case Edit.parWhat

    Case conWhatBugs
         EditFormBugs

    Case conWhatClasses
         EditFormClasses

    Case conWhatPosition
         EditFormClassPosition

    Case conWhatIdeias
         EditFormIdeias

    Case conWhatMelhorias
         EditFormMelhorias

    Case conWhatPessoas
         EditFormPessoas

    Case Else

  End Select

  Edit.RedirectActions

  Default.HTMLBegin
  Default.HeadAll "Sugestões - HiperTools30"
    Show.HTML "<STYLE>"
    Show.HTML "A:LINK{font-style: normal; text-decoration: none; color:#ff5d05;font-size: 12}"
    Show.HTML "A:ACTIVE{font-style: normal; text-decoration: none; color:red;font-size: 12}"
    Show.HTML "A:VISITED{font-style: normal; text-decoration: none; color:#000000;font-size: 12}"
    Show.HTML "A:HOVER{font-style: normal; text-decoration: underline; color:Green;font-size: 12}"
    Show.HTML "</STYLE>"
  Default.BodyBegin

  Show.BR
  Show.Center

	If ( Edit.IsMyAction ) Then

    ShowFirstPage

	End If

  Default.PageFooterDefault
  Show.CenterEnd

  Default.BodyEnd
  Default.HTMLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do ZeusIII.asp
REM =========================================================================
%>