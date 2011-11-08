<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE VIRTUAL="/HiperTools/HiperTools30.inc"-->
<!--#INCLUDE VIRTUAL="/HiperTools/Objects.inc"-->
<!--#INCLUDE FILE="Chamado.inc"-->
<!--#INCLUDE FILE="Constants.inc"-->
<%

REM =========================================================================
REM  /Chamado.asp
REM -------------------------------------------------------------------------
REM Descricao: Chamado Técnico
REM Criacao  : 2/12/0 5:14PM
REM Autor    : Zevallos Tecnologia em Informacao
REM Versao   : 1.1.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by Zevallos(r) Tecnologia em Informacao
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
REM Mostra a Primeira Pagina
REM -------------------------------------------------------------------------
Private Sub ShowFirstPage

  Table.Padding = 4
  Table.Spacing = 0
  Table.BeginTable "65%", ,2 ,True

  Table.ColumnColSpan = 2
  Table.BeginRow 2, True
    Table.Column "Painel de Controle"
  Table.EndRow

  Table.BeginRow 2
    Table.BeginColumn

      URL.BeginURL Initializer.ScriptName
        URL.Add efQueryStrAction, efQSActionEditor
        URL.Add efQueryStrWhat, conWhatChamaTecnico
        URL.Show "Chamado Técnico", efQueryStrEditableStr
      URL.EndURL

    Table.EndColumn
  Table.EndRow

  Table.BeginRow 2, True
    Table.Column "Relatórios"
  Table.EndRow

  Table.ColumnColSpan = ""
  Table.BeginRow 2
    Table.BeginColumn

      URL.BeginURL Initializer.ScriptName

        URL.Add efQueryStrAction, efQSActionSummary
        URL.Add efQueryStrWhat, conWhatChamaTecnico

        URL.Show "Relatório por Setor", URL.Equal(efQueryStrList, conListSetor)
        Show.BR

        URL.Show "Relatório por Usuário", URL.Equal(efQueryStrList, conListUsuario)
        Show.BR

        URL.Show "Relatório por Tempo de Atendimento", URL.Equal(efQueryStrList, conListTempo)
        Show.BR

        URL.Show "Relatório por Atendente", URL.Equal(efQueryStrList, conListAtendente)
        Show.BR

        URL.Show "Relatório por Tipo de Problema", URL.Equal(efQueryStrList, conListProblema)

      URL.EndURL

    Table.EndColumn
  Table.EndRow

  Table.EndTable

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FistPage
REM =========================================================================

REM =========================================================================
REM  Guarda os dados das tabelas de Chamados Tecnicos
REM -------------------------------------------------------------------------
Public Sub sncChamaTecnico

  Edit.DataBegin "sncChamaTecnico"

  Edit.DataAddField "chtCodigo",          efDataTypeInt,       "4", efRequired
  Edit.DataAddField "chtSolicitante",     efDataTypeVarChar, "100", efRequired
  Edit.DataAddField "chtSetor",           efDataTypeVarChar,  "50", efNotRequired
  Edit.DataAddField "chtDataSolicita",    efDataTypeDateTime, "10", efNotRequired
  Edit.DataAddField "chtHoraSolicita",    efDataTypeVarChar,   "5", efNotRequired
  Edit.DataAddField "chtProblemaTipo",    efDataTypeVarChar,   "1", efNotRequired
  Edit.DataAddField "chtDescricao",       efDataTypeText,   "3000", efNotRequired

  Edit.DataAddField "chtAtendente",       efDataTypeVarChar, "100", efNotRequired
  Edit.DataAddField "chtDataAtende",      efDataTypeDateTime, "10", efNotRequired
  Edit.DataAddField "chtHoraAtende",      efDataTypeVarChar,   "5", efNotRequired
  Edit.DataAddField "chtTempoAtende",     efDataTypeVarChar,   "5", efNotRequired
  Edit.DataAddField "chtProblemaDetecta", efDataTypeText,   "3000", efNotRequired
  Edit.DataAddField "chtSolucao",         efDataTypeText,   "3000", efNotRequired
  Edit.DataAddField "chtPendencias",      efDataTypeText,   "3000", efNotRequired

  Edit.DataAddPrimaryKey "chtCodigo"

  If Not Edit.IsTable( "sncChamaTecnico" ) Then
     Edit.TableCreate "sncChamaTecnico"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub sncChamaTecnico
REM =========================================================================

REM =========================================================================
REM  Guarda os dados do form de Chamados Tecnicos
REM -------------------------------------------------------------------------
Public Sub GetChamaTecnico

  sncChamaTecnico

  Edit.FormBegin "sncChamaTecnico", "Chamado Técnico", 1, Edit.parWhat, efValLocClient

  Edit.FormList "chtSolicitante,chtSetor,chtDataSolicita,chtHoraSolicita,chtProblemaTipo"
  Edit.FormUnit "chtSolicitante;chtSetor,chtDataSolicita,chtHoraSolicita;chtProblemaTipo;chtDescricao|chtAtendente;chtDataAtende,chtHoraAtende,chtTempoAtende;chtProblemaDetecta;chtSolucao;chtPendencias"
  Edit.FormFind "chtSolicitante,chtSetor,chtDataSolicita,chtHoraSolicita,chtProblemaTipo,chtDescricao,chtAtendente,chtDataAtende,chtHoraAtende,chtTempoAtende,chtProblemaDetecta,chtSolucao,chtPendencias"
  Edit.FormTabs "Solicitação|Atendimento"

  Edit.AddList "chtSetor,#1(l=Nº,gc)", "Por Setor", conListSetor
  Edit.AddList "chtSolicitante,#1(l=Nº,gc)", "Por Usuário", conListUsuario
  Edit.AddList "chtTempoAtende,#1(l=Nº,gc)", "Por Tempo de Atendimento", conListTempo
  Edit.AddList "chtAtendente,#1(l=Nº,gc)", "Por Atendente", conListAtendente
  Edit.AddList "chtProblemaTipo,#1(l=Nº,gc)", "Por Tipo de Problema", conListProblema

  Edit.AddField "chtCodigo",       "Código",  efFldTypeText, , , efNext
  Edit.AddField "chtSolicitante",  "Solicitante"
  Edit.AddField "chtSetor",        "Setor"
  Edit.AddField "chtDataSolicita", "Data",    efFldTypeText, efValOptDate, ,  Strings.ZTILongDate(Now)
  Edit.AddField "chtHoraSolicita", "Horário", efFldTypeText, , ,  Strings.LeadingZeroes(Hour(Now), 2) & ":" &  Strings.LeadingZeroes(Minute(Now), 2)
  Edit.AddField "chtProblemaTipo", "Tipo de Problema Detectado", efFldTypeRadio, , , conTypeHardware
  Edit.AddField "chtDescricao",    "Descrição", efFldTypeTextArea

  Edit.AddField "chtAtendente",       "Atendido Por", efFldTypeText
  Edit.AddField "chtTempoAtende",     "Tempo de Atendimento", efFldTypeText
  Edit.AddField "chtDataAtende",      "Data",    efFldTypeText, efValOptDate, ,  Strings.ZTILongDate(Now)
  Edit.AddField "chtHoraAtende",      "Horário", efFldTypeText, , ,  Strings.LeadingZeroes(Hour(Now), 2) & ":" &  Strings.LeadingZeroes(Minute(Now), 2)
  Edit.AddField "chtProblemaDetecta", "Problemas Detectados", efFldTypeTextArea
  Edit.AddField "chtSolucao",         "Solução",    efFldTypeTextArea
  Edit.AddField "chtPendencias",      "Pendências", efFldTypeTextArea

  Edit.FieldRadioColumns "chtProblemaTipo", conNumColumns

  Edit.FieldAddValue "chtProblemaTipo", conTypeHardware, "Hardware"
  Edit.FieldAddValue "chtProblemaTipo", conTypeSoftware, "Software"
  Edit.FieldAddValue "chtProblemaTipo", conTypeMicroinf, "Microinformática"

  Edit.FieldShowSize "chtSetor", "40"

  Edit.FieldShowSize "chtDescricao", "80"
  Edit.FieldTextAreaHeight "chtDescricao", 2

  Edit.FieldShowSize "chtProblemaDetecta", "80"
  Edit.FieldTextAreaHeight "chtProblemaDetecta", 4

  Edit.FieldShowSize "chtSolucao", "80"
  Edit.FieldTextAreaHeight "chtSolucao", 4

  Edit.FieldShowSize "chtPendencias", "80"
  Edit.FieldTextAreaHeight "chtPendencias", 2

  Edit.FieldInternalLink "chtSolicitante", Edit.parWhat, "chtCodigo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub GetChamaTecnico
REM =========================================================================

REM =========================================================================
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  Edit.ConnectionString = Session("ConnectionString")

  Edit.QueryString

  Select Case Edit.parWhat
    Case conWhatChamaTecnico
      GetChamaTecnico

'      Para criar a tabela retire o rem abaixo
REM      Edit.TableCreate "sncChamaTecnico"

  End Select

  Edit.RedirectActions

  Default.BodyText  = "navy"
  Default.BodyLink  = "navy"
  Default.BodyVLink = "#0B8D94"
  Default.BodyALink = "#0B8D94"

  Default.HTMLBegin
  Default.HeadAll "Suporte de Informática"
  Default.BodyBegin
  Show.Center

  Table.BeginTable "65%"
  Table.Style.Color2 = ""
  Table.BeginRow 4

  Table.ColumnVAlign = "Middle"
  Table.Column "<img src=""/HiperTools/img/Livro.gif"">"
  Table.Column "<Font Color=""DarkBlue"">Suporte de Informática"

  Table.EndRow
  Table.EndTable

  FormatTable

  Set Edit.Style = Table.Style

  If Edit.IsMyAction Then
     ShowFirstPage
  End If

  Default.PageFooterDefault
  Default.BodyEnd
  Default.HTMLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do Senac.asp
REM =========================================================================
%>