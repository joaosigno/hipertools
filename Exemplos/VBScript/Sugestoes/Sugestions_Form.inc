<%
REM =========================================================================
REM /Sugestions_Form.inc
REM -------------------------------------------------------------------------
REM Nome     : Sugestхes HiperTools
REM Descricao:
REM Home     : www.zevallos.com.br/sugest
REM Criacao  : 2000/04/27 12:18AM
REM Autor    : Ridai Govinda <ridai@zevallos.com.br>
REM Versao   : 1.0.0
REM Local    :  - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Inicio do procedimento que desenha o formulбrio de ediзгo da tabela
REM "Bugs"
REM -------------------------------------------------------------------------
Public Sub EditFormBugs

  dataBugs

  Edit.FormBegin  "sugBugs", "Bugs do HiperTools30", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "bugRemetente;bugSumario;bugDescricao;bugDataHora,bugClass,bugReferente"
    Edit.FormFind "bugRemetente,bugSumario,bugDescricao,bugDataHora,bugClass,bugReferente"
    Edit.FormList "bugRemetente,bugSumario,bugDescricao,bugDataHora,bugClass,bugReferente"

      Edit.AddField "bugCodigo",    "&Cуdigo", , , ,efNext
      Edit.AddField "bugRemetente", "&Remetente", efFldTypeLookup
      Edit.AddField "bugDescricao", "&Descriзгo", efFldTypeTextArea
      Edit.AddField "bugSumario",   "&Sumбrio"
      Edit.AddField "bugDataHora",  "&Data", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "bugClass",     "Cla&sse", efFldTypeLookup
      Edit.AddField "bugReferente", "Re&ferente а", efFldTypeLookup

       Edit.FieldInternalLink "bugRemetente", Edit.parWhat, "bugCodigo"

  Edit.FieldLookup "bugRemetente", "sugPessoas", "pssCodigo", "pssNome"
  Edit.FieldLookup "bugClass", "sugClasses", "clsCodigo", "clsNome"
  Edit.FieldLookup "bugReferente", "sugBugs", "bugCodigo", "bugSumario"
  Edit.FieldDisableValue "bugDataHora", "bugDataHora", Now(), False

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Bugs
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulбrio de ediзгo da tabela
REM "Classes"
REM -------------------------------------------------------------------------
Public Sub EditFormClasses

  dataClasses

  Edit.FormBegin  "sugClasses", "Classes (objetos)", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "clsNome;clsVersao,clsTamanho,clsPosition;clsDescricao"
    Edit.FormFind "clsNome,clsVersao,clsTamanho,clsDescricao"
    Edit.FormList "clsNome,clsVersao,clsTamanho,clsDescricao"

      Edit.AddField "clsCodigo", "&Cуdigo", , , , efNext
      Edit.AddField "clsNome", "&Nome da Classe"
      Edit.AddField "clsVersao", "&Versгo"
      Edit.AddField "clsTamanho", "&Tamanho"
      Edit.AddField "clsPosition",  "&Posiзгo", efFldTypeLookUp
      Edit.AddField "clsDescricao", "&Descriзгo",efFldTypeTextArea

       Edit.FieldInternalLink "clsNome", Edit.parWhat, "clsCodigo"

    Edit.FieldLookup "clsPosition", "sugClassPosition", "posCodigo", "posDescricao"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Classes
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulбrio de ediзгo da tabela
REM "ClassPosition"
REM -------------------------------------------------------------------------
Public Sub EditFormClassPosition

  dataClassPosition

  Edit.FormBegin  "sugClassPosition", "Posiзгo da Classe", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "posDescricao;posEstimEntrega,posResponsavel"
    Edit.FormFind "posDescricao,posEstimEntrega,posResponsavel"
    Edit.FormList "posDescricao,posEstimEntrega,posResponsavel"

      Edit.AddField "posCodigo", "&Cуdigo", , , ,efNext
      Edit.AddField "posDescricao", "&Resumo da Situaзгo"
      Edit.AddField "posEstimEntrega", "Data de Entrega Esti&mada", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "posResponsavel", "&Responsбvel", efFldTypeLookUp

      Edit.FieldInternalLink "posDescricao", Edit.parWhat, "posCodigo"

    Edit.FieldLookup "posResponsavel", "sugPessoas", "pssCodigo", "pssNome"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "ClassPosition"
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulбrio de ediзгo da tabela
REM "Ideias"
REM -------------------------------------------------------------------------
Public Sub EditFormIdeias

  dataIdeias

  Edit.FormBegin  "sugIdeias", "Idйias para o HiperTools30", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "ideRemetente;ideSumario;ideDescricao;ideDataHora,ideClass,ideReferente"
    Edit.FormFind "ideRemetente,ideSumario,ideDescricao,ideDataHora,ideClass,ideReferente"
    Edit.FormList "ideRemetente,ideSumario,ideDescricao,ideDataHora,ideClass,ideReferente"

      Edit.AddField "ideCodigo", "&Cуdigo", , , , efNext
      Edit.AddField "ideRemetente", "&Remetente", efFldTypeLookUp
      Edit.AddField "ideDescricao", "&Descriзгo", efFldTypeTextArea
      Edit.AddField "ideSumario", "&Sumбrio"
      Edit.AddField "ideDataHora", "&Data", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "ideClass", "Cla&sse", efFldTypeLookUp
      Edit.AddField "ideReferente", "Re&ferente а", efFldTypeLookUp

      Edit.FieldInternalLink "ideRemetente", Edit.parWhat, "ideCodigo"

    Edit.FieldLookup "ideRemetente", "sugPessoas", "pssCodigo", "pssNome"
    Edit.FieldLookup "ideClass", "sugClasses", "clsCodigo", "clsNome"
    Edit.FieldLookup "ideReferente", "sugIdeias", "ideCodigo", "ideSumario"

    REM Edit.DoFPLPLink = False
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Ideias
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulбrio de ediзгo da tabela
REM "Melhorias"
REM -------------------------------------------------------------------------
Public Sub EditFormMelhorias

  dataMelhorias

  Edit.FormBegin  "sugMelhorias", "Melhorias implementadas no HiperTools30", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "mlhAutor;mlhSumario;mlhDescricao;mlhData;mlhClass,mlhReferente"
    Edit.FormFind "mlhAutor,mlhSumario,mlhDescricao,mlhData,mlhClass,mlhReferente"
    Edit.FormList "mlhAutor,mlhSumario,mlhDescricao,mlhData,mlhClass,mlhReferente"

      Edit.AddField "mlhCodigo", "&Cуdigo", , , , efNext
      Edit.AddField "mlhAutor", "&Autor/Resposбvel", efFldTypeLookup
      Edit.AddField "mlhDescricao", "&Descriзгo",efFldTYpeTextArea
      Edit.AddField "mlhSumario", "&Sumбrio"
      Edit.AddField "mlhData", "&Data", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "mlhClass", "Cla&sse", efFldTypeLookup
      Edit.AddField "mlhReferente", "&Referente а", efFldTypeLookup

      Edit.FieldInternalLink "mlhAutor", Edit.parWhat, "mlhCodigo"

    Edit.FieldLookup "mlhAutor", "sugPessoas", "pssCodigo", "pssNome"
    Edit.FieldLookup "mlhClass", "sugClasses", "clsCodigo", "clsNome"
    Edit.FieldLookup "mlhReferente", "sugMelhorias", "mlhCodigo", "mlhSumario"

    REM Edit.DoFPLPLink = False
  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Melhorias"
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formulбrio de ediзгo da tabela
REM "Pessoas"
REM -------------------------------------------------------------------------
Public Sub EditFormPessoas

  dataPessoas

  Edit.FormBegin  "sugPessoas", "PeopleWare (Pessoas envolvidas)", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "pssNome;pssEmail;pssFone,pssIsDeveloper"
    Edit.FormFind "pssNome,pssEmail,pssFone,pssIsDeveloper"
    Edit.FormList "pssNome,pssEmail,pssFone,pssIsDeveloper"

      Edit.AddField "pssCodigo",      "&Cуdigo", , , , efNext
      Edit.AddField "pssNome",        "&Nome"
      Edit.AddField "pssEmail",       "&Email", efFldTypeEmail, efValOptEmail
      Edit.AddField "pssFone",        "&Fone"
      Edit.AddField "pssIsDeveloper", "Й &desenvolvedor?", efFldTypeCheck, , , , "N"

        Edit.FieldAddValue "pssIsDeveloper","S", "Sim"
        Edit.FieldAddValue "pssIsDeveloper","N", "Nгo"

      Edit.FieldInternalLink "pssNome", Edit.parWhat, "pssCodigo"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Pessoas
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do /Sugestions_Form.inc
REM =========================================================================
%>