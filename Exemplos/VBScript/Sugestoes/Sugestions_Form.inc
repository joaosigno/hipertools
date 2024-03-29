<%
REM =========================================================================
REM /Sugestions_Form.inc
REM -------------------------------------------------------------------------
REM Nome     : Sugest�es HiperTools
REM Descricao:
REM Home     : www.zevallos.com.br/sugest
REM Criacao  : 2000/04/27 12:18AM
REM Autor    : Ridai Govinda <ridai@zevallos.com.br>
REM Versao   : 1.0.0
REM Local    :  - DF
REM Companhia: Zevallos
REM -------------------------------------------------------------------------

REM =========================================================================
REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela
REM "Bugs"
REM -------------------------------------------------------------------------
Public Sub EditFormBugs

  dataBugs

  Edit.FormBegin  "sugBugs", "Bugs do HiperTools30", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "bugRemetente;bugSumario;bugDescricao;bugDataHora,bugClass,bugReferente"
    Edit.FormFind "bugRemetente,bugSumario,bugDescricao,bugDataHora,bugClass,bugReferente"
    Edit.FormList "bugRemetente,bugSumario,bugDescricao,bugDataHora,bugClass,bugReferente"

      Edit.AddField "bugCodigo",    "&C�digo", , , ,efNext
      Edit.AddField "bugRemetente", "&Remetente", efFldTypeLookup
      Edit.AddField "bugDescricao", "&Descri��o", efFldTypeTextArea
      Edit.AddField "bugSumario",   "&Sum�rio"
      Edit.AddField "bugDataHora",  "&Data", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "bugClass",     "Cla&sse", efFldTypeLookup
      Edit.AddField "bugReferente", "Re&ferente �", efFldTypeLookup

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
REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela
REM "Classes"
REM -------------------------------------------------------------------------
Public Sub EditFormClasses

  dataClasses

  Edit.FormBegin  "sugClasses", "Classes (objetos)", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "clsNome;clsVersao,clsTamanho,clsPosition;clsDescricao"
    Edit.FormFind "clsNome,clsVersao,clsTamanho,clsDescricao"
    Edit.FormList "clsNome,clsVersao,clsTamanho,clsDescricao"

      Edit.AddField "clsCodigo", "&C�digo", , , , efNext
      Edit.AddField "clsNome", "&Nome da Classe"
      Edit.AddField "clsVersao", "&Vers�o"
      Edit.AddField "clsTamanho", "&Tamanho"
      Edit.AddField "clsPosition",  "&Posi��o", efFldTypeLookUp
      Edit.AddField "clsDescricao", "&Descri��o",efFldTypeTextArea

       Edit.FieldInternalLink "clsNome", Edit.parWhat, "clsCodigo"

    Edit.FieldLookup "clsPosition", "sugClassPosition", "posCodigo", "posDescricao"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "Classes
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela
REM "ClassPosition"
REM -------------------------------------------------------------------------
Public Sub EditFormClassPosition

  dataClassPosition

  Edit.FormBegin  "sugClassPosition", "Posi��o da Classe", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "posDescricao;posEstimEntrega,posResponsavel"
    Edit.FormFind "posDescricao,posEstimEntrega,posResponsavel"
    Edit.FormList "posDescricao,posEstimEntrega,posResponsavel"

      Edit.AddField "posCodigo", "&C�digo", , , ,efNext
      Edit.AddField "posDescricao", "&Resumo da Situa��o"
      Edit.AddField "posEstimEntrega", "Data de Entrega Esti&mada", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "posResponsavel", "&Respons�vel", efFldTypeLookUp

      Edit.FieldInternalLink "posDescricao", Edit.parWhat, "posCodigo"

    Edit.FieldLookup "posResponsavel", "sugPessoas", "pssCodigo", "pssNome"

  Edit.FormEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "ClassPosition"
REM =========================================================================

REM =========================================================================
REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela
REM "Ideias"
REM -------------------------------------------------------------------------
Public Sub EditFormIdeias

  dataIdeias

  Edit.FormBegin  "sugIdeias", "Id�ias para o HiperTools30", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "ideRemetente;ideSumario;ideDescricao;ideDataHora,ideClass,ideReferente"
    Edit.FormFind "ideRemetente,ideSumario,ideDescricao,ideDataHora,ideClass,ideReferente"
    Edit.FormList "ideRemetente,ideSumario,ideDescricao,ideDataHora,ideClass,ideReferente"

      Edit.AddField "ideCodigo", "&C�digo", , , , efNext
      Edit.AddField "ideRemetente", "&Remetente", efFldTypeLookUp
      Edit.AddField "ideDescricao", "&Descri��o", efFldTypeTextArea
      Edit.AddField "ideSumario", "&Sum�rio"
      Edit.AddField "ideDataHora", "&Data", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "ideClass", "Cla&sse", efFldTypeLookUp
      Edit.AddField "ideReferente", "Re&ferente �", efFldTypeLookUp

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
REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela
REM "Melhorias"
REM -------------------------------------------------------------------------
Public Sub EditFormMelhorias

  dataMelhorias

  Edit.FormBegin  "sugMelhorias", "Melhorias implementadas no HiperTools30", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "mlhAutor;mlhSumario;mlhDescricao;mlhData;mlhClass,mlhReferente"
    Edit.FormFind "mlhAutor,mlhSumario,mlhDescricao,mlhData,mlhClass,mlhReferente"
    Edit.FormList "mlhAutor,mlhSumario,mlhDescricao,mlhData,mlhClass,mlhReferente"

      Edit.AddField "mlhCodigo", "&C�digo", , , , efNext
      Edit.AddField "mlhAutor", "&Autor/Respos�vel", efFldTypeLookup
      Edit.AddField "mlhDescricao", "&Descri��o",efFldTYpeTextArea
      Edit.AddField "mlhSumario", "&Sum�rio"
      Edit.AddField "mlhData", "&Data", , efValOptDate, , Strings.LongDate(Now)
      Edit.AddField "mlhClass", "Cla&sse", efFldTypeLookup
      Edit.AddField "mlhReferente", "&Referente �", efFldTypeLookup

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
REM Inicio do procedimento que desenha o formul�rio de edi��o da tabela
REM "Pessoas"
REM -------------------------------------------------------------------------
Public Sub EditFormPessoas

  dataPessoas

  Edit.FormBegin  "sugPessoas", "PeopleWare (Pessoas envolvidas)", 1, Edit.parWhat, efValLocClient

    Edit.FormUnit "pssNome;pssEmail;pssFone,pssIsDeveloper"
    Edit.FormFind "pssNome,pssEmail,pssFone,pssIsDeveloper"
    Edit.FormList "pssNome,pssEmail,pssFone,pssIsDeveloper"

      Edit.AddField "pssCodigo",      "&C�digo", , , , efNext
      Edit.AddField "pssNome",        "&Nome"
      Edit.AddField "pssEmail",       "&Email", efFldTypeEmail, efValOptEmail
      Edit.AddField "pssFone",        "&Fone"
      Edit.AddField "pssIsDeveloper", "� &desenvolvedor?", efFldTypeCheck, , , , "N"

        Edit.FieldAddValue "pssIsDeveloper","S", "Sim"
        Edit.FieldAddValue "pssIsDeveloper","N", "N�o"

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