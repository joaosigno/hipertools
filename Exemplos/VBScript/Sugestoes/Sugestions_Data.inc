<%

REM =========================================================================
REM /Sugestions_Data.inc
REM -------------------------------------------------------------------------
REM Nome     : Sugestões HiperTools
REM Descricao:
REM Home     : www.zevallos.com.br/sugest
REM Criacao  : 2000/04/27 12:18AM
REM Autor    : Ridai Govinda <ridai@zevallos.com.br>
REM Versao   : 1.0.0
REM Local    : Brasilia - DF
REM Companhia: Zevallos Tecnologia em Inforamcao
REM -------------------------------------------------------------------------


Public sblnCreateTable

sblnCreateTable = False

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Bugs"
REM -------------------------------------------------------------------------
Public Sub dataBugs

  Edit.DataBegin "sugBugs"

  Edit.DataAddField "bugCodigo", efDataTypeInt, 4, efRequired
  Edit.DataAddField "bugRemetente", efDataTypeInt, 4, efRequired
  Edit.DataAddField "bugDescricao", efDataTypeText, 6000, efNotRequired
  Edit.DataAddField "bugSumario", efDataTypeVarChar, 100, efNotRequired
  Edit.DataAddField "bugDataHora", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "bugClass", efDataTypeInt, 4, efRequired
  Edit.DataAddField "bugReferente", efDataTypeInt, 4, efNotRequired

  Edit.DataAddPrimaryKey "bugCodigo"

  Edit.DataAddIndex "FK_bugRemetente",  "bugRemetente"
  Edit.DataAddIndex "FK_bugClass", "bugClass"
  Edit.DataAddIndex "FK_bugReferente", "bugReferente"

  If ( Not Edit.IsTable( "sugBugs" ) )  Or sblnCreateTable Then
      Edit.TableCreate "sugBugs"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataBugs
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Classes"
REM -------------------------------------------------------------------------
Public Sub dataClasses

  Edit.DataBegin "sugClasses"

  Edit.DataAddField "clsCodigo",    efDataTypeInt, 4, efRequired
  Edit.DataAddField "clsNome",      efDataTypeVarChar, 50, efNotRequired
  Edit.DataAddField "clsVersao",    efDataTypeVarChar, 20, efNotRequired
  Edit.DataAddField "clsTamanho",   efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "clsPosition",  efDataTypeInt, 4, efRequired
  Edit.DataAddField "clsDescricao", efDataTypeText, 6000, efRequired

  Edit.DataAddPrimaryKey "clsCodigo"

  Edit.DataAddIndex "FK_clsPosition",  "clsPosition"
  Edit.DataAddIndex "IE_clsNome", "clsNome"

  If ( Not Edit.IsTable( "sugClasses" ) )  Or sblnCreateTable Then
      Edit.TableCreate "sugClasses"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataClasses
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "ClassPosition"
REM -------------------------------------------------------------------------
Public Sub dataClassPosition

  Edit.DataBegin "sugClassPosition"

  Edit.DataAddField "posCodigo",        efDataTypeInt, 4, efRequired
  Edit.DataAddField "posDescricao",     efDataTypeVarChar, 80, efNotRequired
  Edit.DataAddField "posEstimEntrega",  efDataTypeVarChar, 20, efNotRequired
  Edit.DataAddField "posResponsavel",   efDataTypeInt, 4, efNotRequired

  Edit.DataAddPrimaryKey "posCodigo"

  Edit.DataAddIndex "FK_posResp",  "posResponsavel"
  Edit.DataAddIndex "IE_posEstEnt", "posEstimEntrega"

  If ( Not Edit.IsTable( "sugClassPosition" ) )  Or sblnCreateTable Then
      Edit.TableCreate "sugClassPosition"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataClassPosition
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Ideias"
REM -------------------------------------------------------------------------
Public Sub dataIdeias

  Edit.DataBegin "sugIdeias"

  Edit.DataAddField "ideCodigo", efDataTypeInt, 4, efRequired
  Edit.DataAddField "ideRemetente", efDataTypeInt, 4, efRequired
  Edit.DataAddField "ideDescricao", efDataTypeText, 6000, efNotRequired
  Edit.DataAddField "ideSumario", efDataTypeVarChar, 100, efNotRequired
  Edit.DataAddField "ideDataHora", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "ideClass", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "ideReferente", efDataTypeInt, 4, efNotRequired

  Edit.DataAddPrimaryKey "ideCodigo"

  Edit.DataAddIndex "FK_ideRemetente",  "ideRemetente"
  Edit.DataAddIndex "FK_ideClass", "ideClass"
  Edit.DataAddIndex "FK_ideReferente", "ideReferente"

  If ( Not Edit.IsTable( "sugIdeias" ) )  Or sblnCreateTable Then
      Edit.TableCreate "sugIdeias"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataIdeias
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Ideias"
REM -------------------------------------------------------------------------
Public Sub dataMelhorias

  Edit.DataBegin "sugMelhorias"

  Edit.DataAddField "mlhCodigo", efDataTypeInt, 4, efRequired
  Edit.DataAddField "mlhAutor", efDataTypeInt, 4, efRequired
  Edit.DataAddField "mlhDescricao", efDataTypeText, 6000, efNotRequired
  Edit.DataAddField "mlhSumario", efDataTypeVarChar, 100, efNotRequired
  Edit.DataAddField "mlhData", efDataTypeDateTime, 10, efNotRequired
  Edit.DataAddField "mlhClass", efDataTypeInt, 4, efNotRequired
  Edit.DataAddField "mlhReferente", efDataTypeInt, 4, efNotRequired

  Edit.DataAddPrimaryKey "mlhCodigo"

  Edit.DataAddIndex "FK_mlhAutor",  "mlhAutor"
  Edit.DataAddIndex "FK_mlhClass", "mlhClass"
  Edit.DataAddIndex "FK_mlhReferente", "mlhReferente"

  If ( Not Edit.IsTable( "sugMelhorias" ) )  Or sblnCreateTable Then
      Edit.TableCreate "sugMelhorias"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataIdeias
REM =========================================================================

REM =========================================================================
REM Procedimento que define a estrutura de dados tabela "Pessoas"
REM -------------------------------------------------------------------------
Public Sub dataPessoas

  Edit.DataBegin "sugPessoas"

  Edit.DataAddField "pssCodigo",      efDataTypeInt,       4, efRequired
  Edit.DataAddField "pssNome",        efDataTypeVarChar, 100, efRequired
  Edit.DataAddField "pssEmail",       efDataTypeVarChar, 100, efNotRequired
  Edit.DataAddField "pssFone",        efDataTypeVarChar,  15, efNotRequired
  Edit.DataAddField "pssIsDeveloper", efDataTypeVarChar,   1, efNotRequired

  Edit.DataAddPrimaryKey "pssCodigo"

  Edit.DataAddIndex "IE_pssNome",  "pssNome"

  If ( Not Edit.IsTable( "sugPessoas" ) )  Or sblnCreateTable Then
      Edit.TableCreate "sugPessoas"
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim do dataPessoas
REM =========================================================================

REM =========================================================================
REM Procedimento que define cria todas as tabelas do sistema
REM -------------------------------------------------------------------------
Public Sub CriaTodasTabelas()
  sblnCreateTable = True

  dataPessoas
  dataBugs
  dataClasses
  dataIdeias
  dataClassPosition
  dataMelhorias

  sblnCreateTable = False
End Sub
REM -------------------------------------------------------------------------
REM Fim do CriaTodasTabelas
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do /Sugestions_Form.inc
REM =========================================================================
%>