<%
REM =========================================================================
REM  /Import.inc
REM -------------------------------------------------------------------------
REM Nome     : Importa a estrutura de dados para o ZEUS III
REM Descricao: Programa para importacao dos dados para os Banco de Dados do
REM          : ZEUSIII
REM Home     : http://www.hipertools.com.br/
REM Criacao  : 2/12/0 5:14PM
REM Autor    : Zevallos Tecnologia em Informacao
REM Versao   : 1.1.0.0
REM Local    : Brasília - DF
REM Companhia: 2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

  Dim sobjRS, sobjConn

  Set sobjRS = Server.CreateObject("ADODB.RecordSet")
  Set sobjConn = Server.CreateObject("ADODB.Connection")

REM =========================================================================
REM Procedimento que mostra a "Primeira Pagina"
REM -------------------------------------------------------------------------
Private Sub ShowImportLogon

  Session( "CompanyCode" )      = Request.QueryString( conImportCompany )
  Session( "SystemCode" )       = Request.QueryString( conImportSystem )

  Show.Center
  Table.Padding = 2
  Table.Border =  0
  Table.BeginTable "40%", "Base de dados a importar", , True

    Form.BeginForm Initializer.ScriptName & URL.Equal( conOptions, conOptionTable), "frmDataBase"

      Table.BeginRow
        Table.BeginCell
          Show.HTMLCR Strings.BoldText( "Server" )
          Show.BR
          If Request.Cookies( "ServerName" ) > "" Then
             Form.Text "Server", 30, Request.Cookies( "ServerName" )
          Else
             Form.Text "Server", 30, "ServerName"
          End If

      Table.BeginRow
        Table.BeginCell
          Show.HTMLCR Strings.BoldText( "User" )
          Show.BR
          If Request.Cookies( "UserName" ) > "" Then
             Form.Text "Userid", 30, Request.Cookies( "UserName" )
          Else
             Form.Text "Userid", 30, "UserName"
          End If

      Table.BeginRow
        Table.BeginCell
          Show.HTMLCR Strings.BoldText( "Password" )
          Show.BR
          Form.Password "Password", 30

      Table.BeginRow
        Table.BeginCell
          Show.HTMLCR Strings.BoldText( "Database" )
          Show.BR
          If Request.Cookies( "DatabaseName" ) > "" Then
             Form.Text "Database", 30, Request.Cookies( "DatabaseName" )
          Else
             Form.Text "Database", 30, "Database"
          End If

      Table.BeginRow
        Table.BeginCell
          Show.HTMLCR Strings.BoldText( "Tipo de banco de dados" )
          Show.BR
          Form.Radio "DatabaseType", "SQLServer", , ( ( Not Request.Cookies( "DatabaseType" ) > "" ) Or ( Request.Cookies( "DatabaseType" ) = "SQLServer" ) )
          Show.HTML "SQL-Server"

      Table.BeginRow
        Table.BeginCell
          Show.Center
          Show.HTMLCR "<HR>"
          Form.Submit "bmtConnect", "Próximo >>"
        Table.EndCell
      Table.EndRow

    Form.EndForm

  Table.EndTable

  Show.CenterEnd

End Sub
REM -------------------------------------------------------------------------
REM Fim do "ShowFirstPage"
REM =========================================================================

REM =========================================================================
REM Mostra as Tabelas
REM -------------------------------------------------------------------------
Private Sub ShowTables

  Dim i

  sobjConn.ConnectionTimeout = conConnectionTimeout
  sobjConn.CommandTimeout    = conCommandTimeout

  Select Case Request.Form( "DatabaseType" )
    Case "SQLServer"
         Session(conImportConnection) = Strings.FormatText( "DRIVER={SQL Server};SERVER=$s;DATABASE=$s;UID=$s;PWD=$s", _
                                                               Request.Form( "Server"   ), _
                                                               Request.Form( "Database" ), _
                                                               Request.Form( "Userid"   ), _
                                                               Request.Form( "Password" ) )
    Case Else
         Exit Sub
  End Select

  sobjConn.Open Session( conImportConnection )

  Form.BeginForm Initializer.ScriptName & URL.Equal(conOptions, conOptionRename), "frmTables"

    Table.BeginTable "40%", "Base de Dados", , True

      Table.HeaderRow "Tabelas"

      Table.BeginRow 2
        Table.BeginCell

          Show.Center

            sobjRS.Open "Select Name From SysObjects Where SysStat & 0xf = 3 Order By 1", sobjConn

            Form.BeginSelect "Tabelas", 10, True
              Do While Not sobjRS.EOF
                 Form.SelectOption sobjRS("Name"), sobjRS("Name")
                 sobjRS.MoveNext
              Loop
              sobjRS.Close
            Form.EndSelect

            Show.BR

            Form.CheckBox "Sobrepor", "Sobrepor", , 1
            Show.HTML "Sobrepor"

        Table.EndCell
      Table.EndRow

      Table.BeginRow , True
        Table.BeginCell
           Form.Submit "bmtTable", "Próximo >>"
        Table.EndCell
      Table.EndRow

    Table.EndTable
  Form.EndForm

End Sub
REM -------------------------------------------------------------------------
REM Fim do "ShowTables"
REM =========================================================================

REM =========================================================================
REM Insere a estrutura da tabela dentro do Zeus III
REM -------------------------------------------------------------------------
Private Sub ShowRenameTables

  Dim i, strTableName

  Form.BeginForm Initializer.ScriptName & URL.Equal(conOptions, conOptionImport), "frmTables"

    Table.BeginTable "40%", "Definição das tabelas", , True

      Table.CellColor = Table.Style.HeaderColor
      Table.BeginRow
        Table.BeginCell
          Show.HTML "Sigla"
          Show.nbsp 2
          Show.HTML "Nome"
        Table.EndCell
      Table.EndRow
      Table.CellColor = ""

      Table.BeginRow
        Table.BeginCell
          For i = 1 To Request.Form( "Tabelas" ).Count
              strTableName = Request.Form( "Tabelas" )( i )
              Form.Text "Sigla_" & strTableName, 3, LCase( Mid( strTableName, 4, 3 ) ), 3
              Show.nbsp
              Show.HTML Mid( strTableName, 4, Len( strTableName ) )
              Show.BR
          Next
        Table.EndCell
      Table.EndRow

      Table.BeginRow , True
        Table.BeginCell
           Form.Submit "bmtTables", "Próximo >>"
        Table.EndCell
      Table.EndRow

    Table.EndTable
  Form.EndForm

End Sub
REM -------------------------------------------------------------------------
REM Fim do "ImportTables"
REM =========================================================================

REM =========================================================================
REM Insere a estrutura da tabela dentro do Zeus III
REM -------------------------------------------------------------------------
Private Sub ImportTables

  Dim strItem, j, intTableCode
  Dim objConn2, objRS2
  Dim strFieldName, strFieldType, strFieldSize, strFieldNull, strFieldKey

  Set objConn2 = Server.CreateObject("ADODB.Connection")
  Set objRS2   = Server.CreateObject("ADODB.RecordSet")

  sobjConn.ConnectionTimeout = conConnectionTimeout
  sobjConn.CommandTimeout    = conCommandTimeout
  sobjConn.Open Session( "ConnectionString" )

  objConn2.ConnectionTimeout = conConnectionTimeout
  objConn2.CommandTimeout    = conCommandTimeout
  objConn2.Open Session( conImportConnection )

  For Each strItem In Request.Form
      If Left( strItem, 6 ) = "Sigla_" Then
         If InsertTableName( Mid( strItem, 7 ), Session( "SystemCode" ), intTableCode, Request.Form( strItem ) ) Then

            sobjRS.Open Mid( strItem, 7 ), objConn2, adOpenForwardOnly, adLockReadOnly, adCmdTable

            For j = 0 To sobjRS.Fields.Count - 1

                objRS2.Open "SELECT ( Max( camCodigo ) + 1 ) AS NextCode FROM zeuCampos" , sobjConn, adOpenDynamic, adLockReadOnly
                strFieldName = Mid( sobjRS( j ).Name, 4 )
                If sobjRS( j ).Type = adChar Then
                   strFieldType =  0
                   strFieldSize = Cstr( sobjRS( j ).DefinedSize )
                ElseIf sobjRS( j ).Type = adVarChar Then
                   strFieldType =  1
                   strFieldSize = Cstr( sobjRS( j ).DefinedSize )
                ElseIf sobjRS( j ).Type = adDBTimeStamp Then
                   strFieldType =  2
                   strFieldSize = "10"
                ElseIf sobjRS( j ).Type = adLongVarChar Then
                   strFieldType =  3
                   strFieldSize = "6000"
                ElseIf sobjRS( j ).Type = adInteger Then
                   strFieldType =  4
                   strFieldSize = "4"
                ElseIf sobjRS( j ).Type = adSmallInt Then
                   strFieldType =  5
                   strFieldSize = "2"
                ElseIf sobjRS( j ).Type = adUnsignedTinyInt Then
                   strFieldType =  6
                   strFieldSize = "1"
                ElseIf sobjRS( j ).Type = adDouble Then
                   strFieldType =  7
                   strFieldSize = "10"
                ElseIf sobjRS( j ).Type = adNumeric Then
                   strFieldType =  8
                   strFieldSize = "10"
                ElseIf sobjRS( j ).Type = adCurrency Then
                   strFieldType =  9
                   strFieldSize = "10"
                ElseIf sobjRS( j ).Type = adBoolean Then
                   strFieldType =  10
                   strFieldSize = "1"
                Else
                   strFieldType = 3
                   strFieldSize = "200"
                   Show.Message strFieldName & "=" & Cstr( sobjRS( j ).Type ) &  "=" & Cstr( sobjRS( j ).DefinedSize )
                End If

                If CBool( sobjRS( j ).Attributes And adFldIsNullable ) Then
                   strFieldNull = 0
                Else
                   strFieldNull = 1
                End If

                strFieldKey = 0

                sobjConn.Execute "INSERT INTO zeuCampos ( camTabela, camCodigo, camNome, camTipo, camTamanho, camOrdem, camRotulo, camPrimario, " & _
                                 "camRequerido, camDelimitador, camTipoEdicao, camTipoValidacao, camListagem, camMostrar, camLocalizacao ) VALUES ( " & _
                                 intTableCode  & ", " & objRS2("NextCode")  & ", " & Strings.SQLString( strFieldName  ) & ", " & strFieldType  & ", " & _
                                 strFieldSize  & ", " & j & ", " & Strings.SQLString( strFieldName ) & ", " & strFieldKey & ", " & strFieldNull  & ", ';', 1, 1, 1, 1, 1 )" , , adCmdText

                objRS2.Close

            Next
         End If
      End If
  Next

  Set objConn2 = Nothing
  Set objRS2   = Nothing

End Sub
REM -------------------------------------------------------------------------
REM Fim do "ImportTables"
REM =========================================================================

REM =========================================================================
REM Insere a estrutura da tabela dentro do Zeus III
REM -------------------------------------------------------------------------
Private Function InsertTableName( ByVal strTabeName, ByVal strSystemCode, intTableCode, ByVal strSigla )

  Dim Code

  If sobjRS.State = adStateOpen Then
     sobjRS.Close
  End If

  sobjRS.Open "SELECT tabCodigo FROM zeuTabelas WHERE tabNome = " & Strings.SQLString( Mid( strTabeName, 4 ) ) & _
              " AND tabSistema = " & strSystemCode, sobjConn, adOpenForwardOnly, adLockReadOnly

  InsertTableName = False

  If Not ( sobjRS.EOF And sobjRS.BOF ) Then

     If Request.Form( "Sobrepor" ) <> "Sobrepor" Then

        Show.MessageError "Não foi possível importar a tabela " & Request.Form( "Tabelas" )( i ) & " pois esta já existe neste sistema."

        Exit Function

    End If

    Code = sobjRS( "tabCodigo" )

    sobjRS.Close

  Else

     sobjRS.Close

     sobjRS.Open "SELECT ( Max( tabCodigo ) + 1 ) AS NewCode FROM zeuTabelas", sobjConn, adOpenForwardOnly, adLockReadOnly

     Code = sobjRS( "NewCode" )

     sobjConn.Execute "INSERT INTO zeuTabelas ( tabCodigo, tabSistema, tabSigla, tabNome ) VALUES (" & Code & ", " & strSystemCode & ", " & _
                         Strings.SQLString( strSigla ) & ", " & Strings.SQLString( Mid( strTabeName, 4 ) ) & " )", , adCmdText

     sobjRS.Close


  End If

  InsertTableName =  True

  intTableCode = Code

  sobjConn.Execute "DELETE FROM zeuCampos WHERE camTabela = " & Code, , adCmdText

End Function
REM -------------------------------------------------------------------------
REM Fim do "InsertTables"
REM =========================================================================

REM -------------------------------------------------------------------------
REM Fim do Import.inc
REM =========================================================================
%>