<%
REM -------------------------------------------------------------------------
REM  /HiperTools/Table.inc
REM -------------------------------------------------------------------------
REM  Descricao  : Constantes do HiperTools30.Table / HiperTools30.TableStyle
REM  Criação    : 19/02/2000 11:21AM
REM  Local      : Brasilia/DF
REM  Elaborado  : Ruben Zevallos <ruben@zevallos.com.br>
REM               Fernando Aquino <fernando@zevallos.com.br>
REM               Kelson Medeiros <kelson@zevallos.com.br>
REM               Ueslei Silva <ueslei@zevallos.com.br>
REM               Eduardo Silva <ueslei@zevallos.com.br>
REM               Ridai Govinda <ridai@zevallos.com.br>
REM  Copyright  : 2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM Style.BorderFormat:
REM ------------------------------------------------------
Const tbBdFormatCompleteGrid = 0    'Grade completa
Const tbBdFormatOnlyLines = 1       'Apenas linhas
Const tbBdFormatOutterGridLines = 2 'Grade externa + linhas
Const tbBdFormatInvisible = 3       'Sem grade

REM HiperTools30.Table.Style.ColorFormat:
REM ------------------------------------------------------
Const tbStFormatHeader = 0                 'Cabecalho cor 1, resto cor 2
Const tbStFormatHeaderAlternateColumns = 1 'Alterna colunas
Const tbStFormatAlternateColumns = 2       'Cabecalho fixo, alterna colunas
Const tbStFormatHearderAlternateLines = 3  'Alterna linhas
Const tbStFormatAlternateLines = 4         'Cabecalho fixo, alterna linhas
Const tbStFormatNothing = 5                'Todo da cor 1(menos o Título)
Const tbStFormatTitle = 6                  'Todo da cor 1(inclusive Título)
%>