<%
REM -------------------------------------------------------------------------
REM  /HiperTools/Validate.inc
REM -------------------------------------------------------------------------
REM  Descricao  : Constantes do HiperTools30.Validate
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

REM Precedência de validação de CPF ou CGC
REM ---------------------------------------
Const vlCpCPFFirst = 0
Const vlCpCGCFirst = 1

REM Valores de Retorno do Validate.CompareDate
REM ---------------------------------------
Const vlCmpDateFirstInvalid = 1
Const vlCmpDateSecondInvalid = 2
Const vlCmpDateBothInvalid = 3
Const vlCmpDateOutOfRange = 4

%>