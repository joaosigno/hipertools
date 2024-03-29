<%
REM -------------------------------------------------------------------------
REM  /HiperTools/Registry.inc
REM -------------------------------------------------------------------------
REM  Descricao  : Constantes do HiperTools30.Registry
REM  Cria��o    : 19/02/2000 11:21AM
REM  Local      : Brasilia/DF
REM  Elaborado  : Ruben Zevallos <ruben@zevallos.com.br>
REM               Fernando Aquino <fernando@zevallos.com.br>
REM               Kelson Medeiros <kelson@zevallos.com.br>
REM               Ueslei Silva <ueslei@zevallos.com.br>
REM               Eduardo Silva <ueslei@zevallos.com.br>
REM               Flavio Costa <flavio@zevallos.com.br>
REM               Ridai Govinda <ridai@zevallos.com.br>
REM  Copyright  : 2000 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM Entradas da Registry
REM -----------------------------
Const rgHKEYClassesRoot = &H80000000
Const rgHKEYCurrentUser = &H80000001
Const rgHKEYLocalMachine = &H80000002
Const rgHKEYUsers = &H80000003
Const rgHKEYPerformanceData = &H80000004
Const rgHKEYCurrentConfig = &H80000005
Const rgHKEYDynData = &H80000006

REM Tipos de Valores do Registry
REM -----------------------------
Const rgNone = 0                        ' Sem tipo de dado
Const rgSz = 1                          ' String tipo Unicode terminada com nulo
Const rgExpandSz = 2                    ' String tipo Unicode terminada com nulo com
                                        ' refer�ncias � vari�veis de ambiente: %windir% p/ exemplo
Const rgBinary = 3                      ' Dado bin�rio (livre)
Const rgDword = 4                       ' N�mero 32-bit (Long Integer)
Const rgDwordLittleEndian = 4           ' N�mero 32-bit  (o mesmo que rgDword)
Const rgDwordBigEndian = 5              ' N�mero 32-bit
Const rgLink = 6                        ' Link simb�lico (unicode)
Const rgMultiSz = 7                     ' Strings tipo Unicode m�ltiplas
Const rgResourceList = 8                ' Resource list in the resource map
Const rgFullResourceDescriptor = 9      ' Resource list in the hardware description
Const rgResourceRequirementsList = 10   ' Resource requirements list
%>