<%
// -------------------------------------------------------------------------
// /HiperTools/EditForm.inc
// -------------------------------------------------------------------------
// Descricao : Constantes do HiperTools30.EditForm
// Criação  : 19/02/2000 11:21AM
// Local  : Brasilia/DF
// Elaborado : Ruben Zevallos <ruben@zevallos.com.br>
//    Fernando Aquino <fernando@zevallos.com.br>
//    Kelson Medeiros <kelson@zevallos.com.br>
//    Ueslei Silva <ueslei@zevallos.com.br>
//    Eduardo Silva <ueslei@zevallos.com.br>
//    Flavio Costa <flavio@zevallos.com.br>
//    Ridai Govinda <ridai@zevallos.com.br>
// Copyright : 2000 by Zevallos(r) Tecnologia em Informacao
// -------------------------------------------------------------------------

// Entradas da Registry
// -----------------------------
var rgHKEYClassesRoot = 0x80000000;
var rgHKEYCurrentUser = 0x80000001;
var rgHKEYLocalMachine = 0x80000002;
var rgHKEYUsers = 0x80000003;
var rgHKEYPerformanceData = 0x80000004;
var rgHKEYCurrentConfig = 0x80000005;
var rgHKEYDynData = 0x80000006;

// Tipos de Valores do Registry
// -----------------------------
var rgNone = 0;                      // Sem tipo de dado
var rgSz = 1;                        // String tipo Unicode terminada com nulo
var rgExpandSz = 2;                  // String tipo Unicode terminada com nulo com
                                     // referências à variáveis de ambiente: %windir% p/ exemplo
var rgBinary = 3;                    // Dado binário (livre)
var rgDword = 4;                     // Número 32-bit (Long Integer)
var rgDwordLittleEndian = 4;         // Número 32-bit (o mesmo que rgDword)
var rgDwordBigEndian = 5;            // Número 32-bit
var rgLink = 6;                      // Link simbólico (unicode)
var rgMultiSz = 7;                   // Strings tipo Unicode múltiplas
var rgResourceList = 8;              // Resource list in the resource map
var rgFullResourceDescriptor = 9;    // Resource list in the hardware description
var rgResourceRequirementsList = 10; // Resource requirements list
%>