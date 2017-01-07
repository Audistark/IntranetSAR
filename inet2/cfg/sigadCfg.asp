<%
'Arquivo de configuraзгo de variбveis

'URL do WebService do Prуton
'const urlWSProton = "http://www.webservicex.net/usaddressverification.asmx?WSDL" ' Teste para ver os erros
'const urlWSProton = "http://homologacao.anac.gov.br/proton/proton.asmx?op=" 'URL de Homologaзгo
'const urlWSProton = "http://sdadf1004.anac.gov.br/wsproton/Proton.asmx?op=" 'URL de Desenvolvimento
'Const urlWSProton = "http://sigad.anac.gov.br/proton/Proton.asmx?op=" 'URL de Produзгo
Const urlWSProton = "none" 'nгo tem SIGAD

'--- usuбrios para cadastro
Dim cod_usuarioProton : cod_usuarioProton = "3168" 'SINTAC_HABILITACAO Desenv e Homolog й o mesmo ID
const CONST_USUARIO_PAINEL_VISTORIA = 2497 '3672 Produзгo
%>