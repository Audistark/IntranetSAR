<%
'Arquivo de configura��o de vari�veis

'URL do WebService do Pr�ton
'const urlWSProton = "http://www.webservicex.net/usaddressverification.asmx?WSDL" ' Teste para ver os erros
'const urlWSProton = "http://homologacao.anac.gov.br/proton/proton.asmx?op=" 'URL de Homologa��o
'const urlWSProton = "http://sdadf1004.anac.gov.br/wsproton/Proton.asmx?op=" 'URL de Desenvolvimento
'Const urlWSProton = "http://sigad.anac.gov.br/proton/Proton.asmx?op=" 'URL de Produ��o
Const urlWSProton = "none" 'n�o tem SIGAD

'--- usu�rios para cadastro
Dim cod_usuarioProton : cod_usuarioProton = "3168" 'SINTAC_HABILITACAO Desenv e Homolog � o mesmo ID
const CONST_USUARIO_PAINEL_VISTORIA = 2497 '3672 Produ��o
%>