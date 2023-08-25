/*
  Direitos Autorais (c) 2023 Renata Veras Venturim

  Licença MIT

  A permissão é concedida, gratuitamente, a qualquer pessoa que obtenha uma cópia
  deste software e dos arquivos de documentação associados (o "Software"), para negociar
  no Software sem restrições, incluindo, sem limitação, os direitos de uso, cópia,
  modificação, fusão, publicação, distribuição, sublicenciamento e/ou venda de cópias do Software,
  e de permitir que as pessoas a quem o Software é fornecido o façam, sujeitas às seguintes condições:

  O aviso de direitos autorais acima e este aviso de permissão devem ser incluídos em todas as cópias
  ou partes substanciais do Software.

  **Atribution Clause:**
  Os usuários deste software são obrigados a manter os créditos originais e as atribuições do projeto
  em quaisquer cópias ou derivações do Software.

  O SOFTWARE É FORNECIDO "COMO ESTÁ", SEM GARANTIA DE QUALQUER TIPO, EXPRESSA OU IMPLÍCITA,
  INCLUINDO, MAS NÃO SE LIMITANDO ÀS GARANTIAS DE COMERCIALIZAÇÃO, ADEQUAÇÃO A UM PROPÓSITO ESPECÍFICO
  E NÃO VIOLAÇÃO. EM NENHUM CASO OS AUTORES OU DETENTORES DOS DIREITOS AUTORAIS SERÃO RESPONSÁVEIS
  POR QUALQUER REIVINDICAÇÃO, DANOS OU OUTRAS RESPONSABILIDADES, SEJA EM AÇÃO DE CONTRATO, DELITO
  OU DE OUTRA FORMA, DECORRENTES DE, OU EM CONEXÃO COM O SOFTWARE OU O USO OU OUTRAS NEGOCIAÇÕES NO PROGRAMAS.
*/

var url = "https://docs.google.com/spreadsheets/d/1do17u9OwnWRzyCMaDWWK5d0SjZqxbkFFdSF877zlCW0/edit?usp=sharing"

var nomeguia = "Planilha de Controle";
var linhainicial = 2;
var colunainicial = 1;
var linhacabecalho = 1;

var colunaID = 1;

function PesquisarDados(criteriopesquisa){
  var planilha = SpreadsheetApp.openByUrl(url);
  var guiadados = planilha.getSheetByName(nomeguia); 
  var dados = guiadados.getRange(linhainicial, colunainicial, guiadados.getLastRow()-linhacabecalho,13).getValues(); 
  var resultados = [];

  for(var linha = 0; linha<dados.length; linha++){
    if(dados[linha][colunaID].toString().trim().toUpperCase() == criteriopesquisa.trim().toString().toUpperCase() || dados[linha][1].toString().toUpperCase() == criteriopesquisa.toString().trim().toUpperCase()){
      var Carregar = {};

      Carregar.Campo1 = dados[linha][colunaID];
      Carregar.Campo2 = dados[linha][0]; 
      Carregar.Campo3 = dados[linha][3];      
      Carregar.Campo4 = dados[linha][6];         
      Carregar.Campo5 = dados[linha][9];
      
      if (dados[linha][10] !== "") {
        Carregar.Campo6 = Utilities.formatDate(dados[linha][10], Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        Carregar.Campo6 = "";
      }
      
      Carregar.Campo7 = dados[linha][11];  
      Carregar.Campo8 = dados[linha][12];

      resultados.push(Carregar);
    }
  }

  if (resultados.length > 0) {
    return resultados;
  } else {
    return ["Não encontrado!"];
  }
}
