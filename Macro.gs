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
    if(dados[linha][colunaID].toString().toLowerCase() == criteriopesquisa.toString().toLowerCase() || dados[linha][1].toString().toLowerCase() == criteriopesquisa.toString().toLowerCase()){
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
