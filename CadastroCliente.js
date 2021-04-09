function CadastroCliente(){

  var Form = HtmlService.createTemplateFromFile("PaginaCadastro")

  var mostrarForm = Form.evaluate();

  mostrarForm.setTitle("Cadastro de Cliente").setHeight(600).setWidth(1200);

  SpreadsheetApp.getUi().showModalDialog(mostrarForm,"Cadastro de Cliente");
}

function consultaCNPJ(cnpj) {

  const cnpjNumeroStr = cnpj.replace(/[^0-9]/g,'');
  
  var response = UrlFetchApp.fetch("https://www.receitaws.com.br/v1/cnpj/" + cnpjNumeroStr);
  const resultado = JSON.parse(response.getContentText());

  return resultado;
}

const TABELA_CADASTRO_CLIENTES = "Cadastro de Clientes";
const INDEX_INCLUDE_DATA = 3;
const INDEX_INCLUDE_RAZAO_SOCIAL = 4;
const INDEX_INCLUDE_CNPJ = 5;
const INDEX_INCLUDE_ENDERECO = 6;
const INDEX_INCLUDE_CIDADE = 7;
const INDEX_INCLUDE_ESTADO = 8;

function salvarDados(resultObject) {

  var sheetCadastro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABELA_CADASTRO_CLIENTES);

  var row=4;    

  while(!sheetCadastro.getRange(row,INDEX_INCLUDE_CNPJ).isBlank()) {             //Encontra a próxima linha com CNPJ vazio
    row++; 
  }

  sheetCadastro.getRange(row,INDEX_INCLUDE_DATA).setValue(todayDate);
  sheetCadastro.getRange(row,INDEX_INCLUDE_RAZAO_SOCIAL).setValue(resultObject.nome);
  sheetCadastro.getRange(row,INDEX_INCLUDE_CNPJ).setValue(resultObject.cnpj);
  sheetCadastro.getRange(row,INDEX_INCLUDE_ENDERECO).setValue(resultObject.enderecoCompleto);
  sheetCadastro.getRange(row,INDEX_INCLUDE_CIDADE).setValue(resultObject.municipio);
  sheetCadastro.getRange(row,INDEX_INCLUDE_ESTADO).setValue(resultObject.uf);
}



