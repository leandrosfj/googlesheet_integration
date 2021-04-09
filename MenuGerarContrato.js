function onOpen() {

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Gerar Contrato');
  menu.addItem('Gerar Contrato Padrão','generateDefaultContract');
  menu.addToUi();
}

const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
const MODELO_ID = "135z0ZBfjqoEbGXY8R_tVvqqS0LojCGlZgnkyL3UGup0";
const PASTA_DESTINO_ID = "1jvk0wmlB4RBb-Ddg2b3XqnLBT27gIYml"
const TABELA_CONTRATOS = "Contratos";
const INDEX_RAZAO_SOCIAL = 3;
const INDEX_CNPJ = 4;
const INDEX_ENDERECO = 5;
const INDEX_VALOR_RASTREAMENTO = 18;
const INDEX_VALOR_LOGISTICA = 19;
const INDEX_VALOR_FIXO = 20;
const INDEX_VALOR_CADASTRO = 21;
const INDEX_VALOR_CONSULTA = 22;
const INDEX_LINK_CONTRATO = 25;

// Data de hoje
const todayDate = new Date();
const dataHoje = todayDate.getDate() + ' de ' + monthNames[todayDate.getMonth()] + ' de ' + todayDate.getFullYear();

//Formatador de Real R$
var realBR = new Intl.NumberFormat('pt-br', {
  style: 'currency',
  currency: 'BRL',
});

function generateDefaultContract() {

  var docX = DriveApp.getFileById(MODELO_ID);
  var newDoc = Drive.newFile();
  var blob = docX.getBlob();

  const destinationFolder = DriveApp.getFolderById(PASTA_DESTINO_ID);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABELA_CADASTRO_CLIENTES);
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row,index) {
    
    if (index === 0) return;   
    if (index === 1) return;                                                                           //Se for primeira linha, ignora.
    if (row[INDEX_LINK_CONTRATO]) return;                                                              // Se já existir link gerado, ignora.

    var copy = Drive.Files.insert(newDoc,blob,{convert:true});                                         // Cria cópia do arquivo
    DocumentApp.openById(copy.id).setName('Contrato Certify RDL - ' + row[INDEX_RAZAO_SOCIAL]);        // Altera nome
    var file = DriveApp.getFileById(copy.getId());                                                     // Abre arquivo
    file.moveTo(destinationFolder);                                                                    // Muda para pasta destino
    const doc = DocumentApp.openById(file.getId());                                                    // Reabre arquivo
    var body = doc.getBody();                                                                          // Seleciona conteúdo
    
    // Altera campos
    body.replaceText('{{razaoSocial}}', row[INDEX_RAZAO_SOCIAL].toString().toUpperCase());                                                   
    body.replaceText('{{cnpj}}', row[INDEX_CNPJ]);                                                                                           
    body.replaceText('{{endereco}}', row[INDEX_ENDERECO]);                                                                                    
    
    if (row[INDEX_VALOR_RASTREAMENTO]) {

      body.replaceText('{{rastreamento}}', realBR.format(row[INDEX_VALOR_RASTREAMENTO]) + ' por embarque de veículos terceiros rastreados;');         
    } else {
      body.replaceText('{{rastreamento}}', "");
    } 

    if (row[INDEX_VALOR_LOGISTICA]) {

      body.replaceText('{{logistica}}', realBR.format(row[INDEX_VALOR_LOGISTICA]) + ' por embarque na modalidade logística;');       
    } else {
      body.replaceText('{{logistica}}', "");
    }

    if (row[INDEX_VALOR_FIXO]) {

      body.replaceText('{{fixo}}', realBR.format(row[INDEX_VALOR_FIXO]) + ' por veículo fixo rastreado;');       
    } else {
      body.replaceText('{{fixo}}', "");
    }

    if (row[INDEX_VALOR_CADASTRO]) {

      body.replaceText('{{cadastro}}', realBR.format(row[INDEX_VALOR_CADASTRO]) + ' por cadastro na política autônomo;');       
    } else {
      body.replaceText('{{cadastro}}', "");
    }

    if (row[INDEX_VALOR_CONSULTA]) {

      body.replaceText('{{consulta}}', realBR.format(row[INDEX_VALOR_CONSULTA]) + ' por consulta na política autônomo;');       
    } else {
      body.replaceText('{{consulta}}', "");
    }
  
    body.replaceText('{{dataHoje}}', dataHoje);                                                                                              
    body.replaceText('{{assinatura}}', row[INDEX_RAZAO_SOCIAL].toString().toUpperCase());                                                    
    
    // Salva e inclui na tabela o link gerado
    doc.saveAndClose();
    const url = doc.getUrl();
    sheet.getRange(index + 1, INDEX_LINK_CONTRATO + 1).setValue(url);
  })
}