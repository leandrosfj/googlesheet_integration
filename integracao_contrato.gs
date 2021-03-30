function onOpen() {

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Gerar Contrato');
  menu.addItem('Gerar Contrato Padrão','generateDefaultContract');
  menu.addToUi();
}

const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

function generateDefaultContract() {

  var docX = DriveApp.getFileById('1CALdrtGiyIwsobVrQRixXovQJlKcumBU');
  var newDoc = Drive.newFile();
  var blob = docX.getBlob();

  const destinationFolder = DriveApp.getFolderById('1jvk0wmlB4RBb-Ddg2b3XqnLBT27gIYml');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes Novos');
  const rows = sheet.getDataRange().getValues();

  
  rows.forEach(function(row,index) {
    
    if (index === 0) return;                                                                           // Se estiver vazio, ignora
    if (row[16]) return;                                                                               // Se já existir link gerado, ignora.

    var copy = Drive.Files.insert(newDoc,blob,{convert:true});                                         // Cria cópia do arquivo
    DocumentApp.openById(copy.id).setName('Contrato Certify RDL - ' + row[1]);                         // Altera nome
    var file = DriveApp.getFileById(copy.getId());                                                     // Abre arquivo
    file.moveTo(destinationFolder);                                                                    // Muda para pasta destino
    const doc = DocumentApp.openById(file.getId());                                                    // Reabre arquivo
    var body = doc.getBody();                                                                          // Seleciona conteúdo

    // Data de hoje
    const todayDate = new Date();
    const date = todayDate.getDate() + ' de ' + monthNames[todayDate.getMonth()] + ' de ' + todayDate.getFullYear();

    //Formatador de Real R$
    var realBR = new Intl.NumberFormat('pt-br', {
      style: 'currency',
      currency: 'BRL',
    });
    
    // Altera campos
    body.replaceText('{{razaoSocial}}', row[1].toString().toUpperCase());                                             // Coluna B
    body.replaceText('{{cnpj}}', row[2]);                                                                             // Coluna C
    body.replaceText('{{endereco}}', row[3]);                                                                         // Coluna D
    body.replaceText('{{rastreamento}}', realBR.format(row[8]) + ' por embarque de veículos terceiros rastreados;');  // Coluna I
    body.replaceText('{{logistica}}', realBR.format(row[9]) + ' por embarque na modalidade logística;');              // Coluna J
    body.replaceText('{{cadastro}}', realBR.format(row[11]) + ' por cadastro na política autônomo;');                 // Coluna L
    body.replaceText('{{consulta}}', realBR.format(row[12]) + ' por consulta na política autônomo;');                 // Coluna M
    body.replaceText('{{dataHoje}}', date);                                                                           // Data de hoje
    body.replaceText('{{assinatura}}', row[1].toString().toUpperCase());                                              // Coluna B
    
    // Salva e inclui na tabela o link gerado
    doc.saveAndClose();
    const url = doc.getUrl();
    sheet.getRange(index + 1, 17).setValue(url);
  })

  // TO DO: Consult. Club
  // TO DO: Fixo
  // TO DO: Assinatura
}