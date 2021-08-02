// EDITAR TABELA GOOGLE SHEETS EM HTML, UTILIZANDO BOOTSTRAP-TABLE
// Janeiro - 2021;
 
function doGet() {

  return HtmlService.createHtmlOutputFromFile('page')
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

}


// função chamada pelo botão da Planilah.
function executarHTML() {
 
  var html = HtmlService.createHtmlOutputFromFile('page')
    .setWidth(1360)
    .setHeight(755)
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'EXEMPLO DE INDICADORES');

}




//--------------------------------------------------------------------------
// Lista das Empresas:
// somente os de STATUS em branco.
//--------------------------------------------------------------------------
function getEmpresas() {
  try {
    let aEmpresas = []
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadastros");
    var values = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues();
    for (i = 0; i <= values[0].length - 2; i+=2) {
      if ( values[0][i+1] == '' ) {  // se não estiver em branco, ter como DESATIVADO!!
        aEmpresas.push(values[0][i])
      }
    }
    return aEmpresas;
  }
  catch(err) {
    Logger.log(err);
  }
}
//--------------------------------------------------------------------------

//--------------------------------------------------------------------------
// Lista dos DEPTo.s da Empresa solicitada:
// somente os de STATUS em branco:
//--------------------------------------------------------------------------
function getDepartamentos(nomeEmpresa) {
  try {
    row = 2  // linha do "empresa"
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cadastros");
    var data = sheet.getDataRange().getValues();
    var Departamentos = []
    var col = data[row].indexOf(nomeEmpresa);
    if (col != -1) {
      for (i = 3; i <= data.length -1; i++) {  // vai até o tamanho da maior;
        if (data[i][col] != '' &  data[i][col+1] == '' ) { // se não estiver em branco, ter como DESATIVADO!!
          Departamentos.push(data[i][col])
        }
      }
    }
    return Departamentos
  }
  catch(err) {
    Logger.log(err);
  }
}
//--------------------------------------------------------------------------


//--------------------------------------------------------------------------
// pegar o ANO para os lançamentos:.....
function getAno() {
  var ano = [SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ano").getRange(2, 2,).getValue()]
  return ano
}
//--------------------------------------------------------------------------


//--------------------------------------------------------------------------
// pegar última linha com dados de uma determinada coluna.....
function getLastDataRow(sheet) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A" + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }              
}
//--------------------------------------------------------------------------


//--------------------------------------------------------------------------
function getDados(empresa, departamento, ano, mes) {

  var dados = getDadosIndicadores(empresa, departamento, ano, mes)

  if (dados.length <= 1) {  // se vazio (ou só o header), pegar o MODELO
    dados = getModelo()
  }

  return dados

}
//--------------------------------------------------------------------------

//--------------------------------------------------------------------------
// Buscar os dados solicitados:
function getDadosIndicadores(empresa, departamento, ano, mes) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados")
  var range = sheet.getRange( 1, 1, sheet.getLastRow(), sheet.getLastColumn() )
  var d = range.getValues()
  var k = range.getRowIndex() + 1  // pular 1 = tirar o header...

  //Preencher a coluna[0]=Index (com o mesmo valor da LINHA)), em TODO "d"
  //  #TO-DO: se ficar lento pra grande quantidade de dados, fazer após o filtro, com dados menor!
  for ( i = 1; i < d.length; i++ ){
    d[i][0] = k++;
  }

  // concatenar todos os critérios em um só REGEX...
  var S = '('   + empresa + ')' +
          '(.+' + departamento +')' +
          '(.+' + ano +')' +
          '(.+' + mes +')'
  var reg = S;
  var T = d.filter(function(item) {
    return (item.toString().search(reg) != -1) ||
           (item[0] == 'id')  // trazer o header também...
  });

  return T

}
//--------------------------------------------------------------------------


//--------------------------------------------------------------------------
// Salvar os dados:
function setDados(tabela) {

  tabela.splice(0, 1); // remove header

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados")

  // se INDEX = 0, então é o modelo. Adicionar ao LastRow() e atualizar INDEX;
  if (tabela[0][0] == 0) {
    //Preencher a coluna[0]=Index (com o mesmo valor da LINHA))
    var k = sheet.getLastRow() + 1
    if (k < 5) // se vazio ainda... segunda linha em diante... 
        k = 2;

    for ( i = 0; i < tabela.length; i++ ){
      tabela[i][0] = k++;
    }
  }

  sheet.getRange(tabela[0][0], 1, tabela.length, tabela[0].length ).setValues( tabela );   // escrever....

  SpreadsheetApp.flush();

}
//--------------------------------------------------------------------------



//--------------------------------------------------------------------------
// Buscar Modelo:
function getModelo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Modelo")
  var range = sheet.getRange( 1, 1, sheet.getLastRow(), sheet.getLastColumn() )
  var d = range.getValues()
  // var k = range.getRowIndex()
  // //Preencher a coluna[0]=Index (com o mesmo valor da LINHA))
  // for ( i = 0; i < d.length; i++ ){
  //   d[i][0] = k++;
  // }
  return d
}
//--------------------------------------------------------------------------



//--------------------------------------------------------------------------
// Função Filtrar - Google:
//https://sites.google.com/site/nnillixxsource/Notable/PivotChartsLib/ArrayLib/filterByText_js
// Adaptado para pesquisar TODOS - (E) - dos strings passados no vetor "values[]"
//-------------------------------------------------------------------------
function filterByText(data, values) {

    var r = [];

    if ((data.length > 0) & (values.length > 0)) {

      // concatenar todos os critérios em um só REGEX...
      var S = '(' + values[0] + ')'  // 1ª expressão...
      for (var j = 1; j < values.length; j++) {
          S += '(.+' + values[j] +')'
      }

      var reg = S;

      for (var i = 0; i < data.length; i++) {
        if ( data[i].toString().search(reg) != -1 ) {

          //  adicionar header 
          if (r.length == 0) {
            r.push(data[0]);
          }

          r.push(data[i]);
 
          if ( (i % 2 ) == 1 ) {

            var zzz = r[ r.length - 1 ][5]
            var kkk = r[ i ][5]

            r[ r.length - 1 ][5] = ''  // nas linhas ímpares, deixar em branco... 

            var aaa = r[ r.length - 1 ][5]
            var bbb = r[ i ][5]

             var t = 0

          }


        }
      }
    }

    if (r.length == 0) {
      r = getModelo()
    }

    return r;

}
//--------------------------------------------------------------------------

