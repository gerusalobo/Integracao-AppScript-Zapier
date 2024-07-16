//Função para criação da Trigger junto ao Forms - roda apenas uma vez.

function setUpFormSubmitTrigger() {
    var form = FormApp.openById("1YLjK7yTrUaxApXccATR5JoIETKWRuwybpnHT8dakMOA");
    ScriptApp.newTrigger('dadosForms')
   .forForm(form)
   .onFormSubmit()
   .create();
}

//Função para pegar os dados do Forms e escrever nas abas: ResumoPedido e Pedidos

function dadosForms(e) {
 var formResponse = e.response;
 Logger.log(JSON.stringify(e));
 var itemResponses = formResponse.getItemResponses();

 // Obter as planilhas ativas
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
 var sheetresumo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ResumoPedido"); 
 var nextRow2 = sheetresumo.getLastRow() + 1;
 var body = "\n";

 // Pegar a data hora e formatar - Timestamp
 var currentDateTime = new Date();
 var formattedDateTime = Utilities.formatDate(currentDateTime, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

 // Iterar sobre cada resposta do formulário
 itemResponses.forEach(r => {
   var title = r.getItem().getTitle();
   var response = r.getResponse();  
   
   // Pegar os dados do Array e colocar na tabela Detalhes 
   if (title === "Eventos" && Array.isArray(response)) {
     // Iterar sobre o array de eventos
     response.forEach((evento, index) => {
       // Verificar se o evento é diferente de zero
       if (evento !== '0' && evento !== null) {
         var nextRow = sheet.getLastRow() + 1;
         var id = index + 1; // criar o id do evento pela posição no array
         sheet.getRange(nextRow, 1).setValue(formattedDateTime); // Coluna A - Timestamp
         sheet.getRange(nextRow, 2).setValue('Evento ' + id); // Coluna B - Evento
         sheet.getRange(nextRow, 3).setValue(evento); // Coluna C - Quantidade (Ingressos do evento)
         var nome = itemResponses.find(item => item.getItem().getTitle() === "Nome").getResponse();
         sheet.getRange(nextRow, 4).setValue(nome); // Coluna D - Nome
       }
     });
   }

   // Pegar os dados do Array e colocar na tabela Resumo 
   if (title === "Eventos" && Array.isArray(response)) {
     // Iterar sobre o array de eventos
     response.forEach((evento, index) => {
       // Verificar se o evento é diferente de zero
       if (evento !== '0' && evento !== null) {
         var id = index + 1; // criar o id do evento pela posição no array
         body += ('     # Evento ' + id + ': ' + evento + ' Ingressos\n');
       }
     });
     
     sheetresumo.getRange(nextRow2, 1).setValue(formattedDateTime); // Coluna A - Timestamp
     var nome = itemResponses.find(item => item.getItem().getTitle() === "Nome").getResponse();
     sheetresumo.getRange(nextRow2, 2).setValue(nome); // Coluna B - Nome
     sheetresumo.getRange(nextRow2, 3).setValue(body); // Coluna C - Resumo
     Logger.log(body);
   }
 });

 // Chama a função de cálculo
 calculo();
}

//Função para calcular os dados e criar a tabela resultado na aba: Calculo

function calculo() {
 const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 const calculoSheet = spreadsheet.getSheetByName("Calculo");
 const pedidosSheet = spreadsheet.getSheetByName("Pedidos");
 const tabelaSheet = spreadsheet.getSheetByName("tabelaEventos");

 const lastRow = pedidosSheet.getLastRow();

 //iterar sobre as celulas.
 for (var i=2; i <= lastRow; i++) {
  
 //Calculando o custo e o Recebido
 var custo = '=C' + i +'*VLOOKUP(B'+ i +',TabelaEventos!A:B,2,FALSE)';
 var pedido = '=C'+ i +'*VLOOKUP(B'+ i +',TabelaEventos!A:C,3,FALSE)';
 var saldo = '=F'+ i + '- E'+ i;

 //complementando os dados da lista de pedidos.
 pedidosSheet.getRange(i, 5).setValue(custo);
 pedidosSheet.getRange(i, 6).setValue(pedido);
 pedidosSheet.getRange(i, 7).setValue(saldo);
 }

 //Criar Pivot Table Resultado
 
 // Obter o intervalo dos dados
  const rangeDados = pedidosSheet.getRange("B1:G" + lastRow);

 // Obtendo os valores da coluna B e removendo os valores nulos
 const colunaB = rangeDados.getValues().map(row => row[0]).filter(value => value !== null && value !== '');
 const rangeDadosFiltrados = rangeDados.offset(0, 0, colunaB.length);

 // Resetando a aba Calculo
 calculoSheet.getRange(1,1,5,50).clear;
 calculoSheet.getRange(6,2,10,50).clear;

 // Definindo a célula onde a tabela dinâmica será inserida
 const celulaInicial = calculoSheet.getRange("A1");

 // Criando a tabela dinâmica
 const tabelaDinamica = celulaInicial.createPivotTable(rangeDadosFiltrados);

 tabelaDinamica.addRowGroup(2); // Agrupa pela segunda coluna do intervalo de dados (coluna B)
 tabelaDinamica.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Soma os valores da terceira coluna (coluna C)
 tabelaDinamica.addPivotValue(5, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Soma os valores da quinta coluna (coluna E)
 tabelaDinamica.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Soma os valores da sexta coluna (coluna F)
 tabelaDinamica.addPivotValue(7, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Soma os valores da sétima coluna (coluna G)

 //calcular a quantidade de ingressos minimos na tabela Eventos.
 for (var i=2; i <= 6; i++) {
 var qtdingressos = '=ROUNDUP(D'+i+'/(C'+i+'-B'+i+'))'
 tabelaSheet.getRange(i, 5).setValue(qtdingressos);
 }

 //calcular se houve lucro, indicar quantidade de ingressos para o Lucro e o Lucro - comparando com os custos fixos da tabela de Eventos e inserindo na aba calulo
 for (var i=2; i <= 6; i++) {
 var custo = '=VLOOKUP(A'+ i +',TabelaEventos!A:D,4,FALSE)'
 var ingressosfaltantes = '=IF((VLOOKUP(A' + i +',TabelaEventos!A:E,5,FALSE)-B' + i + ')>0,VLOOKUP(A' + i +',TabelaEventos!A:E,5,FALSE)-B' + i + ',"")'
 var calculolucro = '=E'+ i +'-VLOOKUP(A'+ i +',TabelaEventos!A:D,4,FALSE)';
 var testelucro = '=IF(H'+ i +'> 0 ,true,false)';

 calculoSheet.getRange(i, 6).setValue(custo);
 calculoSheet.getRange(i, 7).setValue(ingressosfaltantes);
 calculoSheet.getRange(i, 8).setValue(calculolucro);
 calculoSheet.getRange(i, 9).setValue(testelucro);
 }

 //calcular totalizadores
 var custototal = '=SUM(F2:F6)';
 var ingressos = '=SUM(G2:G6)';
 var lucrototal = '=SUM(H2:H6)';
 var teste = '=and(I2:I6)';
 calculoSheet.getRange(7, 6).setValue(custototal);
 calculoSheet.getRange(7, 7).setValue(ingressos);
 calculoSheet.getRange(7, 8).setValue(lucrototal);
 calculoSheet.getRange(7, 9).setValue(teste);

}