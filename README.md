# Integração: Forms - GoogleSheets - Discord usando AppScript & Zapier 



Data: 25/05/2024

Pós Tech Fiap - Dev Foundations - Tech Challenge 1

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXeuQI68JATP91uCyV3JNK51pkaWhB_OgCmtVre-YOobyXh2xVQcT2xUI_q8-12Gfv02Fuk0SnGKjuQ8A64wYzhH5ydc7rkkpKV78ftPfQvAaegD7bV8Vf7lMasHtpwHP2TOA1MvBaX-tLWuulNk9vulwmma?key=umjyYsAnoWfRwPFluedkpA)

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXchpYKZ0Bh3wOlvQgt38QSlvncYWKTy_BWf_2uRJLys2II_08RUCOg80fOg3Iopk7wDc-Ov6zyQL_KlfOdgqakEohkGvAX0s_Exn2GP1Eaop2wbN9QLsue7kEygJqDxicoui-mjE0jCifUBaQ1DMOslFx69?key=umjyYsAnoWfRwPFluedkpA)

Mais algumas considerações, conforme esclarecimentos do Discord:

- Deve ser possível comprar ingressos para mais de um evento;

- A mensagem no Discord deve ser o resumo da compra. Apenas uma mensagem por compra;

- Foi considerado um custo fixo para os eventos, de forma a calcular a quantidade mínima de ingressos, o valor arrecadado e se houve lucro.

# 

# Google Forms:

https://forms.gle/h89C7GQ1tyv6FGHdA

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXeSEcO2NclI-adrZ3I6ceQ6o4rkummbH0suc1Rrf-W070sRz1pf1g-mHGXiLJggBSU1PeTdpgMSQ_jenDuRe20p2PjBeBmM213-ljiVN26fJa6IBiC7A8R0BiYoZWn6OiysSZwvglk2dJV0fwmFa2ENbc7r?key=umjyYsAnoWfRwPFluedkpA)

![Forms response chart. Question title: Nome. Number of responses: 12 responses.](https://lh7-us.googleusercontent.com/docsz/AD_4nXfXQlAMOvm0TwA_eSpJsDb5uk1OewtBbt2MJWrgHCnJL0AKMuy_FgZ3tKLu13y1qPTJtav7ciMWjoRyZ47NaSltQXK1L-suPNO0JHLFt7SQdXEpLXAk3xzITD_qY5IivZdRQQCAEx_xFg-8zRDOtaRAjInn?key=umjyYsAnoWfRwPFluedkpA)

![Forms response chart. Question title: Eventos. Number of responses: .](https://lh7-us.googleusercontent.com/docsz/AD_4nXeU5QpOm11L3fowOIElPklaDEn2SCEYdMqL9sqtNzZKQI4XmLSrg1lllsbA3F9NURIlEtYBVyBTNIlaXnY3OKAiVl-dmrAx5h2z2iU5-vW8W5b_YvrDxCFsgNclBn_4_VYHKzEDxt-BzvCEu6deLSyCEcE?key=umjyYsAnoWfRwPFluedkpA)

Observações e Considerações sobre o Formulário:

- Apesar do formulário permitir a submissão de formulário sem eventos, eles não são registrados como pedidos. Há um tratamento de erro para isso.

- O formulário em questão permite a seleção de múltiplos eventos com a indicação da quantidade para cada um.

# Trigger

Para o vínculo do Google Forms no Google Sheets foi criado uma função no app script de vínculo.

```
//Função para criação da Trigger junto ao Forms - roda apenas uma vez.

function setUpFormSubmitTrigger() {
     var form = FormApp.openById("1YLjK7yTrUaxApXccATR5JoIETKWRuwybpnHT8dakMOA");
     ScriptApp.newTrigger('dadosForms')
    .forForm(form)
    .onFormSubmit()
    .create();}
```



![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXdnahTX2AxjZtH5vODKbpZD6lrMzGzwgVxIw83h6h3r_zj1GhiQkXyZmRLqIOaX7YU-v77U1eHiCZ5ehNX7k60kYFTYVhCGucTBq2jyN_GSiyQQzKjAKEANmJM9KHyCIC0GGaWVfucOfqQFxkt8nJX3UGj9?key=umjyYsAnoWfRwPFluedkpA)

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXcPiinO-a8-AJxE2AlQmKrSG7m_iWzPBrqE4_tUcLvZeFfIeMjApy5kAggF2K_N6D-FZ3cEukyuwLo5nkckLaGpGSEp2cYULCk335v7sWVtfNIs01ofraLvFsnyhHTTrfOTLg4b5QhmkIL2--30JH4zpkyK?key=umjyYsAnoWfRwPFluedkpA)

Execuções:

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXdInaEYha-ccv5Wogx41-2M6ea2_8IYfhsWcAQJd1y3Xr-S8O7lSJC49_HpVv_M2pkSd6hIN8Is7lS7gMik9C0EYIWqoGFFzrgf0kkw4YS0d63KpMDP9awcjFEmsvbZ22YryHeSHnXctsjhcpvhAZqHboWL?key=umjyYsAnoWfRwPFluedkpA)

# 

# Planilha

[Link Planilha](https://docs.google.com/spreadsheets/d/114G-TALOQ0yWi2uk8eptQfwEeqfFPtIfLEPlOnB-2Io/edit?usp=sharing)

A planilha usada tem 4 Abas:

Aba 1: Tabela de Eventos

Aba 2: Resumo Pedido

Aba 3: Pedidos

Aba 4: Calculo

Aba 1: Tabela de Eventos

Nessa aba tem os dados dos eventos com os valores de custo por ingresso, preço do Ingresso, custo fixo, e onde é calculado a quantidade mínima de ingressos para ter Lucro, considerando o preço, e os custos envolvidos.

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXd0yUdiIR8zCSlLPq9EkqE-7S-m3hPu_zT7ucfDppVYWz-ouCwSzwTccRfur_R8vNmZZ2cFd4GD2abQ99RsSxpshyH3Rxsry49foOonEW4qR1eNeMzcptD-T56bFgguHLvVlJbW9bZ7qnO4q6IIwp5xW7SD?key=umjyYsAnoWfRwPFluedkpA)

Aba 2: Resumo Pedido

Na aba 2 é registrado o resumo dos pedidos. Essa é a aba de conexão com o Zapier.

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXcavaFk1VqXr42QZXvDi9aVPtH6YZGQIApNwH_i_bGCXB-6ZihtaegycEeOdtlM_-P2hdsVvc33jCo1fJCqRwsS-njtC9FgZKAW2c6HGRrr_mghHJNZznrmbfV4Q7_hMICesE6Zc-M4ol2M1s1cPuoc2uw?key=umjyYsAnoWfRwPFluedkpA)



Aba 3: Pedidos

Na aba Pedidos, é segmentado os itens adquiridos dos pedidos, e já calculado o custo, valor pago e saldo total (saldo = valor pago - custo do ingresso) por item.

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXdFPbNsCS8Gwh_fqP_jQ51VCf9ZhjgBd5pXB0RA8rgGRRUWk3BZv1g-ATv3aMnn1dtcbSGUAiqsAaXfnRkuHGThV0EzoNE3ceDw8lQDGEOmHWVpQyS9iKHl0X4vyvq9GefjUg5OUR5JC5bqeBr2cCY-Kak?key=umjyYsAnoWfRwPFluedkpA)

Aba 4: Calculo

A aba 4 tem a tabela resultado com todos os cálculos solicitados: Soma dos totais por evento, com quantidade total, custo total de ingresso, valor recebido dos ingressos, saldo (recebido - custo por ingresso), Custo Fixo indicado da tabela da Aba 1, quantidade de ingressos para o Lucro, Cálculo do Lucro e se tem Lucro.

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXcfRfKFhmrq1oi_AclHkd1tNr1VpQyo8PZ7C0lQBAlmSkTCRHLnX7J6G3o3iTjYehv_RE4wsaPqHdokQjfF3Gqvq0AxNzBQijUHjTguF_x46xsbUcaV7dSSf_oZzmFZov0mwAByITUpjkv1RAtRoVNZp20B?key=umjyYsAnoWfRwPFluedkpA)

Observação: Foi incluído um valor de custo Fixo para permitir o cálculo de Lucro.

# Códigos AppScript

Função dadosForms()

Essa função inicia através da trigger configurada previamente, recebendo os dados do Forms, tratando a informação e escrevendo nas abas: ResumoPedido e Pedidos.

Ao final, essa função chama a função de cálculo.

```
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
```

Função Cálculo()

Essa função implementa todos os cálculos da planilha:

Na primeira etapa: Ela adiciona na tabela Pedidos, os valores de custo, valor recebido (pedido) e saldo.

Na sequência ela cria uma tabela dinâmica (pivot table) com os totalizados na aba: cálculo

E complementa essa tabela com mais 3 colunas calculadas: 

- Custo Fixo - da tabela de Eventos
- Quantidade de Ingressos para ter Lucro
- O valor do Lucro/Prejuízo
- A indicação se tem Lucro.

Adicionalmente, essa função também calcula a quantidade de ingressos mínimos para ter Lucro, na Tabela de Eventos.

```
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
```

Observações:

- Todas as fórmulas da planilha são criadas via código na função cálculo().

- Em um esclarecimento do discord, foi permitido o uso de pivot table para cálculo da tabela resultado.

# Zapier

A conexão com o Zapier é feita com a Planilha na aba Resumo Pedidos.

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXdb6_J6ml8Wvd21yT_1P4zo6gg1EGl3YWFWpSFnAW0qSNX6fvOyVohF7jSrePvlKgE5XoaDtwbuPd86qWwx12OM2IbVeMma7SOj07bhSIiqWvO0oS9FyoYS084XJkqyeEDrW-4pmClSkqcHYyvJLh4m7Zj0?key=umjyYsAnoWfRwPFluedkpA)![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXcF_7iB1O7r5kTM2vI8rTZN3PjWcd2akPlixQC7FZItW1RstqHXwIrwPoqDPVZmO6jsUNXY6cfIdAXxzYaaR8LEm0pLP845PX1MFVFO1d9FvBdgYHkqM4SgUwt7DObkkP3GTTvtnUN-4VJonUP5ko0jmQtC?key=umjyYsAnoWfRwPFluedkpA)

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXddJOcZIdODMQ_rZvMTbPBNqKpJS5VkUiwJK5xLpHDobJDTppi4QdEX-119B5ozlg7AI4yo46T5JflnZdT4nsuZX4M3ysueH-m5PJ7luiC54GC0AjhCbPKJin1YrhwZtRZueHnu1RTAT3lbfzKKHx7mFId2?key=umjyYsAnoWfRwPFluedkpA)

​	

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXewCkhpfcWwsChVBmEArMIw0lTUlpo35uI_R0DcRQ7_GOwmjR07SbVtMChhv8JelH6t3hzea5Rx0Yr2cMbEYNRSHt5bGhhS9TPATAyd74eupt-WEADPcnNPDGdCE7AV0M5xcRhifN2zmxJdzS3TQl9XJ1s?key=umjyYsAnoWfRwPFluedkpA)





E com o Discord Tech Brabos no canal testes-tech-challenge

![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXeEuaR4okrIBdTcjmKQ3Q0d6VsGtikOyRjY9eXl_xisb4__0f2WCkj3pL-ekNrJ1twHtVVv4gW1PeTJkkTxkiVXRZj1xHUgIM9x1LQwXtfwVCOdqDZj94svgJYWFMPJNSUazinn8LnNvWeeHO_FoC7CbzSr?key=umjyYsAnoWfRwPFluedkpA)![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXcBIGkd6riXopqjd7CRNEtm9DuzHgqyJjw1OWlgP72ZGe6VDCwt-oz6BMm_U78HtwRWXiGnxgFFAthkzEX0kR1X-lEN7HpxTH5P1BX8gbmEVZ7qn_elBo02GQm7q-zvKRwMBi3W07369oFN9KclOS9kdvfl?key=umjyYsAnoWfRwPFluedkpA)![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXdLkwFEPAFwV-hDygaQKljrvSDFijQsP6b9m08u9bkIGoqXlrEa59UfRb1-Jt5hm_iNuNJb034gT2PohSdcva5anwkBjwt6ZasIrYIKJZ8xvuCwJp_gurleRKK8XU9CiRNw0LFuL9tcvGGltCCMGCKrjL_O?key=umjyYsAnoWfRwPFluedkpA)![img](https://lh7-us.googleusercontent.com/docsz/AD_4nXe7U9K_JQZ6xwfqlMlbdlg4136qbtnqPmGwM_oAXcAttZeDOXkr7W2ZS_4EhPtHogtsIWU0g82puG8EbcjMDmhyyyAnuiH9eBdShbnm7ExM25nFQSwM7P6nPOtCMn75Xu3eNDuZHYsJHImJet9UCloFch1c?key=umjyYsAnoWfRwPFluedkpA)

 
