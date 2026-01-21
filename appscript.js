const SUBDOMAIN = 'SEU_SUBDOMINIO'; 
const ACCESS_TOKEN = 'SEU_ACCESS_TOKEN';

function localizarDuplicatasInteligentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.clear();
  sheet.appendRow(["ID do Contato", "Nome", "Telefone Original", "Chave de Comparação", "Status"]);

  let page = 1;
  let contatosProcessados = {}; 
  let duplicatasEncontradas = [];
  let hasNext = true;

  while (hasNext) {
    const url = `https://${SUBDOMAIN}.kommo.com/api/v4/contacts?limit=250&page=${page}`;
    const options = {
      method: 'get',
      headers: { 'Authorization': `Bearer ${ACCESS_TOKEN}` },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) break;

    const data = JSON.parse(response.getContentText());
    if (!data._embedded) break;

    data._embedded.contacts.forEach(contato => {
      let telBruto = extrairTelefone(contato);
      if (telBruto) {
        let chaveEstrategica = gerarChaveComparacao(telBruto);
        
        if (contatosProcessados[chaveEstrategica]) {
          duplicatasEncontradas.push([
            contato.id, 
            contato.name, 
            telBruto, 
            chaveEstrategica,
            "Duplicado (Original ID: " + contatosProcessados[chaveEstrategica] + ")"
          ]);
        } else {
          contatosProcessados[chaveEstrategica] = contato.id;
        }
      }
    });

    page++;
    if (page > 50) hasNext = false; 
  }

  if (duplicatasEncontradas.length > 0) {
    sheet.getRange(2, 1, duplicatasEncontradas.length, 5).setValues(duplicatasEncontradas);
  }
  SpreadsheetApp.getUi().alert("Processo concluído!");
}


function gerarChaveComparacao(telefone) {

  let num = telefone.toString().replace(/\D/g, '');


  if (num.startsWith('55') && (num.length === 12 || num.length === 13)) {
    num = num.substring(2); 
    return tratarNonoDigitoBr(num);
  } 
  

  if (num.length === 10 || num.length === 11) {
    return tratarNonoDigitoBr(num);
  }


  return num; 
}

function tratarNonoDigitoBr(num) {
  if (num.length === 11 && num[2] === '9') {

    return num.substring(0, 2) + num.substring(3);
  }
  return num;
}

function extrairTelefone(contato) {
  if (!contato.custom_fields_values) return null;
  const field = contato.custom_fields_values.find(f => f.field_code === 'PHONE');
  return (field && field.values) ? field.values[0].value : null;
}