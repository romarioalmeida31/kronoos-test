import fs from 'fs'
import csvParser from 'csv-parser'
import * as excel4node from 'excel4node'


const conversionToReal = valor => {
  let numero = new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  }).format(valor);

  return numero
}


const validateCPF = cpf => {


  cpf = cpf.replace(/[^\d]+/g, '');


  if (cpf.length !== 11) return false;

  let sum = 0;
  let remainder;


  for (let i = 0; i < 9; i++) {
    sum += parseInt(cpf.charAt(i)) * (10 - i);
  }


  remainder = (sum * 10) % 11;


  if (remainder === 10 || remainder === 11) remainder = 0;


  if (remainder !== parseInt(cpf.charAt(9))) return false;

  sum = 0;
  for (let i = 0; i < 10; i++) {

    sum += parseInt(cpf.charAt(i)) * (11 - i);
  }


  remainder = (sum * 10) % 11;


  if (remainder === 10 || remainder === 11) remainder = 0;


  if (remainder !== parseInt(cpf.charAt(10))) return false;

  return true;
}


const validateCNPJ = cnpj => {


  cnpj = cnpj.replace(/[^\d]+/g, '');


  if (cnpj.length !== 14) return false;


  if (/^(\d)\1+$/.test(cnpj)) return false;

  let sum = 0;
  let factor = 5;


  for (let i = 0; i < 12; i++) {
    sum += parseInt(cnpj.charAt(i)) * factor;
    factor = factor === 2 ? 9 : factor - 1;
  }

  let remainder = sum % 11;
  let firstVerifier = remainder < 2 ? 0 : 11 - remainder;


  if (parseInt(cnpj.charAt(12)) !== firstVerifier) return false;

  sum = 0;
  factor = 6;


  for (let i = 0; i < 13; i++) {
    sum += parseInt(cnpj.charAt(i)) * factor;
    factor = factor === 2 ? 9 : factor - 1;
  }


  remainder = sum % 11;
  let secondVerifier = remainder < 2 ? 0 : 11 - remainder;
  return parseInt(cnpj.charAt(13)) === secondVerifier;
}


const validatePrest = (vlPrest, qPrest, vlTotal) => {
  return vlPrest * qPrest === vlTotal;
}

const conversionDate = num => {
  const dateFormat = `${num.slice(0, 4)}-${num.slice(4, 6)}-${num.slice(6, 8)}`;
  return new Date(dateFormat);
}

const results = [];

fs.createReadStream('data.csv')
  .pipe(csvParser())
  .on('data', (row) => {
    const originalRow = { ...row };


    row.validaCpfCnpj = validateCPF(row.nrCpfCnpj) || validateCNPJ(row.nrCpfCnpj);


    row.validaPrest = validatePrest(row.vlPresta, row.qtPrestacoes, row.vlTotal)


    row.dtContrato = conversionDate(row.dtContrato)
    row.dtVctPre = conversionDate(row.dtVctPre)

  
    row.vlTotal = conversionToReal(row.vlTotal);
    row.vlPresta = conversionToReal(row.vlPresta);
    row.vlMora = conversionToReal(row.vlMora);
    row.vlMulta = conversionToReal(row.vlMulta);
    row.vlOutAcr = conversionToReal(row.vlOutAcr);
    row.vlIof = conversionToReal(row.vlIof);
    row.vlDescon = conversionToReal(row.vlDescon);
    row.vlAtual = conversionToReal(row.vlAtual);

    results.push(row);

  })
  .on('end', () => {
    const data = results
    const xl = excel4node;
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('Worksheet Name')

    const headingColumnNames = [
      "nrInst",
      "Agencia:",
      "cod Cliente:",
      "Numero Cliente",
      "Cpf/Cnpj",
      "Contrato:",
      "data Contrato",
      "qtPrestacoes:",
      "vlTotal",
      "cdProduto:",
      "dsProduto",
      "cdCarteira",
      "dsCarteira",
      "nrProposta",
      "nrPresta",
      "tpPresta",
      "nrSeqPre",
      "dtVctPre",
      "vlPresta",
      "vlMora",
      "vlMulta",
      "vlOutAcr",
      "vlIof",
      "vlDescon",
      "vlAtual",
      "idSituac",
      "idSitVen",
      "validaCpfCnpj",
      "validaPrest"

    ]

    let headingColumnIndex = 1;
    headingColumnNames.forEach(heading => {
      ws.cell(1, headingColumnIndex++).string(heading)
    });

    let rowIndex = 2;
    data.forEach(record => {
      let columnIndex = 1;
      console.log(columnIndex)
      Object.keys(record).forEach(columnName => {
        ws.cell(rowIndex, columnIndex++)
          .string(record[columnName])
      });
      rowIndex++;
    });

    wb.write('kronoos.xlsx')
  });


