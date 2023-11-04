const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');

class ExcelWriter {
  async createExcelFile(data, outputFileName) {
    try {
      const workbook = await XlsxPopulate.fromBlankAsync();      
      const sheet = workbook.sheet(0);

      // Escribimos la cabecera
      this.writeHeader(sheet);

      // Sepramos ingresos y gastos
      let ingresos = data.filter(item => item.Importe > 0);
      let gastos = data.filter(item => item.Importe < 0);
      // Dibujamos ingresos
      ingresos.forEach((item, index) => {
        sheet.cell(`A${index + 2}`).value(item.Movimiento);
        sheet.cell(`B${index + 2}`).value(item.Importe);
      });

      // Dibujamos gastos   
      gastos.forEach((item, index) => {
        sheet.cell(`C${index + 2}`).value(item.Movimiento);
        sheet.cell(`D${index + 2}`).value(item.Importe);
      });

      // Calcular el total de ingresos en la celda E2
      sheet.cell("E2").formula(`SUM(B2:B${ingresos.length + 1})`);

      // Calcular el total de gastos en la celda F2
      sheet.cell("F2").formula(`SUM(D2:D${gastos.length + 1})`);
      // Calculamos el
      sheet.cell("G2").formula(`SUM(E2:F2)`);

      this.applyBoldFont(sheet, "A1:G1");

      await workbook.toFileAsync(outputFileName);
      console.log('Archivo Excel creado y rellenado con datos.');
    } catch (error) {
      console.error('Error: ', error);
    }
  }

  writeHeader(sheet) {
    sheet.cell("A1").value("Ingresos");
    sheet.cell("C1").value("Gastos");
    sheet.cell("E1").value("Total Ingresos");
    sheet.cell("F1").value("Total Gastos");
    sheet.cell("G1").value("Resto");
  }

  applyBoldFont(sheet, range) {
    sheet.range(range).style({ bold: true });
  }
}

module.exports = ExcelWriter;