const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '20mb' }));

app.post('/generar-excel', async (req, res) => {
  try {
    const datos = req.body;
    const workbook = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, 'plantilla', '08_M_OATransv (1).xlsx');
    await workbook.xlsx.readFile(templatePath);
    const hoja = workbook.getWorksheet('OA-1');

    // --- DATOS GENERALES ---
    hoja.getCell('C3').value = datos.ruta;
    hoja.getCell('C4').value = datos.sector;
    hoja.getCell('O3').value = datos.deDm;
    hoja.getCell('P3').value = ''; // P3 vacío si es sólo una celda de muestra
    hoja.getCell('O4').value = datos.aDm;
    hoja.getCell('P4').value = '';

    const mapObra = (obra, filaBase, bocaEntrada, bocaSalida, obsCelda) => {
      hoja.getCell(`A${filaBase}`).value = obra.numero;
      hoja.getCell(`B${filaBase}`).value = obra.dm;
      hoja.getCell(`C${filaBase}`).value = obra.tipo;
      hoja.getCell(`H${filaBase}`).value = obra.dimensiones.diametro;
      hoja.getCell(`I${filaBase}`).value = obra.dimensiones.altura;
      hoja.getCell(`J${filaBase}`).value = obra.dimensiones.ancho;
      hoja.getCell(`K${filaBase}`).value = obra.esviaje;
      hoja.getCell(`L${filaBase}`).value = obra.largo.total;
      hoja.getCell(`M${filaBase}`).value = obra.largo.ejeAlaIzq;
      hoja.getCell(`N${filaBase}`).value = obra.largo.ejeAlaDer;
      hoja.getCell(`O${filaBase}`).value = obra.tipoObraArte === 'riego' ? 'X' : '';
      hoja.getCell(`P${filaBase}`).value = obra.tipoObraArte === 'drenaje' ? 'X' : '';
      hoja.getCell(`Q${filaBase}`).value = obra.estado.ducto;
      hoja.getCell(`R${filaBase}`).value = obra.estado.muros;
      hoja.getCell(`S${filaBase}`).value = obra.estado.guardaRuedas;
      hoja.getCell(`T${filaBase}`).value = obra.estado.alas;

      // Boca de entrada
      hoja.getCell('A' + (filaBase + 2)).value = bocaEntrada.foto ? 'FOTO TOMADA' : '';
      hoja.getCell('F' + (filaBase + 3)).value = bocaEntrada.muro.alturaClave;
      hoja.getCell('F' + (filaBase + 5)).value = bocaEntrada.muro.largo;
      hoja.getCell('F' + (filaBase + 7)).value = bocaEntrada.muro.espesor;
      hoja.getCell('F' + (filaBase + 10)).value = bocaEntrada.alas.altura1;
      hoja.getCell('F' + (filaBase + 12)).value = bocaEntrada.alas.altura2;
      hoja.getCell('F' + (filaBase + 14)).value = bocaEntrada.alas.espesor;
      hoja.getCell('F' + (filaBase + 16)).value = bocaEntrada.alas.largo;
      hoja.getCell('G' + (filaBase + 2)).value = bocaEntrada.embancamiento;
      hoja.getCell('G' + (filaBase + 10)).value = bocaEntrada.descripcion;

      // Boca de salida
      hoja.getCell('K' + (filaBase + 2)).value = bocaSalida.foto ? 'FOTO TOMADA' : '';
      hoja.getCell('Q' + (filaBase + 3)).value = bocaSalida.muro.alturaClave;
      hoja.getCell('Q' + (filaBase + 5)).value = bocaSalida.muro.largo;
      hoja.getCell('Q' + (filaBase + 7)).value = bocaSalida.muro.espesor;
      hoja.getCell('Q' + (filaBase + 10)).value = bocaSalida.alas.altura1;
      hoja.getCell('Q' + (filaBase + 12)).value = bocaSalida.alas.altura2;
      hoja.getCell('Q' + (filaBase + 14)).value = bocaSalida.alas.espesor;
      hoja.getCell('Q' + (filaBase + 16)).value = bocaSalida.alas.largo;
      hoja.getCell('R' + (filaBase + 2)).value = bocaSalida.embancamiento;
      hoja.getCell('R' + (filaBase + 10)).value = bocaSalida.descripcion;

      // Observación si corresponde
      if (obsCelda && obra.observacion) {
        hoja.getCell(obsCelda).value = obra.observacion;
      }
    };

    // --- OBRA 1 ---
    mapObra(datos.obras[0], 8, datos.obras[0].bocaEntrada, datos.obras[0].bocaSalida, 'C45');

    // --- OBRA 2 ---
    mapObra(datos.obras[1], 26, datos.obras[1].bocaEntrada, datos.obras[1].bocaSalida, 'C48');

    // Guardar y enviar
    const outputPath = path.join(__dirname, 'reporte_generado.xlsx');
    await workbook.xlsx.writeFile(outputPath);
    const fileBuffer = fs.readFileSync(outputPath);
    res.setHeader('Content-Disposition', 'attachment; filename=reporte.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(fileBuffer);
  } catch (error) {
    console.error('Error al generar Excel:', error);
    res.status(500).send('Error al generar Excel');
  }
});

app.listen(PORT, () => {
  console.log(`Servidor backend corriendo en http://localhost:${PORT}`);
});


