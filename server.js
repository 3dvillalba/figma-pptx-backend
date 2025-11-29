const express = require('express');
const cors = require('cors');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '100mb' }));

const PORT = process.env.PORT || 3000;

app.post('/generate-pptx', async (req, res) => {
  try {
    console.log('Generando PPTX con orientación vertical...');
    const { slides, fileName } = req.body;
    
    if (!slides || slides.length === 0) {
      return res.status(400).json({ error: 'No hay slides para exportar' });
    }

    // Crear archivo temporal JSON con datos de slides
    const tempData = {
      slides: slides,
      fileName: fileName || 'presentation'
    };

    const tempDir = '/tmp';
    const tempFile = path.join(tempDir, `slide_data_${Date.now()}.json`);
    fs.writeFileSync(tempFile, JSON.stringify(tempData));

    // Ejecutar script Python que genera el PPTX
    const pythonScript = path.join(__dirname, 'generate_pptx.py');
    const outputFile = path.join(tempDir, `output_${Date.now()}.pptx`);

    try {
      execSync(`python3 ${pythonScript} ${tempFile} ${outputFile}`, { 
        encoding: 'utf-8',
        maxBuffer: 100 * 1024 * 1024,
        stdio: 'pipe'
      });

      if (!fs.existsSync(outputFile)) {
        throw new Error('Python script no generó el archivo PPTX');
      }

      // Leer archivo PPTX generado
      const pptxBuffer = fs.readFileSync(outputFile);

      // Enviar al cliente
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename="${fileName || 'presentation'}.pptx"`);
      res.send(pptxBuffer);

      // Limpiar archivos temporales
      try {
        fs.unlinkSync(tempFile);
        fs.unlinkSync(outputFile);
      } catch (e) {
        console.log('No se pudieron limpiar temp files:', e);
      }

      console.log('✓ PPTX vertical generado exitosamente');

    } catch (pythonError) {
      console.error('Error ejecutando Python:', pythonError.message);
      res.status(500).json({ error: 'Error generando PPTX: ' + pythonError.message });
    }

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/', (req, res) => {
  res.send('Backend funcionando - PPTX Vertical con Texto Editable');
});

app.listen(PORT, () => {
  console.log(`✓ Servidor en puerto ${PORT}`);
});
