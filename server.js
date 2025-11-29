const express = require('express');
const cors = require('cors');
const PptxGenJS = require('pptxgenjs');

const app = express();
app.use(cors());
app.use(express.json({ limit: '100mb' }));

const PORT = 3000;

app.post('/generate-pptx', async (req, res) => {
  try {
    console.log('Generando PPTX...');
    const { slides, fileName } = req.body;
    
    if (!slides || slides.length === 0) {
      return res.status(400).json({ error: 'No hay slides para exportar' });
    }
    
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'A4_LANDSCAPE', width: 10, height: 7.5 });
    pptx.layout = 'A4_LANDSCAPE';
    
    for (let i = 0; i < slides.length; i++) {
      const slideData = slides[i];
      const slide = pptx.addSlide();
      slide.background = { color: 'FFFFFF' };
      
      if (slideData.imageBase64) {
        slide.addImage({
          data: 'data:image/png;base64,' + slideData.imageBase64,
          x: 0,
          y: 0,
          w: '100%',
          h: '100%',
          sizing: { type: 'contain', w: 10, h: 7.5 }
        });
      }
      
      console.log('Slide agregado:', slideData.title);
    }
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    console.log('PPTX generado:', buffer.length, 'bytes');
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName || 'presentation'}.pptx"`);
    res.send(buffer);
    
    console.log('PPTX enviado al cliente');
    
  } catch (error) {
    console.error('Error generando PPTX:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/', (req, res) => {
  res.send('Backend funcionando');
});

app.listen(PORT, () => {
  console.log(`✓ Servidor escuchando en http://localhost:${PORT}`);
  console.log('Esperando solicitudes de generación de PPTX...');
});
