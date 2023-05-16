const fs = require('fs');
const { google } = require('googleapis');
const Docxtemplater = require('docxtemplater');

async function authenticate() {
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json',
    scopes: ['https://www.googleapis.com/auth/presentations'],
  });
  const authClient = await auth.getClient();
  google.options({
    auth: authClient,
  });
}

function extractMathContentFromDocx(docxFile) {
  const content = fs.readFileSync(docxFile, 'binary');
  const doc = new Docxtemplater(content);
  doc.setData({});
  doc.render();

  const xmlContent = doc.getZip().generate({ type: 'string' });
  const paragraphs = xmlContent.match(/<w:p\b[^>]*>.*?<\/w:p>/gs) || [];
  const mathContent = paragraphs.filter((p) => p.startsWith('<w:pict'));

  return mathContent;
}

async function createSlidesWithMathContent(mathContent) {
  const presentation = await google.slides({ version: 'v1' }).presentations.create({
    requestBody: {
      title: 'Mathematical Slides',
    },
  });
  const presentationId = presentation.data.presentationId;

  for (const xml of mathContent) {
    await google.slides({ version: 'v1' }).presentations.pages.batchUpdate({
      presentationId: presentationId,
      requestBody: {
        requests: [
          {
            createSlide: {
              slideLayoutReference: {
                predefinedLayout: 'BLANK',
              },
            },
          },
          {
            createImage: {
              objectId: `Math_${Date.now()}`,
              url: `data:image/png;base64,${Buffer.from(xml).toString('base64')}`,
              elementProperties: {
                pageObjectId: presentation.data.slides[0].objectId,
              },
            },
          },
        ],
      },
    });
  }

  console.log(`Slides created: https://docs.google.com/presentation/d/${presentationId}`);
}

// Usage example
const docxFile = 'math_document.docx';
authenticate()
  .then(() => {
    const mathContent = extractMathContentFromDocx(docxFile);
    return createSlidesWithMathContent(mathContent);
  })
  .catch((err) => console.error(err));
