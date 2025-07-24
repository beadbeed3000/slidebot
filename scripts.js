// scripts.js - Complete and Simplified Slide Generator Logic

// 1. Extract `.docx` text and populate textarea
const fileInput = document.getElementById('uploadDocx');
fileInput.addEventListener('change', (event) => {
  const reader = new FileReader();
  reader.onload = () => {
    mammoth.extractRawText({ arrayBuffer: reader.result })
      .then((result) => {
        document.getElementById('outlineText').value = result.value;
      })
      .catch((err) => alert('Error reading DOCX: ' + err.message));
  };
  reader.readAsArrayBuffer(event.target.files[0]);
});

// 2. Handle Generate button
document.getElementById('generateBtn').addEventListener('click', generatePresentation);

// 3. Main function to generate the PPT
function generatePresentation() {
  const text = document.getElementById('outlineText').value.trim();
  if (!text) {
    return alert('Please paste text or upload a file.');
  }

  // Read custom inputs
  const bgColor    = document.getElementById('bgColor').value;
  const textColor  = document.getElementById('textColor').value;
  const bulletSize = +document.getElementById('bulletSize').value;
  const logoUrl    = document.getElementById('logoUrl').value;
  const fileName   = document.getElementById('fileName').value;
  const spinner    = document.getElementById('spinner');

  // Show spinner
  spinner.hidden = false;

  try {
    const pptx = new PptxGenJS();
    const lines = text.split(/\r?\n/);
    let slide = null;
    let title = '';
    let bullets = [];

    // Pattern: numeric, Roman, letter headings
    const headingPattern = /^\s*(?:\d+(?:\.\d*)?|[IVXLCDM]+|[A-Z]\.|#+)\s*(.+)/;

    // Add content to current slide
    const finalizeSlide = () => {
      if (!slide) return;
      slide.background = { color: bgColor };
      if (title) {
        slide.addText(title, { x: 0.5, y: 0.3, w: '90%', h: 1, fontSize: 28, bold: true, color: textColor, align: 'center' });
      }
      if (bullets.length) {
        const items = bullets.map(idx => bullets[idx] = parseBullet(bullets[idx]));
        slide.addText(bullets.map(parseBullet), { x: 0.75, y: 1.2, w: 8.5, h: 5.5, fontSize: bulletSize, color: textColor });
      }
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 6.7, w: '100%', h: 0.3, fill: { color: '71AC4A' } });
      slide.addImage({ data: logoUrl, x: 0.5, y: 6.5, w: 1.2, h: 0.6 });
    };

    // Process lines
    lines.forEach((line) => {
      const trimmed = line.trim();
      const match = trimmed.match(headingPattern);
      if (match) {
        finalizeSlide();
        slide = pptx.addSlide();
        title = match[1];
        bullets = [];
      } else if (trimmed) {
        bullets.push(trimmed);
      }
    });
    finalizeSlide();

    pptx.writeFile(fileName);
  } catch (err) {
    alert('Error: ' + err.message);
  } finally {
    spinner.hidden = true;
  }
}

// Parse bullets with indentation
function parseBullet(line) {
  const indent = (line.match(/^\s*/) || [''])[0].length;
  const level = Math.floor(indent / 2) + 1;
  const text = line.trim().replace(/^[-*]\s*/, '');
  return { text, options: { bullet: true, indentLevel: level } };
}
