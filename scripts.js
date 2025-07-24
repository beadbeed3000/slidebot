// Extract DOCX text
document.getElementById('uploadDocx').addEventListener('change', event => {
  const reader = new FileReader();
  reader.onload = () => {
    mammoth.extractRawText({ arrayBuffer: reader.result })
      .then(result => document.getElementById('outlineText').value = result.value)
      .catch(err => alert('Error reading DOCX: ' + err.message));
  };
  reader.readAsArrayBuffer(event.target.files[0]);
});

// Generate presentation on button click
document.getElementById('generateBtn').addEventListener('click', generatePresentation);

function generatePresentation() {
  const text = document.getElementById('outlineText').value;
  if (!text.trim()) {
    alert('Please paste text or upload a file.');
    return;
  }

  // Read customization values
  const bgColor   = document.getElementById('bgColor').value;
  const textColor = document.getElementById('textColor').value;
  const bulletSize = +document.getElementById('bulletSize').value;
  const logoUrl   = document.getElementById('logoUrl').value;
  const fileName  = document.getElementById('fileName').value;
  const spinner   = document.getElementById('spinner');

  // Show spinner
  spinner.hidden = false;

  try {
    const pptx = new PptxGenJS();
    const lines = text.split(/\n+/);
    let slide = null, title = '', bullets = [];

    // Regex for headings: numeric (1.), Roman (I.), Markdown (#)
    const headingPattern = /^(?:\s*)(?:\d+(?:\.\d*)?|[IVXLCDM]+|#+)\.?\s+(.+)/;

    // Add slide content when ready
    function finalizeSlide() {
      if (!slide) return;
      slide.background = { color: bgColor };
      if (title) {
        slide.addText(title, {
          x: 0.5, y: 0.3, w: '90%', h: 1,
          fontSize: 28, bold: true, color: textColor,
          align: 'center'
        });
      }
      if (bullets.length) {
        const items = bullets.map(parseBullet);
        slide.addText(items, {
          x: 0.75, y: 1.2, w: 8.5, h: 5.5,
          fontSize: bulletSize, color: textColor
        });
      }
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 6.7, w: '100%', h: 0.3,
        fill: { color: '71AC4A' }
      });
      slide.addImage({ data: logoUrl, x: 9.2, y: 6.5, w: 1.2, h: 0.6 });
    }

    // Process each line
    lines.forEach(line => {
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

    // Save file
    pptx.writeFile(fileName);
  } catch (err) {
    alert('Error generating presentation: ' + err.message);
  } finally {
    spinner.hidden = true;
  }
}

// Parse nested bullets by indentation or markers
function parseBullet(line) {
  const indent = line.match(/^\s*/)[0].length;
  const level  = Math.floor(indent / 2) + 1;
  const text   = line.trim().replace(/^[-*]\s*/, '');
  return { text, options: { bullet: true, indentLevel: level } };
}
