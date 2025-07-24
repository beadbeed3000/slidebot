// scripts.js - Updated Slide Generator Logic

// 1. Extract `.docx` text
const fileInput = document.getElementById('uploadDocx');
fileInput.addEventListener('change', () => {
  const reader = new FileReader();
  reader.onload = () => {
    mammoth.extractRawText({ arrayBuffer: reader.result })
      .then(result => {
        document.getElementById('outlineText').value = result.value;
      })
      .catch(err => alert('Error reading DOCX: ' + err.message));
  };
  reader.readAsArrayBuffer(fileInput.files[0]);
});

// 2. Hook up Generate button
document.getElementById('generateBtn').addEventListener('click', generatePresentation);

// 3. Main generation function
function generatePresentation() {
  const rawText = document.getElementById('outlineText').value;
  if (!rawText.trim()) {
    alert('Please paste text or upload a file.');
    return;
  }

  // Customization inputs
  const bgColor    = document.getElementById('bgColor').value;
  const textColor  = document.getElementById('textColor').value;
  const bulletSize = parseInt(document.getElementById('bulletSize').value, 10);
  const logoUrl    = document.getElementById('logoUrl').value;
  const fileName   = document.getElementById('fileName').value;
  const spinner    = document.getElementById('spinner');

  spinner.hidden = false;
  try {
    const pptx = new PptxGenJS();
    const lines = rawText.split(/\r?\n/);
    let slide = null;
    let title = '';
    let bullets = [];

    // Prefix regex to detect heading lines (with dot)
    const headingPrefix = /^(?:\d+(?:\.\d*)?|[IVXLCDM]+|[A-Z])\.(?:\s*)(.*)$/;

    // Finalize and render current slide
    function finalizeSlide() {
      if (!slide) return;
      slide.background = { color: bgColor };
      if (title) {
        slide.addText(title, { x:0.5, y:0.3, w:'90%', h:1, fontSize:28, bold:true, color:textColor, align:'center' });
      }
      if (bullets.length) {
        const items = bullets.map(line => parseBullet(line));
        slide.addText(items, { x:0.75, y:1.2, w:8.5, h:5.5, fontSize:bulletSize, color:textColor });
      }
      slide.addShape(pptx.ShapeType.rect, { x:0, y:6.7, w:'100%', h:0.3, fill:{ color:'71AC4A' }});
      slide.addImage({ data:logoUrl, x:0.5, y:6.5, w:1.2, h:0.6 });
    }

    // Build slides: start on matching headingPrefix
    lines.forEach(textLine => {
      const match = textLine.match(headingPrefix);
      if (match) {
        finalizeSlide();
        slide = pptx.addSlide();
        title = match[1].trim();
        bullets = [];
      } else if (textLine.trim()) {
        bullets.push(textLine);
      }
    });
    finalizeSlide();
    pptx.writeFile(fileName);
  } catch (err) {
    alert('Error generating presentation: ' + err.message);
  } finally {
    spinner.hidden = true;
  }
}

// 4. Parse bullets, supports nesting via indentation
function parseBullet(line) {
  const indent = (line.match(/^\s*/) || [''])[0].length;
  const level  = Math.floor(indent / 2) + 1;
  const text   = line.replace(/^\s*[-*]?\s*/, '').trim();
  return { text, options:{ bullet:true, indentLevel:level }};
}
