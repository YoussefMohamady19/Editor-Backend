// server.js
const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const { JSDOM } = require('jsdom');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const { Document, Packer, Paragraph, HeadingLevel } = require('docx');

const app = express();
app.use(cors()); // Allow frontend requests
app.use(express.json({ limit: '10mb' }));

// Multer setup - store temp files in ./uploads
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + path.extname(file.originalname));
  }
});
const upload = multer({ storage });

// --- Helper: build tree from HTML produced by mammoth
function buildTreeFromHtml(html) {
  const dom = new JSDOM(html);
  const doc = dom.window.document;
  const bodyChildren = Array.from(doc.body.children);

  const tree = [];
  const stack = [];

  function popToLevel(level) {
    while (stack.length > 0 && stack[stack.length - 1].level >= level) {
      stack.pop();
    }
  }

  function createNode(title, level) {
    return {
      id: cryptoId(),
      title: title || '',
      level: level,
      contentHtml: '',
      children: []
    };
  }

  function cryptoId() {
    // simple unique id
    return 'n_' + Math.random().toString(36).slice(2, 9);
  }

  for (const node of bodyChildren) {
    const tag = node.tagName.toUpperCase();

    const headingMatch = tag.match(/^H([1-6])$/);
    if (headingMatch) {
      const level = parseInt(headingMatch[1], 10);
      const title = node.textContent.trim();

      popToLevel(level);

      const newNode = createNode(title, level);
      if (stack.length === 0) {
        tree.push(newNode);
      } else {
        const parent = stack[stack.length - 1];
        parent.children.push(newNode);
      }
      stack.push(newNode);
    } else {
      // not a heading -> treat as content, append to current top of stack
      const htmlFragment = node.outerHTML;
      if (stack.length > 0) {
        const current = stack[stack.length - 1];
        current.contentHtml += htmlFragment;
      } else {
        // paragraphs before any heading -> create an intro node if none
        if (tree.length === 0) {
          const intro = createNode('Intro', 1);
          intro.contentHtml += htmlFragment;
          tree.push(intro);
          stack.push(intro);
        } else {
          // attach to first top-level node
          tree[0].contentHtml += htmlFragment;
        }
      }
    }
  }

  return tree;
}

// --- Endpoint: upload .docx and return JSON tree


app.post("/api/upload-file", upload.single("file"), async (req, res) => {
    try {
        const filePath = req.file.path;

        const result = await mammoth.convertToHtml(
            { path: filePath },
            {
                includeDefaultStyleMap: true,
                includeEmbeddedStyleMap: true,
                ignoreEmptyParagraphs: false,
                preserveLang: true,

                styleMap: [
                    "b => strong",
                    "i => em",
                    "u => u",
                    "strike => s",
                    "table => table.fresh",
                    "p[style-name='Normal'] => p:fresh"
                ],

                convertImage: mammoth.images.inline(async (element) => {
                    const buffer = await element.read("base64");
                    return {
                        src: "data:" + element.contentType + ";base64," + buffer,
                    };
                }),
            }
        );

        const html = result.value;

        fs.unlinkSync(filePath);

        // Build your tree like before
        const dom = new JSDOM(html);
        const body = dom.window.document.body;
        const tree = buildTreeFromHtml(body.innerHTML);

        res.json({ tree });

    } catch (err) {
        console.error("UPLOAD ERROR:", err);
        res.status(500).json({ error: "Failed to process Word file" });
    }
});


app.post("/api/upload-base64", async (req, res) => {
    try {
        const { base64 } = req.body;

        if (!base64)
            return res.status(400).json({ error: "No base64 provided" });

        // decode base64
        const fileBuffer = Buffer.from(base64.split(",")[1], "base64");

        const tempPath = path.join(uploadDir, 
            "base64-" + Date.now() + ".docx"
        );

        fs.writeFileSync(tempPath, fileBuffer);

        // Now read it with mammoth just like normal file
        const result = await mammoth.convertToHtml(
            { path: tempPath },
            {
                convertImage: mammoth.images.inline((element) =>
                    element.read("base64").then((imageBuffer) => ({
                        src: "data:" + element.contentType + ";base64," + imageBuffer,
                    }))
                ),
                includeDefaultStyleMap: true,
            }
        );

        const html = result.value;

        // Remove temp file
        fs.unlinkSync(tempPath);

        // Build tree
        const dom = new JSDOM(html);
        const body = dom.window.document.body;
        const tree = buildTreeFromHtml(body.innerHTML);

        res.json({ tree });

    } catch (err) {
        console.error("BASE64 UPLOAD ERROR:", err);
        res.status(500).json({ error: "Failed to parse base64 Word file" });
    }
});


// --- (Optional) Endpoint: export tree -> docx
// Note: simple naive export that writes headings and plain paragraphs (no rich formatting)
// If you want a full-featured export (preserve inline styles/images) we can add more logic.
// const { Document, Packer, Paragraph, HeadingLevel } = require('docx');

function buildDocxFromTree(tree) {
    const paragraphs = [];

    function append(nodes) {
        nodes.forEach(node => {
            paragraphs.push(
                new Paragraph({
                    text: node.title,
                    heading: HeadingLevel["HEADING_" + node.level]
                })
            );

            // تحويل الـ HTML إلى نص فقط (مؤقتاً)
            const plain = node.contentHtml
                .replace(/<br\s*\/?>/gi, "\n")
                .replace(/<\/p>/gi, "\n")
                .replace(/<[^>]+>/g, "");

            if (plain.trim().length > 0) {
                paragraphs.push(new Paragraph(plain));
            }

            if (node.children.length > 0) append(node.children);
        });
    }

    append(tree);

    const doc = new Document({
        sections: [{ children: paragraphs }]
    });

    return Packer.toBuffer(doc);
}

const htmlToDocx = require("html-to-docx");

function cleanHtml(html) {
    // Remove dangerous MS Word namespaces
    html = html.replace(/w:[a-zA-Z0-9\-]+="[^"]*"/gi, "");

    // Remove mso- styles
    html = html.replace(/mso-[^:;"]+:[^;"]+;?/gi, "");

    // Remove attributes starting with @
    html = html.replace(/@[a-zA-Z0-9\-]+="[^"]*"/gi, "");

    // Remove attributes starting with _
    html = html.replace(/_[a-zA-Z0-9\-]+="[^"]*"/gi, "");

    // Remove comments
    html = html.replace(/<!--[\s\S]*?-->/gi, "");

    // Cleanup table widths: width="@w", width="auto", etc.
    html = html.replace(/width="[^"]*"/gi, "");

    // Cleanup style attributes containing invalid values
    html = html.replace(/style="[^"]*?@[^"]*?"/gi, "");
    html = html.replace(/style="[^"]*?mso-[^"]*?"/gi, "");

    return html;
}

function sanitizeTable(html) {
    const dom = new JSDOM(html);
    const doc = dom.window.document;

    const elements = doc.querySelectorAll("table, tr, td, th, thead, tbody, tfoot");

    elements.forEach((el) => {
        // Remove all attributes completely
        while (el.attributes.length > 0) {
            el.removeAttribute(el.attributes[0].name);
        }
    });

    return doc.body.innerHTML;
}


// EXPORT UPDATED DOCUMENT TO DOCX
app.post("/api/export", async (req, res) => {
    try {
        const nodes = req.body;

        // Combine all sections into one big HTML document
        let finalHtml = "";

        const buildHtml = (list, level = 1) => {
            list.forEach((n) => {
                finalHtml += `<h${n.level}>${n.title}</h${n.level}>`;

                if (n.contentHtml) {
                    let safe = cleanHtml(n.contentHtml);
                    safe = sanitizeTable(safe);
                    finalHtml += safe;
                }

                if (n.children && n.children.length) {
                    buildHtml(n.children, level + 1);
                }
            });
        };

        buildHtml(nodes);

        const docxBuffer = await htmlToDocx(finalHtml, null, {
            table: { row: { cantSplit: true } },
            footer: true,
            pageNumber: true,
            font: "Calibri",
            getImage: async (url) => {
                // For Base64 images inserted from JoditEditor:
                if (url.startsWith("data:image")) {
                    const base64 = url.split(",")[1];
                    return Buffer.from(base64, "base64");
                }
                return null;
            }
        });

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.setHeader("Content-Disposition", "attachment; filename=document.docx");
        res.send(docxBuffer);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Failed to export Word" });
    }
});


const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`word-backend listening on ${PORT}`));
