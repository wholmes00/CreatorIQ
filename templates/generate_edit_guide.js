#!/usr/bin/env node

/**
 * Edit Guide DOCX Generator — LOCKED FORMAT
 *
 * Renders the edit guide JSON into a DOCX matching the approved gold standard:
 *   - Title page with product name, video count, analysis basis, date
 *   - SECTION A: Hero Videos (2 heroes)
 *     Each: Hook, Duration + Audio, SCRIPT, B-ROLL USED, ON-SCREEN TEXT
 *   - SECTION B: Remix Videos (10 remixes)
 *     Each: Hook, Duration + Audio, SCRIPT (optional), B-ROLL USED, ON-SCREEN TEXT
 *
 * NO CTA strategy. NO visual timeline. NO upload details.
 */

const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  BorderStyle,
  LevelFormat,
  PageBreak,
} = require('docx');

const args = process.argv.slice(2);
if (args.length !== 2) {
  console.error('Usage: node generate_edit_guide.js input.json output.docx');
  process.exit(1);
}

const inputPath = args[0];
const outputPath = args[1];

const jsonData = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const COLORS = {
  primary: '2563EB',
  dark: '1A1A2E',
  secondary: '6B7280',
  green: '059669',
  red: 'DC2626',
  purple: '7C3AED',
  text: '333333',
};

const children = [];

// ========== TITLE PAGE ==========

// Product name large
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 2000, after: 0 },
    children: [
      new TextRun({
        text: (jsonData.productName || 'Product').toUpperCase(),
        font: 'Arial',
        size: 52,
        bold: true,
        color: COLORS.dark,
      }),
    ],
  })
);

// "TikTok Affiliate Edit Guide"
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 0 },
    border: {
      top: {
        color: COLORS.primary,
        space: 8,
        style: BorderStyle.SINGLE,
        size: 4,
      },
    },
    children: [
      new TextRun({
        text: 'TikTok Affiliate Edit Guide',
        font: 'Arial',
        size: 24,
        bold: true,
        color: COLORS.primary,
      }),
    ],
  })
);

// Video count line
const heroCount = (jsonData.heroVideos || []).length;
const remixCount = (jsonData.remixVideos || []).length;
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 0 },
    children: [
      new TextRun({
        text: `${heroCount} Hero Videos + ${remixCount} Remix Videos`,
        font: 'Arial',
        size: 20,
        color: COLORS.secondary,
      }),
    ],
  })
);

// Analysis basis
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 600, after: 0 },
    children: [
      new TextRun({
        text: 'Based on analysis of top-performing TikTok affiliate videos',
        font: 'Arial',
        size: 18,
        italics: true,
        color: COLORS.secondary,
      }),
    ],
  })
);

// Date
const now = new Date();
const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 0 },
    children: [
      new TextRun({
        text: `${monthNames[now.getMonth()]} ${now.getFullYear()}`,
        font: 'Arial',
        size: 18,
        color: COLORS.secondary,
      }),
    ],
  })
);

// Page break
children.push(new Paragraph({ children: [new PageBreak()] }));

// ========== SECTION A: HERO VIDEOS ==========

// Section header
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 40 },
    children: [
      new TextRun({
        text: 'SECTION A',
        font: 'Arial',
        size: 18,
        bold: true,
        color: COLORS.primary,
      }),
    ],
  })
);

children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [
      new TextRun({
        text: 'Hero Videos',
        font: 'Arial',
        size: 34,
        bold: true,
        color: COLORS.dark,
      }),
    ],
  })
);

children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [
      new TextRun({
        text: 'On-camera storytelling \u2014 creator speaks directly to the audience',
        font: 'Arial',
        size: 20,
        italics: true,
        color: COLORS.secondary,
      }),
    ],
  })
);

// Process each hero video
if (jsonData.heroVideos && jsonData.heroVideos.length > 0) {
  jsonData.heroVideos.forEach((hero, idx) => {
    const num = idx + 1;

    // "HERO VIDEO N"
    children.push(
      new Paragraph({
        spacing: { before: 100, after: 60 },
        children: [
          new TextRun({
            text: `HERO VIDEO ${num}`,
            font: 'Arial',
            size: 18,
            bold: true,
            color: COLORS.primary,
          }),
        ],
      })
    );

    // "Hero N — Concept"
    children.push(
      new Paragraph({
        spacing: { before: 0, after: 100 },
        children: [
          new TextRun({
            text: hero.title || `Hero ${num}`,
            font: 'Arial',
            size: 30,
            bold: true,
            color: COLORS.dark,
          }),
        ],
      })
    );

    // Hook line
    const hookText = typeof hero.hook === 'string' ? hero.hook : (hero.hook && hero.hook.spoken) || '';
    if (hookText) {
      children.push(
        new Paragraph({
          spacing: { before: 60, after: 40 },
          children: [
            new TextRun({ text: 'Hook: ', font: 'Arial', size: 20, bold: true, color: COLORS.dark }),
            new TextRun({ text: `\u201C${hookText}\u201D`, font: 'Arial', size: 20, color: COLORS.text }),
          ],
        })
      );
    }

    // Duration + Audio on same line
    if (hero.duration || hero.audio) {
      const durationAudioParts = [];
      if (hero.duration) durationAudioParts.push(new TextRun({ text: 'Duration: ', font: 'Arial', size: 20, bold: true, color: COLORS.dark }));
      if (hero.duration) durationAudioParts.push(new TextRun({ text: hero.duration, font: 'Arial', size: 20, color: COLORS.text }));
      if (hero.duration && hero.audio) durationAudioParts.push(new TextRun({ text: '  ', font: 'Arial', size: 20 }));
      if (hero.audio) durationAudioParts.push(new TextRun({ text: 'Audio: ', font: 'Arial', size: 20, bold: true, color: COLORS.dark }));
      if (hero.audio) durationAudioParts.push(new TextRun({ text: hero.audio, font: 'Arial', size: 20, color: COLORS.secondary }));

      children.push(
        new Paragraph({
          spacing: { before: 20, after: 80 },
          children: durationAudioParts,
        })
      );
    }

    // ── SCRIPT ──
    const scriptLines = hero.script || [];
    if (scriptLines.length > 0) {
      children.push(
        new Paragraph({
          spacing: { before: 120, after: 60 },
          border: {
            bottom: { color: COLORS.primary, space: 4, style: BorderStyle.SINGLE, size: 6 },
          },
          children: [
            new TextRun({ text: 'SCRIPT', font: 'Arial', size: 22, bold: true, color: COLORS.primary }),
          ],
        })
      );

      scriptLines.forEach((line) => {
        const typeColor = line.type === 'ON CAMERA' ? COLORS.primary : COLORS.purple;
        children.push(
          new Paragraph({
            spacing: { before: 30, after: 30 },
            indent: { left: 400 },
            children: [
              new TextRun({ text: `[${line.type}] `, font: 'Arial', size: 18, bold: true, color: typeColor }),
              new TextRun({ text: line.text, font: 'Arial', size: 19, color: COLORS.text }),
            ],
          })
        );
      });
    }

    // ── B-ROLL USED ──
    const brollUsed = hero.brollUsed || [];
    if (brollUsed.length > 0) {
      children.push(
        new Paragraph({
          spacing: { before: 120, after: 60 },
          border: {
            bottom: { color: COLORS.green, space: 4, style: BorderStyle.SINGLE, size: 6 },
          },
          children: [
            new TextRun({ text: 'B-ROLL USED', font: 'Arial', size: 22, bold: true, color: COLORS.green }),
          ],
        })
      );

      brollUsed.forEach((item) => {
        if (typeof item === 'string') {
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: item, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        } else {
          // Object with time, label, description
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: `[${item.time}] `, font: 'Arial', size: 18, bold: true, color: COLORS.green }),
                new TextRun({ text: `${item.label} `, font: 'Arial', size: 19, bold: true, color: COLORS.green }),
                new TextRun({ text: item.description || '', font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        }
      });
    }

    // ── ON-SCREEN TEXT ──
    const onScreenText = hero.onScreenText || [];
    if (onScreenText.length > 0) {
      children.push(
        new Paragraph({
          spacing: { before: 120, after: 60 },
          border: {
            bottom: { color: COLORS.red, space: 4, style: BorderStyle.SINGLE, size: 6 },
          },
          children: [
            new TextRun({ text: 'ON-SCREEN TEXT', font: 'Arial', size: 22, bold: true, color: COLORS.red }),
          ],
        })
      );

      onScreenText.forEach((item) => {
        if (typeof item === 'string') {
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: item, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        } else {
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: `[${item.time}] `, font: 'Arial', size: 18, bold: true, color: COLORS.red }),
                new TextRun({ text: item.text, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        }
      });
    }

    // Page break between hero videos (except last)
    if (idx < jsonData.heroVideos.length - 1) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  });
}

// Page break before Section B
children.push(new Paragraph({ children: [new PageBreak()] }));

// ========== SECTION B: REMIX VIDEOS ==========

children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 40 },
    children: [
      new TextRun({ text: 'SECTION B', font: 'Arial', size: 18, bold: true, color: COLORS.primary }),
    ],
  })
);

children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [
      new TextRun({ text: 'Remix Videos', font: 'Arial', size: 34, bold: true, color: COLORS.dark }),
    ],
  })
);

children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [
      new TextRun({
        text: 'Short, music-driven edits \u2014 b-roll with voiceover or text only',
        font: 'Arial',
        size: 20,
        italics: true,
        color: COLORS.secondary,
      }),
    ],
  })
);

// Process each remix video
if (jsonData.remixVideos && jsonData.remixVideos.length > 0) {
  jsonData.remixVideos.forEach((remix, idx) => {
    const num = idx + 1;

    // "REMIX VIDEO N"
    children.push(
      new Paragraph({
        spacing: { before: 100, after: 40 },
        children: [
          new TextRun({
            text: `REMIX VIDEO ${num}`,
            font: 'Arial',
            size: 18,
            bold: true,
            color: COLORS.primary,
          }),
        ],
      })
    );

    // "Remix N — Title"
    children.push(
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [
          new TextRun({
            text: remix.title || `Remix ${num}`,
            font: 'Arial',
            size: 30,
            bold: true,
            color: COLORS.dark,
          }),
        ],
      })
    );

    // Hook
    const hookText = typeof remix.hook === 'string' ? remix.hook : '';
    if (hookText) {
      children.push(
        new Paragraph({
          spacing: { before: 40, after: 40 },
          children: [
            new TextRun({ text: 'Hook: ', font: 'Arial', size: 20, bold: true, color: COLORS.dark }),
            new TextRun({ text: `\u201C${hookText}\u201D`, font: 'Arial', size: 20, color: COLORS.text }),
          ],
        })
      );
    }

    // Duration + Audio
    if (remix.duration || remix.audio) {
      const parts = [];
      if (remix.duration) {
        parts.push(new TextRun({ text: 'Duration: ', font: 'Arial', size: 20, bold: true, color: COLORS.dark }));
        parts.push(new TextRun({ text: remix.duration, font: 'Arial', size: 20, color: COLORS.text }));
      }
      if (remix.duration && remix.audio) parts.push(new TextRun({ text: '  ', font: 'Arial', size: 20 }));
      if (remix.audio) {
        parts.push(new TextRun({ text: 'Audio: ', font: 'Arial', size: 20, bold: true, color: COLORS.dark }));
        parts.push(new TextRun({ text: remix.audio, font: 'Arial', size: 20, color: COLORS.secondary }));
      }
      children.push(new Paragraph({ spacing: { before: 20, after: 60 }, children: parts }));
    }

    // ── SCRIPT (optional — text-only remixes won't have this) ──
    const remixScript = remix.script || [];
    if (remixScript.length > 0) {
      children.push(
        new Paragraph({
          spacing: { before: 100, after: 60 },
          border: {
            bottom: { color: COLORS.purple, space: 4, style: BorderStyle.SINGLE, size: 6 },
          },
          children: [
            new TextRun({ text: 'SCRIPT', font: 'Arial', size: 22, bold: true, color: COLORS.purple }),
          ],
        })
      );

      remixScript.forEach((line) => {
        if (typeof line === 'string') {
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: `[VOICEOVER] `, font: 'Arial', size: 18, bold: true, color: COLORS.purple }),
                new TextRun({ text: line, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        } else {
          const typeColor = (line.type || 'VOICEOVER') === 'ON CAMERA' ? COLORS.primary : COLORS.purple;
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: `[${line.type || 'VOICEOVER'}] `, font: 'Arial', size: 18, bold: true, color: typeColor }),
                new TextRun({ text: line.text, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        }
      });
    }

    // ── B-ROLL USED ──
    const brollUsed = remix.brollUsed || [];
    if (brollUsed.length > 0) {
      children.push(
        new Paragraph({
          spacing: { before: 100, after: 60 },
          border: {
            bottom: { color: COLORS.green, space: 4, style: BorderStyle.SINGLE, size: 6 },
          },
          children: [
            new TextRun({ text: 'B-ROLL USED', font: 'Arial', size: 22, bold: true, color: COLORS.green }),
          ],
        })
      );

      brollUsed.forEach((item) => {
        const text = typeof item === 'string' ? item : `${item.label} ${item.description || ''}`;
        children.push(
          new Paragraph({
            spacing: { before: 20, after: 20 },
            indent: { left: 400 },
            children: [
              new TextRun({ text: text, font: 'Arial', size: 19, color: COLORS.text }),
            ],
          })
        );
      });
    }

    // ── ON-SCREEN TEXT ──
    const onScreenText = remix.onScreenText || [];
    if (onScreenText.length > 0) {
      children.push(
        new Paragraph({
          spacing: { before: 100, after: 60 },
          border: {
            bottom: { color: COLORS.red, space: 4, style: BorderStyle.SINGLE, size: 6 },
          },
          children: [
            new TextRun({ text: 'ON-SCREEN TEXT', font: 'Arial', size: 22, bold: true, color: COLORS.red }),
          ],
        })
      );

      onScreenText.forEach((item) => {
        if (typeof item === 'string') {
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: item, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        } else {
          children.push(
            new Paragraph({
              spacing: { before: 20, after: 20 },
              indent: { left: 400 },
              children: [
                new TextRun({ text: `[${item.time}] `, font: 'Arial', size: 18, bold: true, color: COLORS.red }),
                new TextRun({ text: item.text, font: 'Arial', size: 19, color: COLORS.text }),
              ],
            })
          );
        }
      });
    }

    // Page break between remix videos (except last)
    if (idx < jsonData.remixVideos.length - 1) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  });
}

// Create document
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: '\u2022',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`\u2713 Edit Guide created: ${outputPath}`);
});
