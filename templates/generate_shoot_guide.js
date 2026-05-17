#!/usr/bin/env node

/**
 * Shoot Guide DOCX Generator — LOCKED FORMAT
 *
 * Renders the shoot guide JSON into a DOCX matching the approved gold standard:
 *   - ON-CAMERA LINES section (2 heroes, ~15 lines total)
 *   - B-ROLL SHOTS section (~15 clips)
 *   - REMIX VOICEOVERS section (7 full-paragraph scripts)
 *
 * Supports both new format (onCameraLines/brollShots/remixVoiceovers)
 * and legacy format (masterShotList) for backward compatibility.
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
} = require('docx');

const args = process.argv.slice(2);
if (args.length !== 2) {
  console.error('Usage: node generate_shoot_guide.js input.json output.docx');
  process.exit(1);
}

const inputPath = args[0];
const outputPath = args[1];

const jsonData = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

// Normalize data — support both new and legacy JSON formats
const heroes = jsonData.onCameraLines
  ? jsonData.onCameraLines.heroes || []
  : (jsonData.masterShotList || {}).talkingHeadShots || [];

const brollShots = jsonData.brollShots
  || (jsonData.masterShotList || {}).brollShots || [];

const remixVOs = jsonData.remixVoiceovers
  || (jsonData.masterShotList || {}).voiceoverAudio || [];

const totalVideos = jsonData.totalVideos || 12;

// Count total lines across heroes
const totalLines = heroes.reduce((sum, h) => sum + (h.lines || []).length, 0);

const children = [];

// ========== TITLE ==========
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 800, after: 0 },
    children: [
      new TextRun({
        text: 'SHOOT GUIDE',
        font: 'Arial',
        size: 44,
        bold: true,
        color: '1A1A2E',
      }),
    ],
  })
);

// Product name
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 0 },
    children: [
      new TextRun({
        text: jsonData.productName || 'Product',
        font: 'Arial',
        size: 28,
        bold: true,
        color: '2563EB',
      }),
    ],
  })
);

// Subtitle line
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 400 },
    border: {
      bottom: {
        color: 'D1D5DB',
        space: 8,
        style: BorderStyle.SINGLE,
        size: 2,
      },
    },
    children: [
      new TextRun({
        text: `${totalVideos} videos from one shoot \u2014 here\u2019s everything you need to capture`,
        font: 'Arial',
        size: 18,
        italics: true,
        color: '6B7280',
      }),
    ],
  })
);

// ========== ON-CAMERA LINES SECTION ==========
children.push(
  new Paragraph({
    spacing: { before: 300, after: 100 },
    border: {
      bottom: {
        color: '2563EB',
        space: 4,
        style: BorderStyle.SINGLE,
        size: 4,
      },
    },
    children: [
      new TextRun({
        text: `ON-CAMERA LINES (${heroes.length} videos, ${totalLines} lines)`,
        font: 'Arial',
        size: 24,
        bold: true,
        color: '2563EB',
      }),
    ],
  })
);

// Instruction
children.push(
  new Paragraph({
    spacing: { before: 40, after: 100 },
    children: [
      new TextRun({
        text: 'You can film each line, or couple of lines, as its own take \u2014 ',
        font: 'Arial',
        size: 18,
        italics: true,
        color: '6B7280',
      }),
      new TextRun({
        text: 'don\u2019t try to do it all in one shot.',
        font: 'Arial',
        size: 18,
        italics: true,
        bold: true,
        color: '6B7280',
      }),
      new TextRun({
        text: ' The editor will stitch them together.',
        font: 'Arial',
        size: 18,
        italics: true,
        color: '6B7280',
      }),
    ],
  })
);

// Each hero
heroes.forEach((hero) => {
  // Hero title — e.g., "Hero 1 — BFF Questions Discovery"
  const title = hero.title || hero.concept || `Hero ${heroes.indexOf(hero) + 1}`;
  children.push(
    new Paragraph({
      spacing: { before: 200, after: 60 },
      border: {
        bottom: {
          color: 'D1D5DB',
          space: 4,
          style: BorderStyle.SINGLE,
          size: 2,
        },
      },
      children: [
        new TextRun({
          text: title,
          font: 'Arial',
          size: 22,
          bold: true,
          color: '1A1A2E',
        }),
      ],
    })
  );

  // Each script line as blockquote-style
  (hero.lines || []).forEach((line) => {
    const typeColor = line.type === 'ON CAMERA' ? '2563EB' : '7C3AED';
    children.push(
      new Paragraph({
        spacing: { before: 40, after: 40 },
        indent: { left: 400 },
        children: [
          new TextRun({
            text: `${line.type} `,
            font: 'Arial',
            size: 18,
            bold: true,
            color: typeColor,
          }),
          new TextRun({
            text: line.text,
            font: 'Arial',
            size: 20,
            color: '333333',
          }),
        ],
      })
    );
  });
});

// ========== B-ROLL SECTION ==========
children.push(
  new Paragraph({
    spacing: { before: 300, after: 100 },
    border: {
      bottom: {
        color: '059669',
        space: 4,
        style: BorderStyle.SINGLE,
        size: 4,
      },
    },
    children: [
      new TextRun({
        text: `B-ROLL SHOTS (${brollShots.length} clips)`,
        font: 'Arial',
        size: 24,
        bold: true,
        color: '059669',
      }),
    ],
  })
);

brollShots.forEach((shot) => {
  const mainChildren = [
    new TextRun({
      text: shot.label,
      font: 'Arial',
      size: 21,
      bold: true,
      color: '059669',
    }),
    new TextRun({
      text: `  ${shot.description}`,
      font: 'Arial',
      size: 20,
      color: '333333',
    }),
  ];

  children.push(
    new Paragraph({
      spacing: { before: 40, after: 40 },
      indent: { left: 200 },
      children: mainChildren,
    })
  );

  // If shot has extra notes (e.g., specific page references), render them indented
  if (shot.notes) {
    children.push(
      new Paragraph({
        spacing: { before: 10, after: 40 },
        indent: { left: 400 },
        children: [
          new TextRun({
            text: shot.notes,
            font: 'Arial',
            size: 18,
            italics: true,
            color: '6B7280',
          }),
        ],
      })
    );
  }
});

// ========== REMIX VOICEOVERS SECTION ==========
children.push(
  new Paragraph({
    spacing: { before: 300, after: 100 },
    border: {
      bottom: {
        color: '7C3AED',
        space: 4,
        style: BorderStyle.SINGLE,
        size: 4,
      },
    },
    children: [
      new TextRun({
        text: `REMIX VOICEOVERS (${remixVOs.length} clips)`,
        font: 'Arial',
        size: 24,
        bold: true,
        color: '7C3AED',
      }),
    ],
  })
);

// Instruction
children.push(
  new Paragraph({
    spacing: { before: 40, after: 100 },
    children: [
      new TextRun({
        text: 'These are audio-only voiceovers for the remix videos. Record in a quiet space \u2014 natural voice, don\u2019t perform it.',
        font: 'Arial',
        size: 18,
        italics: true,
        color: '6B7280',
      }),
    ],
  })
);

remixVOs.forEach((vo) => {
  // Title line — "Remix 1 — Thoughtful Gift Solution"
  const voTitle = vo.title || `${vo.label || 'Remix'} \u2014 ${vo.concept || ''}`;
  children.push(
    new Paragraph({
      spacing: { before: 140, after: 40 },
      children: [
        new TextRun({
          text: voTitle,
          font: 'Arial',
          size: 22,
          bold: true,
          color: '7C3AED',
        }),
      ],
    })
  );

  // Full voiceover script in quotes, indented
  const scriptText = vo.script || '';
  children.push(
    new Paragraph({
      spacing: { before: 20, after: 60 },
      indent: { left: 400 },
      children: [
        new TextRun({
          text: `\u201C${scriptText}\u201D`,
          font: 'Arial',
          size: 20,
          color: '333333',
        }),
      ],
    })
  );
});

// Create the document
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
            style: {
              paragraph: {
                indent: { left: 720, hanging: 360 },
              },
            },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: {
            width: 12240,
            height: 15840,
          },
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440,
          },
        },
      },
      children: children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`\u2713 Shoot Guide created: ${outputPath}`);
});
