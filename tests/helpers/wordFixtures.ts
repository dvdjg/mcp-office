import path from 'path';
import os from 'os';
import fs from 'fs-extra';
import {
  AlignmentType,
  Document,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

export interface GeneratedWordFixture {
  tempDir: string;
  docxPath: string;
}

export async function createTempWordWorkspace(prefix = 'mcp-office-word-'): Promise<string> {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

export async function createComplexWordFixture(baseDir: string, fileName = 'complex-structure.docx'): Promise<GeneratedWordFixture> {
  await fs.ensureDir(baseDir);
  const docxPath = path.join(baseDir, fileName);

  const document = new Document({
    numbering: {
      config: [
        {
          reference: 'complex-list',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 260 },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 1440, hanging: 260 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        children: [
          new Paragraph({
            text: 'Quarterly Review',
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            children: [
              new TextRun('This summary highlights '),
              new TextRun({ text: 'critical', bold: true }),
              new TextRun(' trends and '),
              new TextRun({ text: 'follow-up actions', italics: true }),
              new TextRun(' for the leadership team.'),
            ],
          }),
          new Paragraph({
            text: 'Regional Insights',
            heading: HeadingLevel.HEADING_2,
          }),
          new Paragraph({
            text: 'North region momentum recovered after the product relaunch.',
          }),
          new Paragraph({
            text: 'South region requires pricing review.',
            numbering: { reference: 'complex-list', level: 0 },
          }),
          new Paragraph({
            text: 'Prepare discount sensitivity analysis.',
            numbering: { reference: 'complex-list', level: 1 },
          }),
          new Paragraph({
            text: 'Schedule stakeholder workshop.',
            numbering: { reference: 'complex-list', level: 1 },
          }),
          new Paragraph({
            text: 'Action Tracker',
            heading: HeadingLevel.HEADING_2,
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Region')] }),
                  new TableCell({ children: [new Paragraph('Revenue')] }),
                  new TableCell({ children: [new Paragraph('Status')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('North')] }),
                  new TableCell({ children: [new Paragraph('1.2M')] }),
                  new TableCell({ children: [new Paragraph('Recovered')] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('South')] }),
                  new TableCell({ children: [new Paragraph('0.8M')] }),
                  new TableCell({ children: [new Paragraph('At Risk')] }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(document);
  await fs.writeFile(docxPath, buffer);

  return { tempDir: baseDir, docxPath };
}
