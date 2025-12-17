import * as DOCX from 'docx';
import { ProfileData } from '../types';

export async function buildDataportDocx(data: ProfileData): Promise<Blob> {
  const heading = new DOCX.Paragraph({
    text: 'Datenblatt und Einschätzung zur Erbringung der Arbeitnehmerüberlassung',
    heading: DOCX.HeadingLevel.HEADING_1,
    spacing: { after: 200 }
  });

  const summaryTable = new DOCX.Table({
    width: { size: 100, type: DOCX.WidthType.PERCENTAGE },
    rows: [
      row2('Name', data.name || ''),
      row2('Rolle/Position', data.role || ''),
      row2('Team/Abteilung', data.team || ''),
      row2('E-Mail', data.email || '')
    ]
  });

  const doc = new DOCX.Document({ sections: [{ children: [heading, summaryTable] }] });
  return DOCX.Packer.toBlob(doc);

  function row2(label: string, value: string): DOCX.TableRow {
    return new DOCX.TableRow({
      children: [
        new DOCX.TableCell({ width: { size: 30, type: DOCX.WidthType.PERCENTAGE }, children: [paraBold(label)] }),
        new DOCX.TableCell({ width: { size: 70, type: DOCX.WidthType.PERCENTAGE }, children: [new DOCX.Paragraph({ text: value || 'Bitte anpassen!' })] })
      ]
    });
  }
  function paraBold(text: string) { return new DOCX.Paragraph({ children: [new DOCX.TextRun({ text, bold: true })] }); }
}
