import PizZip from 'pizzip';

export function extractFirstPhotoFromDocx(buffer: ArrayBuffer): { photoBytes?: Uint8Array; photoExt?: string } {
  try {
    const zip = new PizZip(buffer as any);
    const files = Object.keys(zip.files || {}).filter(p => /^word\/media\//i.test(p) && !zip.files[p].dir);
    if (!files.length) return {};

    const ordered = files.filter(f => /\.(png|jpe?g)$/i.test(f)).sort((a, b) => a.localeCompare(b));
    const pick = ordered[0] || files[0];

    const u8 = zip.file(pick)?.asUint8Array?.();
    if (!u8 || !u8.length) return {};

    const extMatch = pick.match(/\.(png|jpe?g)$/i);
    const photoExt = extMatch ? extMatch[1].toLowerCase() : 'png';
    return { photoBytes: u8, photoExt };
  } catch {
    return {};
  }
}

export async function replaceFirstBodyImageWithSourcePhoto(zip: any, photoBytes: Uint8Array, photoExt: string) {
  const docXmlFile = zip.file('word/document.xml');
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!docXmlFile || !relsFile) return;

  const docXml = docXmlFile.asText();
  const relsXml = relsFile.asText();

  const ridMatch = docXml.match(/r:embed="(rId\d+)"/);
  if (!ridMatch) return;

  const rid = ridMatch[1];
  const relRe = new RegExp(`<Relationship[^>]+Id="${rid}"[^>]+Target="([^"]+)"[^>]*/>`, 'i');
  const relMatch = relsXml.match(relRe);
  if (!relMatch) return;

  const target = relMatch[1]; // e.g. "media/image2.png"
  const targetPath = target.startsWith('word/') ? target : `word/${target.replace(/^\.?\//, '')}`;
  const oldExtMatch = targetPath.match(/\.(png|jpe?g)$/i);
  const oldExt = oldExtMatch ? oldExtMatch[1].toLowerCase() : 'png';

  let finalBytes = photoBytes;
  if (oldExt !== (photoExt || '').toLowerCase()) {
    const want = (oldExt === 'jpg' || oldExt === 'jpeg') ? 'image/jpeg' : 'image/png';
    finalBytes = await convertImageBytes(photoBytes, want);
  }

  zip.file(targetPath, finalBytes);
}

async function convertImageBytes(bytes: Uint8Array, mime: string): Promise<Uint8Array> {
  const blob = new Blob([bytes], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);

  try {
    const img = await loadImage(url);
    const canvas = document.createElement('canvas');
    canvas.width = img.width || 300;
    canvas.height = img.height || 300;

    const ctx = canvas.getContext('2d');
    if (!ctx) return bytes;

    ctx.drawImage(img, 0, 0);

    const outBlob: Blob = await new Promise((resolve) => {
      canvas.toBlob((b) => resolve(b || blob), mime, 0.92);
    });

    const ab = await outBlob.arrayBuffer();
    return new Uint8Array(ab);
  } finally {
    URL.revokeObjectURL(url);
  }
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = (e) => reject(e);
    img.src = url;
  });
}
