import JSZip from 'jszip';
import { generatePdfFromElement } from './pdf-generator';

export async function createCertificatesZip(
  certificates: { filename: string; element: HTMLElement }[],
  onProgress?: (current: number, total: number) => void
): Promise<Blob> {
  const zip = new JSZip();
  const folder = zip.folder('CSEA_Certificates') || zip;

  for (let i = 0; i < certificates.length; i++) {
    const item = certificates[i];
    try {
      const blob = await generatePdfFromElement(item.element, item.filename);
      const arrayBuffer = await blob.arrayBuffer();
      folder.file(item.filename, arrayBuffer);
    } catch (err) {
      console.error(`Failed to add ${item.filename} to ZIP:`, err);
    }

    if (onProgress) {
      onProgress(i + 1, certificates.length);
    }
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  return zipBlob;
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
