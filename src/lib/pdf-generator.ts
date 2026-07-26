import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

export async function generatePdfFromElement(
  element: HTMLElement,
  filename: string = 'certificate.pdf'
): Promise<Blob> {
  const canvas = await html2canvas(element, {
    scale: 3, // High DPI for crisp vector text
    useCORS: true,
    allowTaint: true,
    backgroundColor: '#ffffff',
    logging: false
  });

  const imgData = canvas.toDataURL('image/png', 1.0);
  
  // Landscape A4 dimensions in mm: 297 x 210
  const pdf = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });

  pdf.addImage(imgData, 'PNG', 0, 0, 297, 210);
  const blob = pdf.output('blob');
  return blob;
}

export async function downloadPdfFromElement(
  element: HTMLElement,
  filename: string = 'certificate.pdf'
): Promise<void> {
  const blob = await generatePdfFromElement(element, filename);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
