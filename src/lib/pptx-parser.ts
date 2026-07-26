import JSZip from 'jszip';

export interface ExtractedPptxTemplate {
  rawText: string;
  placeholders: string[];
  slideCount: number;
}

/**
 * Parses uploaded .pptx file and extracts all <<Placeholder>> text boxes found across shapes
 */
export async function parsePptxTemplate(file: File): Promise<ExtractedPptxTemplate> {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    let combinedText = '';
    const slideFiles = Object.keys(zip.files).filter(path => path.startsWith('ppt/slides/slide') && path.endsWith('.xml'));

    for (const slidePath of slideFiles) {
      const xmlText = await zip.files[slidePath].async('string');
      // Extract text content inside <a:t> tags
      const textMatches = xmlText.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
      textMatches.forEach(t => {
        const clean = t.replace(/<[^>]+>/g, '');
        combinedText += clean + ' ';
      });
    }

    // Extract << Placeholder >> patterns
    const placeholdersSet = new Set<string>();
    const matches = combinedText.match(/<<\s*([^>]+?)\s*>>/g) || [];
    matches.forEach(m => {
      const inner = m.replace(/^<<\s*/, '').replace(/\s*>>$/, '').trim();
      if (inner) placeholdersSet.add(inner);
    });

    return {
      rawText: combinedText,
      placeholders: Array.from(placeholdersSet),
      slideCount: slideFiles.length
    };
  } catch (err) {
    console.error('Failed to parse PPTX template:', err);
    return {
      rawText: '',
      placeholders: [],
      slideCount: 0
    };
  }
}
