export interface FontOption {
  value: string;
  label: string;
  category: 'Built-in' | 'Serif' | 'Sans-Serif' | 'Script' | 'Display';
  google?: boolean;
  weights?: string;
  preview?: string;
}

export const CERTIFICATE_FONTS: FontOption[] = [
  // Built-in PDF fonts
  { value: "Helvetica", label: "Helvetica", category: "Built-in" },
  { value: "Helvetica-Bold", label: "Helvetica Bold", category: "Built-in" },
  { value: "Times-Roman", label: "Times Roman", category: "Built-in", preview: "Times New Roman" },
  { value: "Times-Bold", label: "Times Bold", category: "Built-in", preview: "Times New Roman" },
  { value: "Courier", label: "Courier", category: "Built-in" },
  
  // Prominent CSEA & Display Fonts
  { value: "League Gothic", label: "League Gothic", category: "Display", google: true, weights: "400" },
  { value: "ITC Motter Corpus Semicondensed", label: "ITC Motter Corpus Semicondensed", category: "Display", preview: "'ITC Motter Corpus Semicondensed', 'Motter Corpus', 'Paytone One', 'Impact', sans-serif" },
  { value: "TT Hoves Bold", label: "TT Hoves Bold", category: "Sans-Serif", preview: "'TT Hoves Bold', 'TT Hoves', 'Montserrat', 'Inter', sans-serif" },
  { value: "TT Hoves", label: "TT Hoves", category: "Sans-Serif", preview: "'TT Hoves', 'Montserrat', 'Inter', sans-serif" },
  { value: "Paytone One", label: "Paytone One", category: "Display", google: true, weights: "400" },
  { value: "Righteous", label: "Righteous", category: "Display", google: true, weights: "400" },

  // Elegant Serif (Google)
  { value: "Playfair Display", label: "Playfair Display", category: "Serif", google: true, weights: "400;700" },
  { value: "Cinzel", label: "Cinzel", category: "Serif", google: true, weights: "400;700" },
  { value: "Cinzel Decorative", label: "Cinzel Decorative", category: "Serif", google: true, weights: "400;700" },
  { value: "Cormorant Garamond", label: "Cormorant Garamond", category: "Serif", google: true, weights: "400;700" },
  { value: "EB Garamond", label: "EB Garamond", category: "Serif", google: true, weights: "400;700" },
  { value: "Lora", label: "Lora", category: "Serif", google: true, weights: "400;700" },
  { value: "Merriweather", label: "Merriweather", category: "Serif", google: true, weights: "400;700" },
  { value: "Libre Baskerville", label: "Libre Baskerville", category: "Serif", google: true, weights: "400;700" },
  { value: "Crimson Text", label: "Crimson Text", category: "Serif", google: true, weights: "400;700" },
  { value: "Spectral", label: "Spectral", category: "Serif", google: true, weights: "400;700" },
  
  // Modern Sans-Serif (Google)
  { value: "Poppins", label: "Poppins", category: "Sans-Serif", google: true, weights: "400;500;600;700;800" },
  { value: "Montserrat", label: "Montserrat", category: "Sans-Serif", google: true, weights: "400;600;700;800" },
  { value: "Open Sans", label: "Open Sans", category: "Sans-Serif", google: true, weights: "400;600;700" },
  { value: "Roboto", label: "Roboto", category: "Sans-Serif", google: true, weights: "400;500;700" },
  { value: "Raleway", label: "Raleway", category: "Sans-Serif", google: true, weights: "400;600;700" },
  { value: "Nunito", label: "Nunito", category: "Sans-Serif", google: true, weights: "400;600;700" },
  { value: "Lato", label: "Lato", category: "Sans-Serif", google: true, weights: "400;700" },
  { value: "Inter", label: "Inter", category: "Sans-Serif", google: true, weights: "400;600;700;800" },
  { value: "Oswald", label: "Oswald", category: "Sans-Serif", google: true, weights: "400;600;700" },

  // Script / Calligraphy (Google)
  { value: "Dancing Script", label: "Dancing Script", category: "Script", google: true, weights: "400;700" },
  { value: "Great Vibes", label: "Great Vibes", category: "Script", google: true, weights: "400" },
  { value: "Pacifico", label: "Pacifico", category: "Display", google: true, weights: "400" },
  { value: "Sacramento", label: "Sacramento", category: "Script", google: true, weights: "400" },
  { value: "Alex Brush", label: "Alex Brush", category: "Script", google: true, weights: "400" },
  { value: "Allura", label: "Allura", category: "Script", google: true, weights: "400" },
  { value: "Pinyon Script", label: "Pinyon Script", category: "Script", google: true, weights: "400" },
];

export function getGoogleFontStylesheetUrl(fonts: FontOption[]): string {
  const googleFonts = fonts.filter(f => f.google);
  if (googleFonts.length === 0) return '';
  const fontParams = googleFonts.map(f => {
    const encoded = f.value.replace(/ /g, '+');
    const weights = f.weights ? `:wght@${f.weights}` : ':wght@400;700';
    return `family=${encoded}${weights}`;
  });
  return `https://fonts.googleapis.com/css2?${fontParams.join('&')}&display=swap`;
}
