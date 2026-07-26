export interface FontOption {
  value: string;
  label: string;
  category: 'Built-in' | 'Serif' | 'Sans-Serif' | 'Script' | 'Display';
  google?: boolean;
  preview?: string;
}

export const CERTIFICATE_FONTS: FontOption[] = [
  // Built-in PDF fonts
  { value: "Helvetica", label: "Helvetica", category: "Built-in" },
  { value: "Helvetica-Bold", label: "Helvetica Bold", category: "Built-in" },
  { value: "Times-Roman", label: "Times Roman", category: "Built-in", preview: "Times New Roman" },
  { value: "Times-Bold", label: "Times Bold", category: "Built-in", preview: "Times New Roman" },
  { value: "Courier", label: "Courier", category: "Built-in" },
  
  // Elegant Serif (Google)
  { value: "Playfair Display", label: "Playfair Display", category: "Serif", google: true },
  { value: "Cinzel", label: "Cinzel", category: "Serif", google: true },
  { value: "Cinzel Decorative", label: "Cinzel Decorative", category: "Serif", google: true },
  { value: "Cormorant Garamond", label: "Cormorant Garamond", category: "Serif", google: true },
  { value: "EB Garamond", label: "EB Garamond", category: "Serif", google: true },
  { value: "Lora", label: "Lora", category: "Serif", google: true },
  { value: "Merriweather", label: "Merriweather", category: "Serif", google: true },
  { value: "Libre Baskerville", label: "Libre Baskerville", category: "Serif", google: true },
  { value: "Crimson Text", label: "Crimson Text", category: "Serif", google: true },
  { value: "Spectral", label: "Spectral", category: "Serif", google: true },
  
  // Modern Sans-Serif (Google)
  { value: "Montserrat", label: "Montserrat", category: "Sans-Serif", google: true },
  { value: "Open Sans", label: "Open Sans", category: "Sans-Serif", google: true },
  { value: "Roboto", label: "Roboto", category: "Sans-Serif", google: true },
  { value: "Poppins", label: "Poppins", category: "Sans-Serif", google: true },
  { value: "Raleway", label: "Raleway", category: "Sans-Serif", google: true },
  { value: "Nunito", label: "Nunito", category: "Sans-Serif", google: true },
  { value: "Lato", label: "Lato", category: "Sans-Serif", google: true },
  { value: "Inter", label: "Inter", category: "Sans-Serif", google: true },
  { value: "Oswald", label: "Oswald", category: "Sans-Serif", google: true },

  // Script / Calligraphy (Google)
  { value: "Dancing Script", label: "Dancing Script", category: "Script", google: true },
  { value: "Great Vibes", label: "Great Vibes", category: "Script", google: true },
  { value: "Pacifico", label: "Pacifico", category: "Display", google: true },
  { value: "Sacramento", label: "Sacramento", category: "Script", google: true },
  { value: "Alex Brush", label: "Alex Brush", category: "Script", google: true },
  { value: "Allura", label: "Allura", category: "Script", google: true },
  { value: "Pinyon Script", label: "Pinyon Script", category: "Script", google: true },
];

export function getGoogleFontStylesheetUrl(fonts: FontOption[]): string {
  const googleFonts = fonts.filter(f => f.google).map(f => f.value.replace(/ /g, '+'));
  if (googleFonts.length === 0) return '';
  return `https://fonts.googleapis.com/css2?${googleFonts.map(f => `family=${f}:wght@400;700`).join('&')}&display=swap`;
}
