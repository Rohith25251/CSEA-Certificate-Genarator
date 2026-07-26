'use client';

import React, { useRef, useState } from 'react';
import { substitutePlaceholders } from '@/lib/placeholder-engine';
import { downloadPdfFromElement } from '@/lib/pdf-generator';
import { Download, Loader2, Award } from 'lucide-react';

interface CertificateRendererProps {
  htmlTemplate: string;
  context: {
    certificateId: string;
    studentId: string;
    issueDate: string;
    eventDate?: string;
    eventName: string;
    row: Record<string, any>;
  };
  filename?: string;
  allowDownload?: boolean;
}

export const CertificateRenderer: React.FC<CertificateRendererProps> = ({
  htmlTemplate,
  context,
  filename = 'CSEA_Certificate.pdf',
  allowDownload = true
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);

  const renderedHtml = substitutePlaceholders(htmlTemplate, context);

  const handleDownload = async () => {
    if (!containerRef.current) return;
    setIsGeneratingPdf(true);
    try {
      await downloadPdfFromElement(containerRef.current, filename);
    } catch (err) {
      console.error('Failed to generate PDF:', err);
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  return (
    <div className="flex flex-col items-center w-full">
      
      {/* High-DPI Container View */}
      <div
        ref={containerRef}
        className="w-full bg-white shadow-xl rounded-xl overflow-hidden border border-slate-200"
        dangerouslySetInnerHTML={{ __html: renderedHtml }}
      />

      {allowDownload && (
        <div className="mt-4 flex justify-center">
          <button
            onClick={handleDownload}
            disabled={isGeneratingPdf}
            className="flex items-center space-x-2 bg-blue-600 hover:bg-blue-700 text-white font-extrabold px-6 py-3 rounded-xl text-xs uppercase tracking-wider shadow-md transition-colors"
          >
            {isGeneratingPdf ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin" />
                <span>Generating High-DPI PDF...</span>
              </>
            ) : (
              <>
                <Download className="w-4 h-4" />
                <span>Download High-DPI PDF</span>
              </>
            )}
          </button>
        </div>
      )}

    </div>
  );
};
