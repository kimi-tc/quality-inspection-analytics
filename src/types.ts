import React from 'react';

export interface ImportedRow {
  date: string;
  session: string;
  batch: string;
  category: string;
  attribute: string;
  declarations: number;
  ambiguousPasses: number;
  rejects: number;
  proofRejects: number;
}

export interface ParsedWorkbook {
  rows: ImportedRow[];
  importedAt: string;
  sourceName: string;
}

export interface SharedDatasetResponse extends ParsedWorkbook {}

export interface MetricsCardData {
  title: string;
  value: string;
  hint: string;
  icon: React.ReactNode;
  tone: 'slate' | 'emerald' | 'blue' | 'amber' | 'rose';
}
