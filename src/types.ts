import React from 'react';

export interface ImportedRow {
  date: string;
  auditor?: string;
  auditorTeam?: string;
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
  importHistory?: ImportRecord[];
}

export interface SharedDatasetResponse extends ParsedWorkbook {}

export interface EfficiencyRow {
  date: string;
  employee: string;
  team: string;
  session: string;
  batch: string;
  handledCount: number;
  weightedHandledCount: number;
  firstAuditCount: number;
  firstAuditPassCount: number;
  precisionPassCount: number;
  auditNotPassCount: number;
  proofRefusalCount: number;
  ambiguousCount: number;
  passRate: number;
  precisionPassRate: number;
  proofAccuracy: number;
  avgHandleMinutes: number;
  timeoutCount: number;
}

export interface ParsedEfficiencyWorkbook {
  rows: EfficiencyRow[];
  importedAt: string;
  sourceName: string;
  importHistory?: ImportRecord[];
}

export interface EfficiencyDatasetResponse extends ParsedEfficiencyWorkbook {}

export interface ImportRecord {
  id: string;
  sourceName: string;
  importedAt: string;
  rowCount: number;
  dataType: 'quality' | 'efficiency';
}

export interface PropertyCategoryEntry {
  propertyName: string;
  category: string;
}

export interface PropertyCategoryDictionaryResponse {
  entries: PropertyCategoryEntry[];
}

export interface AuditorTeamEntry {
  auditorName: string;
  team: string;
}

export interface AuditorTeamDictionaryResponse {
  entries: AuditorTeamEntry[];
}

export interface AiAnalysisResponse {
  model: string;
  analysis: string;
  generatedAt: string;
}

export interface MetricsCardData {
  title: string;
  value: string;
  hint: string;
  icon: React.ReactNode;
  tone: 'slate' | 'emerald' | 'blue' | 'amber' | 'rose';
}
