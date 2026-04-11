export type LoginStatus = 'guest' | 'authorized' | 'error';

export type LoginUserResponse = {
  status: LoginStatus;
  message?: string;
  hasToken?: boolean;
  user?: {
    name: string;
    email: string;
  };
};

export type Recipient = {
  name: string;
  email: string;
  department?: string;
  grade?: string;
  field?: string;
};

export type Channel = {
  id: string;
  name: string;
  is_private?: boolean;
};

export type UserProfile = {
  name: string;
  nameEn: string;
  email: string;
  studentId: string;
  grade: string;
  field: string;
  phone: string;
  birthday: string;
  almaMater: string;
  carOwner: boolean;
  retired: boolean;
  continueNext: boolean;
  isAdmin: boolean;
};

export type SurveyItem = {
  title: string;
  spreadsheetId: string | null;
  spreadsheetUrl: string | null;
  formUrl?: string | null;
  inChargeOrg?: string;
  inChargeDept?: string;
  collecting?: boolean;
  scoreName?: string | null;
  scoreUnit?: string | null;
  userLatestRowIndex?: number | null;
  available?: boolean;
  latestResponseDate?: number | null;
  latestScore?: number | null;
  latestScoreFormatted?: string | null;
};

export type SurveyDetailResponse = {
  success: boolean;
  message?: string;
  sheetRef?: string;
  rowIndex?: number;
  headers?: string[];
  response?: {
    answers: Record<string, string | number | boolean>;
    timestamp: number | null;
    email: string | null;
    score: string | number | null;
    scoreFormatted?: string | null;
    studentId?: string | null;
  };
  scoreName?: string | null;
  scoreUnit?: string | null;
};

export type CollectionItem = {
  id: string;
  title: string;
  spreadsheetUrl: string;
  inChargeOrg: string;
  inChargeDept: string;
  createdAt: number | string | null;
  createdBy: string;
};

export type CollectionPerson = {
  email: string;
  expected: number;
  collected: number;
  status: string;
};

export type CollectionSummary = {
  success: boolean;
  message?: string;
  expectedTotal?: number;
  expectedCount?: number;
  collectedTotal?: number;
  collectedCount?: number;
  perPerson?: CollectionPerson[];
};

export type RosterCsvResponse = {
  success: boolean;
  message?: string;
  csv?: string;
  csvBase64?: string;
  excelBase64?: string;
  filename?: string;
  encoding?: string;
};
