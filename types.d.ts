type Config = {
    resource: string;
    project: string;
    apiKey: string;
}

type FoundryModel = {
    id: string;
    name: string;
    object: string;
    created?: number;
    owned_by?: string;
}

type ChatMessage = {
    role: 'user' | 'assistant' | 'system';
    content: string;
}

type FoundryResponse = { data?: FoundryModel[]; value?: FoundryModel[] }

type ModelMeta = { icon: string; colorClass: string; provider: string; }

// ─── Quick prompt ───────────────────────────────────────────────────────────

type QuickPrompt = {
    icon: string;
    label: string;
    prompt: string;
}

// ─── Word types ─────────────────────────────────────────────────────────────

type WordFormatOptions = {
    bold?: boolean;
    italic?: boolean;
    /** Underline style string, e.g. 'Single', 'Double', 'None'. */
    underline?: string;
    fontSize?: number;
    /** CSS-style hex color, e.g. '#FF0000'. */
    color?: string;
    fontName?: string;
}

type SelectionInfo = {
    text: string;
    charCount: number;
    wordCount: number;
    isEmpty: boolean;
}

type DocumentStructure = {
    headings: Array<{ text: string; level: number }>;
    paragraphCount: number;
    wordCount: number;
}

// ─── Excel types ────────────────────────────────────────────────────────────

type CellData = string | number | boolean | null;

type RangeInfo = {
    address: string;
    values: CellData[][];
    formulas: string[][];
    rowCount: number;
    colCount: number;
}

type SheetData = {
    name: string;
    address: string;
    values: CellData[][];
    rowCount: number;
    colCount: number;
}

type SheetInfo = {
    name: string;
    allSheets: string[];
    usedAddress: string;
    rowCount: number;
    colCount: number;
}

// ─── PowerPoint types ───────────────────────────────────────────────────────

type SlideData = {
    index: number;
    title: string;
    bodyText: string;
    speakerNotes: string;
}
