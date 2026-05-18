// Minimal Office.js type declarations
// Full types: npm install @types/office-js

declare namespace Office {
  function onReady(callback: () => void): void;

  enum CoercionType { Text = 'text', Html = 'html' }
  enum HostType { Word = 'Word', Excel = 'Excel', PowerPoint = 'PowerPoint' }
  enum FileType { Text = 'text', Compressed = 'compressed', Pdf = 'pdf' }
  enum AsyncResultStatus { Succeeded = 'succeeded', Failed = 'failed' }

  const context: {
    document: OfficeDocument;
    host: HostType;
  };

  interface AsyncResult<T> {
    status: AsyncResultStatus;
    value:  T;
    error?: { message: string };
  }

  interface File {
    sliceCount: number;
    getSliceAsync(sliceIndex: number, callback: (result: AsyncResult<Slice>) => void): void;
    closeAsync(): void;
  }

  interface Slice {
    data: string | Uint8Array;
  }

  interface OfficeDocument {
    getFileAsync(
      fileType: FileType,
      options:  { sliceSize: number },
      callback: (result: AsyncResult<File>) => void
    ): void;
    getSelectedDataAsync(
      coercionType: CoercionType,
      callback: (result: AsyncResult<string>) => void
    ): void;
    setSelectedDataAsync(
      data: string,
      options: { coercionType: CoercionType },
      callback: (result: AsyncResult<void>) => void
    ): void;
  }
}

// ─── OfficeExtension ────────────────────────────────────────────────────────

declare namespace OfficeExtension {
  class Error extends globalThis.Error {
    code: string;
    debugInfo: string;
    message: string;
    name: string;
    traceMessages: string[];
  }

  interface ClientResult<T> {
    value: T;
  }
}

// ─── Word JavaScript API ────────────────────────────────────────────────────

declare namespace Word {
  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;

  type InsertLocation = 'Before' | 'After' | 'Start' | 'End' | 'Replace';

  interface RequestContext {
    document: Document;
    sync(): Promise<void>;
  }

  interface Document {
    body: Body;
    getSelection(): Range;
    save(): void;
  }

  interface Body {
    text: string;
    paragraphs: ParagraphCollection;
    insertParagraph(paragraphText: string, insertLocation: InsertLocation | string): Paragraph;
    insertText(text: string, insertLocation: InsertLocation | string): Range;
    search(searchText: string, searchOptions?: SearchOptions): RangeCollection;
    load(propertyNames: string | string[]): Body;
  }

  interface Range {
    text: string;
    html: string;
    style: string;
    font: Font;
    paragraphs: ParagraphCollection;
    insertText(text: string, insertLocation: InsertLocation | string): Range;
    insertParagraph(paragraphText: string, insertLocation: InsertLocation | string): Paragraph;
    /** Adds a comment anchored to this range. Requires Word API 1.4+. */
    insertComment(commentText: string): Comment;
    select(selectionMode?: string): void;
    delete(): void;
    load(propertyNames: string | string[]): Range;
  }

  interface RangeCollection {
    items: Range[];
    load(propertyNames: string | string[]): RangeCollection;
  }

  interface Font {
    bold: boolean;
    italic: boolean;
    underline: string;
    size: number;
    color: string;
    name: string;
  }

  interface Paragraph {
    text: string;
    style: string;
    font: Font;
    insertText(text: string, insertLocation: InsertLocation | string): Range;
    insertParagraph(paragraphText: string, insertLocation: InsertLocation | string): Paragraph;
    load(propertyNames: string | string[]): Paragraph;
  }

  interface ParagraphCollection {
    items: Paragraph[];
    load(propertyNames: string | string[]): ParagraphCollection;
  }

  interface Comment {
    content: string;
    load(propertyNames: string | string[]): Comment;
  }

  interface SearchOptions {
    ignorePunct?: boolean;
    ignoreSpace?: boolean;
    matchCase?: boolean;
    matchPrefix?: boolean;
    matchSuffix?: boolean;
    matchWholeWord?: boolean;
    matchWildcards?: boolean;
  }
}

// ─── Excel JavaScript API ───────────────────────────────────────────────────

declare namespace Excel {
  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;

  interface RequestContext {
    workbook: Workbook;
    sync(): Promise<void>;
  }

  interface Workbook {
    worksheets: WorksheetCollection;
    getSelectedRange(): Range;
  }

  interface WorksheetCollection {
    items: Worksheet[];
    getActiveSheet(): Worksheet;
    load(propertyNames: string | string[]): WorksheetCollection;
  }

  interface Worksheet {
    name: string;
    id: string;
    getUsedRange(valuesOnly?: boolean): Range;
    getRange(address: string): Range;
    activate(): void;
    load(propertyNames: string | string[]): Worksheet;
  }

  interface Range {
    address: string;
    values: (string | number | boolean)[][];
    formulas: string[][];
    text: string[][];
    rowCount: number;
    columnCount: number;
    numberFormat: string[][];
    format: RangeFormat;
    load(propertyNames: string | string[]): Range;
  }

  interface RangeFormat {
    fill: RangeFill;
    font: RangeFont;
    horizontalAlignment: string;
    verticalAlignment: string;
  }

  interface RangeFill {
    color: string;
  }

  interface RangeFont {
    bold: boolean;
    italic: boolean;
    color: string;
    size: number;
  }
}

// ─── PowerPoint JavaScript API ──────────────────────────────────────────────

declare namespace PowerPoint {
  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;

  interface RequestContext {
    presentation: Presentation;
    sync(): Promise<void>;
  }

  interface Presentation {
    slides: SlideCollection;
    title: string;
    load(propertyNames: string | string[]): Presentation;
  }

  interface SlideCollection {
    items: Slide[];
    load(propertyNames: string | string[]): SlideCollection;
  }

  interface Slide {
    id: string;
    /** Speaker notes. Requires PowerPoint API 1.2+. */
    notes: SlideNotes;
    shapes: ShapeCollection;
    load(propertyNames: string | string[]): Slide;
  }

  interface SlideNotes {
    body: NotesBody;
    load(propertyNames: string | string[]): SlideNotes;
  }

  interface NotesBody {
    text: string;
    paragraphs: NotesParagraphCollection;
    load(propertyNames: string | string[]): NotesBody;
  }

  interface NotesParagraphCollection {
    items: NotesParagraph[];
    load(propertyNames: string | string[]): NotesParagraphCollection;
  }

  interface NotesParagraph {
    text: string;
    load(propertyNames: string | string[]): NotesParagraph;
  }

  interface ShapeCollection {
    items: Shape[];
    load(propertyNames: string | string[]): ShapeCollection;
  }

  interface Shape {
    id: string;
    name: string;
    type: string;
    textFrame: TextFrame;
    load(propertyNames: string | string[]): Shape;
  }

  interface TextFrame {
    hasText: boolean;
    textRange: TextRange;
    load(propertyNames: string | string[]): TextFrame;
  }

  interface TextRange {
    text: string;
    load(propertyNames: string | string[]): TextRange;
  }
}
