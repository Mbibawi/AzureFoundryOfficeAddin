/// <reference path="../office.d.ts" />
import {byID} from "./taskpane.js"

class OfficeHelper {
 
   show(id: string): void {
    const el = byID<HTMLElement>(id);
    if (el) el.classList.remove('hidden');
  }

   hide(id: string): void {
    const el = byID<HTMLElement>(id);
    if (el) el.classList.add('hidden');
  }

   toggle(id: string, visible: boolean): void {
    visible ? this.show(id) : this.hide(id);
  }

  /** Generic document-text reader using file slices — works across all Office hosts. */
   async _getDocumentText(): Promise<string> {
    return new Promise((resolve) => {
      if (typeof Office === 'undefined' || !Office.context?.document) {
        resolve('');
        return;
      }

      Office.context.document.getFileAsync(Office.FileType.Text, { sliceSize: 65536 }, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve('');
          return;
        }

        const file = result.value;
        const slices: string[] = [];
        let idx = 0;

        const getSlice = (): void => {
          file.getSliceAsync(idx, (sr) => {
            if (sr.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              resolve(slices.join(''));
              return;
            }
            slices.push(sr.value.data as string);
            idx++;
            if (idx < file.sliceCount) {
              getSlice();
            } else {
              file.closeAsync();
              resolve(slices.join(''));
            }
          });
        };
        getSlice();
      });
    });
  }

   async _getSelectedText(): Promise<string> {
    return new Promise((resolve) => {
      if (typeof Office === 'undefined' || !Office.context?.document) { resolve(''); return; }
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? (result.value || '') : '');
      });
    });
  }

  async setSelectedText(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      if (typeof Office === 'undefined' || !Office.context?.document) { reject('Office not initialized'); return; }
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error?.message || 'Failed to set selected text');
        }
      });
    });
  }
}

abstract class OfficeAppBase extends OfficeHelper {

  /** 'document' | 'spreadsheet' | 'presentation' — used for UI copy. */
  getHostLabel(): string { return 'document'; }

  /** Host-specific suggested prompts shown on the empty-state screen. */
  getQuickPrompts(): QuickPrompt[] { return []; }

  async replaceSelection(text: string): Promise<void> {
    return this.setSelectedText(text);
  }

  async insertText(text: string): Promise<void> {
    return this.setSelectedText(text);
  }

  abstract getSelectedText(): Promise<string>;
  abstract getDocumentText(): Promise<string>;
  // Add this helper to the OfficeAppBase class in OfficeApp.ts
  async getActiveContext(): Promise<{ text: string; source: 'selection' | 'document' }> {
    const selection = await this.getSelectedText();
    // If the user has highlighted text, focus solely on that (Copilot-style)
    if (selection && selection.trim().length > 0) {
      return { text: selection, source: 'selection' };
    }
    // Fallback to the full document/sheet/slide content
    const fullDoc = await this.getDocumentText();
    return { text: fullDoc, source: 'document' };
  }
}

// ─── Word ───────────────────────────────────────────────────────────────────

export class WordApp extends OfficeAppBase {

  getHostLabel(): string { return 'document'; }

  getQuickPrompts(): QuickPrompt[] {
    return [
      { icon: '📋', label: 'Summarise document', prompt: 'Summarise the key points of this document in bullet points.' },
      { icon: '✍️', label: 'Review & improve',   prompt: 'Review this document and suggest improvements to clarity, structure, and style.' },
      { icon: '⚖️', label: 'Legal review',        prompt: 'List any potential legal, contractual, or compliance issues in this document.' },
      { icon: '🌍', label: 'Translate to French', prompt: 'Translate this document to French.' },
    ];
  }

  /** Uses Word.Body.text for a clean plain-text extraction; falls back to file slices. */
  async getDocumentText(): Promise<string> {
    try {
      return await Word.run(async (context) => {
        context.document.body.load('text');
        await context.sync();
        return context.document.body.text;
      });
    } catch {
      return this._getDocumentText();
    }
  }

  async getSelectedText(): Promise<string> {
    try {
      return await Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        return sel.text;
      });
    } catch {
      return this._getSelectedText();
    }
  }

  /** Overwrites the current selection with text. */
  async replaceSelection(text: string): Promise<void> {
    await Word.run(async (context) => {
      context.document.getSelection().insertText(text, 'Replace');
      await context.sync();
    });
  }

  /** Inserts text immediately after the current selection / cursor position. */
  async insertText(text: string): Promise<void> {
    await Word.run(async (context) => {
      context.document.getSelection().insertText(text, 'After');
      await context.sync();
    });
  }

  /** Appends a new paragraph at the end of the document body. */
  async insertParagraphAtEnd(text: string): Promise<void> {
    await Word.run(async (context) => {
      context.document.body.insertParagraph(text, 'End');
      await context.sync();
    });
  }

  /** Returns character count, word count, and raw text of the current selection. */
  async getSelectionInfo(): Promise<SelectionInfo> {
    try {
      return await Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        const t = sel.text;
        return {
          text: t,
          charCount: t.length,
          wordCount: t.trim() ? t.trim().split(/\s+/).length : 0,
          isEmpty: t.length === 0,
        };
      });
    } catch {
      const t = await this.getSelectedText();
      return {
        text: t,
        charCount: t.length,
        wordCount: t.trim() ? t.trim().split(/\s+/).length : 0,
        isEmpty: t.length === 0,
      };
    }
  }

  /** Applies character formatting to the current selection. */
  async formatSelection(options: WordFormatOptions): Promise<void> {
    await Word.run(async (context) => {
      const font = context.document.getSelection().font;
      if (options.bold      !== undefined) font.bold      = options.bold;
      if (options.italic    !== undefined) font.italic    = options.italic;
      if (options.underline !== undefined) font.underline = options.underline;
      if (options.fontSize  !== undefined) font.size      = options.fontSize;
      if (options.color     !== undefined) font.color     = options.color;
      if (options.fontName  !== undefined) font.name      = options.fontName;
      await context.sync();
    });
  }

  /** Applies a built-in or custom paragraph style (e.g. 'Heading 1', 'Normal') to the selection. */
  async applyStyleToSelection(styleName: string): Promise<void> {
    await Word.run(async (context) => {
      context.document.getSelection().style = styleName;
      await context.sync();
    });
  }

  /**
   * Anchors a comment on the current selection.
   * Requires Word API 1.4+; throws on older versions.
   */
  async insertComment(commentText: string): Promise<void> {
    await Word.run(async (context) => {
      context.document.getSelection().insertComment(commentText);
      await context.sync();
    });
  }

  /**
   * Replaces every occurrence of searchText with replacement text.
   * Returns the number of substitutions made.
   */
  async findAndReplace(
    searchText: string,
    replacement: string,
    options: Word.SearchOptions = {}
  ): Promise<number> {
    return Word.run(async (context) => {
      const results = context.document.body.search(searchText, options);
      results.load('items');
      await context.sync();
      results.items.forEach(r => r.insertText(replacement, 'Replace'));
      await context.sync();
      return results.items.length;
    });
  }

  /** Returns non-empty paragraph strings in document order. */
  async getParagraphs(): Promise<string[]> {
    try {
      return await Word.run(async (context) => {
        const paras = context.document.body.paragraphs;
        paras.load('items/text');
        await context.sync();
        return paras.items.map(p => p.text).filter(t => t.trim().length > 0);
      });
    } catch {
      return [];
    }
  }

  /** Returns heading outline, paragraph count, and total word count for the document. */
  async getDocumentStructure(): Promise<DocumentStructure> {
    try {
      return await Word.run(async (context) => {
        const body  = context.document.body;
        const paras = body.paragraphs;
        body.load('text');
        paras.load('items/text,items/style');
        await context.sync();

        const headings: DocumentStructure['headings'] = [];
        let paragraphCount = 0;

        paras.items.forEach(p => {
          if (p.text.trim()) paragraphCount++;
          const m = (p.style ?? '').match(/Heading\s+(\d+)/i);
          if (m) headings.push({ text: p.text.trim(), level: parseInt(m[1], 10) });
        });

        const wordCount = body.text.trim() ? body.text.trim().split(/\s+/).length : 0;
        return { headings, paragraphCount, wordCount };
      });
    } catch {
      return { headings: [], paragraphCount: 0, wordCount: 0 };
    }
  }
}

// ─── Excel ──────────────────────────────────────────────────────────────────

export class ExcelApp extends OfficeAppBase {

  getHostLabel(): string { return 'spreadsheet'; }

  getQuickPrompts(): QuickPrompt[] {
    return [
      { icon: '📊', label: 'Analyse data',    prompt: 'Analyse the data in this spreadsheet and summarise the key trends and insights.' },
      { icon: '🔢', label: 'Suggest formula', prompt: 'Based on the data structure, suggest useful Excel formulas I could apply and explain how to use them.' },
      { icon: '📋', label: 'Summarise data',  prompt: 'Provide a concise summary of what this data represents and its key statistics.' },
      { icon: '🔍', label: 'Find anomalies',  prompt: 'Identify any anomalies, outliers, or data quality issues in this spreadsheet.' },
    ];
  }

  /** Returns the active sheet's used range as a tab-separated plain-text representation. */
  async getDocumentText(): Promise<string> {
    try {
      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveSheet();
        const used  = sheet.getUsedRange();
        sheet.load('name');
        used.load('address,values,rowCount,columnCount');
        await context.sync();

        const lines: string[] = [`Sheet: ${sheet.name}`, `Range: ${used.address}`, ''];
        for (const row of used.values) {
          lines.push(row.map(v => (v === null || v === undefined) ? '' : String(v)).join('\t'));
        }
        return lines.join('\n');
      });
    } catch {
      return this._getDocumentText();
    }
  }

  /** Returns the selected range as a tab-separated block with an address header line. */
  async getSelectedText(): Promise<string> {
    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address,values');
        await context.sync();

        const lines: string[] = [`Selection: ${range.address}`, ''];
        for (const row of range.values) {
          lines.push(row.map(v => (v === null || v === undefined) ? '' : String(v)).join('\t'));
        }
        return lines.join('\n');
      });
    } catch {
      return this._getSelectedText();
    }
  }

  /**
   * Writes TSV-formatted text into the selected range.
   * Data is parsed by newlines (rows) and tabs (columns) and fitted to the selection shape.
   */
  async replaceSelection(text: string): Promise<void> {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('rowCount,columnCount');
      await context.sync();

      const srcRows = text.trimEnd().split('\n').map(r => r.split('\t'));
      const values: (string | number | boolean)[][] = [];
      for (let r = 0; r < range.rowCount; r++) {
        const row: (string | number | boolean)[] = [];
        for (let c = 0; c < range.columnCount; c++) {
          row.push(srcRows[r]?.[c] ?? '');
        }
        values.push(row);
      }
      range.values = values;
      await context.sync();
    });
  }

  /** Returns address, raw values, formula strings, and dimensions of the selected range. */
  async getSelectedRange(): Promise<RangeInfo> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('address,values,formulas,rowCount,columnCount');
      await context.sync();
      return {
        address:  range.address,
        values:   range.values as CellData[][],
        formulas: range.formulas,
        rowCount: range.rowCount,
        colCount: range.columnCount,
      };
    });
  }

  /** Returns all data from the active sheet's used range. */
  async getActiveSheetData(): Promise<SheetData> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveSheet();
      const used  = sheet.getUsedRange();
      sheet.load('name');
      used.load('address,values,rowCount,columnCount');
      await context.sync();
      return {
        name:     sheet.name,
        address:  used.address,
        values:   used.values as CellData[][],
        rowCount: used.rowCount,
        colCount: used.columnCount,
      };
    });
  }

  async getActiveSheetName(): Promise<string> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveSheet();
      sheet.load('name');
      await context.sync();
      return sheet.name;
    });
  }

  async getAllSheetNames(): Promise<string[]> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      return sheets.items.map(s => s.name);
    });
  }

  /** Writes a single value to a cell on the active sheet (e.g. address = 'B2'). */
  async setCellValue(address: string, value: string | number): Promise<void> {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getActiveSheet().getRange(address).values = [[value]];
      await context.sync();
    });
  }

  /** Writes a 2-D array of values to a range on the active sheet (e.g. address = 'A1:C3'). */
  async setRangeValues(address: string, values: (string | number | boolean)[][]): Promise<void> {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getActiveSheet().getRange(address).values = values;
      await context.sync();
    });
  }

  /** Writes an Excel formula string (e.g. '=SUM(A1:A10)') to a cell on the active sheet. */
  async insertFormula(address: string, formula: string): Promise<void> {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getActiveSheet().getRange(address).formulas = [[formula]];
      await context.sync();
    });
  }

  /** Returns the active sheet name, all sibling sheet names, and used-range dimensions. */
  async getActiveSheetInfo(): Promise<SheetInfo> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const active = sheets.getActiveSheet();
      const used   = active.getUsedRange();
      sheets.load('items/name');
      active.load('name');
      used.load('address,rowCount,columnCount');
      await context.sync();
      return {
        name:        active.name,
        allSheets:   sheets.items.map(s => s.name),
        usedAddress: used.address,
        rowCount:    used.rowCount,
        colCount:    used.columnCount,
      };
    });
  }
}

// ─── PowerPoint ─────────────────────────────────────────────────────────────

export class PowerPointApp extends OfficeAppBase {

  getHostLabel(): string { return 'presentation'; }

  getQuickPrompts(): QuickPrompt[] {
    return [
      { icon: '📋', label: 'Summarise slides',   prompt: 'Summarise the key points from all slides in this presentation.' },
      { icon: '✍️', label: 'Improve slide text', prompt: 'Review all slide text and suggest improvements for clarity and impact.' },
      { icon: '🎤', label: 'Talking points',      prompt: 'Generate concise speaker talking points for each slide.' },
      { icon: '🌍', label: 'Translate to French', prompt: 'Translate all slide content in this presentation to French.' },
    ];
  }

  /** Returns all slide text as a structured plain-text document. */
  async getDocumentText(): Promise<string> {
    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        slides.items.forEach(s => {
          s.shapes.load('items/name,items/textFrame/hasText,items/textFrame/textRange/text');
        });
        await context.sync();

        const parts: string[] = [];
        slides.items.forEach((slide, i) => {
          const texts = slide.shapes.items
            .filter(sh => sh.textFrame?.hasText)
            .map(sh => sh.textFrame.textRange.text.trim())
            .filter(Boolean);

          if (texts.length > 0) {
            parts.push(`--- Slide ${i + 1} ---\n${texts.join('\n')}`);
          }
        });

        return parts.join('\n\n');
      });
    } catch {
      return this._getDocumentText();
    }
  }

  async getSelectedText(): Promise<string> {
    //!need to see if we need to implement this for powerpoint
    return "";

  }

  /**
   * Returns structured data for every slide: index, title, body text.
   * Speaker notes are populated if PowerPoint API 1.2+ is available, otherwise left empty.
   */
  async getAllSlidesText(): Promise<SlideData[]> {
    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        slides.items.forEach(s => {
          s.shapes.load('items/name,items/textFrame/hasText,items/textFrame/textRange/text');
        });
        await context.sync();

        return slides.items.map((slide, i) => {
          let title = '';
          const body: string[] = [];

          slide.shapes.items.forEach(sh => {
            if (!sh.textFrame?.hasText) return;
            const text = sh.textFrame.textRange.text.trim();
            if (!text) return;
            if (!title && sh.name.toLowerCase().includes('title')) {
              title = text;
            } else {
              body.push(text);
            }
          });

          // Promote first body text to title when no title placeholder found
          if (!title && body.length > 0) title = body.shift() ?? '';

          return { index: i + 1, title, bodyText: body.join('\n'), speakerNotes: '' };
        });
      });
    } catch {
      return [];
    }
  }

  async getSlideCount(): Promise<number> {
    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        return slides.items.length;
      });
    } catch {
      return 0;
    }
  }

  /**
   * Returns speaker notes text for every slide.
   * Requires PowerPoint API 1.2+; returns empty strings on older runtimes.
   */
  async getSpeakerNotes(): Promise<string[]> {
    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        slides.items.forEach(s => s.notes.body.load('text'));
        await context.sync();
        return slides.items.map(s => s.notes?.body?.text ?? '');
      });
    } catch {
      return [];
    }
  }

  /**
   * Replaces speaker notes for each slide.
   * noteTexts[i] maps to slide i+1; excess slides are left unchanged.
   * Requires PowerPoint API 1.2+.
   */
  async setSpeakerNotes(noteTexts: string[]): Promise<void> {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const toUpdate = slides.items.slice(0, noteTexts.length);
      toUpdate.forEach(s => s.notes.body.paragraphs.load('items/text'));
      await context.sync();

      // Clear existing paragraphs and write the new text in the first paragraph
      toUpdate.forEach((s, i) => {
        const paras = s.notes.body.paragraphs.items;
        if (paras.length > 0) {
          // Overwrite first paragraph; additional paragraphs are left as-is
          (paras[0] as any).text = noteTexts[i];
        }
      });
      await context.sync();
    });
  }
}
