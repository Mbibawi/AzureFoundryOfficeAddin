import { byID } from "./taskpane.js";
class OfficeHelper {
    show(id) {
        const el = byID(id);
        if (el)
            el.classList.remove('hidden');
    }
    hide(id) {
        const el = byID(id);
        if (el)
            el.classList.add('hidden');
    }
    toggle(id, visible) {
        visible ? this.show(id) : this.hide(id);
    }
    async _getDocumentText() {
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
                const slices = [];
                let idx = 0;
                const getSlice = () => {
                    file.getSliceAsync(idx, (sr) => {
                        if (sr.status !== Office.AsyncResultStatus.Succeeded) {
                            file.closeAsync();
                            resolve(slices.join(''));
                            return;
                        }
                        slices.push(sr.value.data);
                        idx++;
                        if (idx < file.sliceCount) {
                            getSlice();
                        }
                        else {
                            file.closeAsync();
                            resolve(slices.join(''));
                        }
                    });
                };
                getSlice();
            });
        });
    }
    async _getSelectedText() {
        return new Promise((resolve) => {
            if (typeof Office === 'undefined' || !Office.context?.document) {
                resolve('');
                return;
            }
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                resolve(result.status === Office.AsyncResultStatus.Succeeded ? (result.value || '') : '');
            });
        });
    }
    async setSelectedText(text) {
        return new Promise((resolve, reject) => {
            if (typeof Office === 'undefined' || !Office.context?.document) {
                reject('Office not initialized');
                return;
            }
            Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                }
                else {
                    reject(result.error?.message || 'Failed to set selected text');
                }
            });
        });
    }
}
class OfficeAppBase extends OfficeHelper {
    getHostLabel() { return 'document'; }
    getQuickPrompts() { return []; }
    async replaceSelection(text) {
        return this.setSelectedText(text);
    }
    async insertText(text) {
        return this.setSelectedText(text);
    }
    async getActiveContext() {
        const selection = await this.getSelectedText();
        if (selection && selection.trim().length > 0) {
            return { text: selection, source: 'selection' };
        }
        const fullDoc = await this.getDocumentText();
        return { text: fullDoc, source: 'document' };
    }
}
export class WordApp extends OfficeAppBase {
    getHostLabel() { return 'document'; }
    getQuickPrompts() {
        return [
            { icon: '📋', label: 'Summarise document', prompt: 'Summarise the key points of this document in bullet points.' },
            { icon: '✍️', label: 'Review & improve', prompt: 'Review this document and suggest improvements to clarity, structure, and style.' },
            { icon: '⚖️', label: 'Legal review', prompt: 'List any potential legal, contractual, or compliance issues in this document.' },
            { icon: '🌍', label: 'Translate to French', prompt: 'Translate this document to French.' },
        ];
    }
    async getDocumentText() {
        try {
            return await Word.run(async (context) => {
                context.document.body.load('text');
                await context.sync();
                return context.document.body.text;
            });
        }
        catch {
            return this._getDocumentText();
        }
    }
    async getSelectedText() {
        try {
            return await Word.run(async (context) => {
                const sel = context.document.getSelection();
                sel.load('text');
                await context.sync();
                return sel.text;
            });
        }
        catch {
            return this._getSelectedText();
        }
    }
    async replaceSelection(text) {
        await Word.run(async (context) => {
            context.document.getSelection().insertText(text, 'Replace');
            await context.sync();
        });
    }
    async insertText(text) {
        await Word.run(async (context) => {
            context.document.getSelection().insertText(text, 'After');
            await context.sync();
        });
    }
    async insertParagraphAtEnd(text) {
        await Word.run(async (context) => {
            context.document.body.insertParagraph(text, 'End');
            await context.sync();
        });
    }
    async getSelectionInfo() {
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
        }
        catch {
            const t = await this.getSelectedText();
            return {
                text: t,
                charCount: t.length,
                wordCount: t.trim() ? t.trim().split(/\s+/).length : 0,
                isEmpty: t.length === 0,
            };
        }
    }
    async formatSelection(options) {
        await Word.run(async (context) => {
            const font = context.document.getSelection().font;
            if (options.bold !== undefined)
                font.bold = options.bold;
            if (options.italic !== undefined)
                font.italic = options.italic;
            if (options.underline !== undefined)
                font.underline = options.underline;
            if (options.fontSize !== undefined)
                font.size = options.fontSize;
            if (options.color !== undefined)
                font.color = options.color;
            if (options.fontName !== undefined)
                font.name = options.fontName;
            await context.sync();
        });
    }
    async applyStyleToSelection(styleName) {
        await Word.run(async (context) => {
            context.document.getSelection().style = styleName;
            await context.sync();
        });
    }
    async insertComment(commentText) {
        await Word.run(async (context) => {
            context.document.getSelection().insertComment(commentText);
            await context.sync();
        });
    }
    async findAndReplace(searchText, replacement, options = {}) {
        return Word.run(async (context) => {
            const results = context.document.body.search(searchText, options);
            results.load('items');
            await context.sync();
            results.items.forEach(r => r.insertText(replacement, 'Replace'));
            await context.sync();
            return results.items.length;
        });
    }
    async getParagraphs() {
        try {
            return await Word.run(async (context) => {
                const paras = context.document.body.paragraphs;
                paras.load('items/text');
                await context.sync();
                return paras.items.map(p => p.text).filter(t => t.trim().length > 0);
            });
        }
        catch {
            return [];
        }
    }
    async getDocumentStructure() {
        try {
            return await Word.run(async (context) => {
                const body = context.document.body;
                const paras = body.paragraphs;
                body.load('text');
                paras.load('items/text,items/style');
                await context.sync();
                const headings = [];
                let paragraphCount = 0;
                paras.items.forEach(p => {
                    if (p.text.trim())
                        paragraphCount++;
                    const m = (p.style ?? '').match(/Heading\s+(\d+)/i);
                    if (m)
                        headings.push({ text: p.text.trim(), level: parseInt(m[1], 10) });
                });
                const wordCount = body.text.trim() ? body.text.trim().split(/\s+/).length : 0;
                return { headings, paragraphCount, wordCount };
            });
        }
        catch {
            return { headings: [], paragraphCount: 0, wordCount: 0 };
        }
    }
}
export class ExcelApp extends OfficeAppBase {
    getHostLabel() { return 'spreadsheet'; }
    getQuickPrompts() {
        return [
            { icon: '📊', label: 'Analyse data', prompt: 'Analyse the data in this spreadsheet and summarise the key trends and insights.' },
            { icon: '🔢', label: 'Suggest formula', prompt: 'Based on the data structure, suggest useful Excel formulas I could apply and explain how to use them.' },
            { icon: '📋', label: 'Summarise data', prompt: 'Provide a concise summary of what this data represents and its key statistics.' },
            { icon: '🔍', label: 'Find anomalies', prompt: 'Identify any anomalies, outliers, or data quality issues in this spreadsheet.' },
        ];
    }
    async getDocumentText() {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveSheet();
                const used = sheet.getUsedRange();
                sheet.load('name');
                used.load('address,values,rowCount,columnCount');
                await context.sync();
                const lines = [`Sheet: ${sheet.name}`, `Range: ${used.address}`, ''];
                for (const row of used.values) {
                    lines.push(row.map(v => (v === null || v === undefined) ? '' : String(v)).join('\t'));
                }
                return lines.join('\n');
            });
        }
        catch {
            return this._getDocumentText();
        }
    }
    async getSelectedText() {
        try {
            return await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load('address,values');
                await context.sync();
                const lines = [`Selection: ${range.address}`, ''];
                for (const row of range.values) {
                    lines.push(row.map(v => (v === null || v === undefined) ? '' : String(v)).join('\t'));
                }
                return lines.join('\n');
            });
        }
        catch {
            return this._getSelectedText();
        }
    }
    async replaceSelection(text) {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('rowCount,columnCount');
            await context.sync();
            const srcRows = text.trimEnd().split('\n').map(r => r.split('\t'));
            const values = [];
            for (let r = 0; r < range.rowCount; r++) {
                const row = [];
                for (let c = 0; c < range.columnCount; c++) {
                    row.push(srcRows[r]?.[c] ?? '');
                }
                values.push(row);
            }
            range.values = values;
            await context.sync();
        });
    }
    async getSelectedRange() {
        return Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('address,values,formulas,rowCount,columnCount');
            await context.sync();
            return {
                address: range.address,
                values: range.values,
                formulas: range.formulas,
                rowCount: range.rowCount,
                colCount: range.columnCount,
            };
        });
    }
    async getActiveSheetData() {
        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveSheet();
            const used = sheet.getUsedRange();
            sheet.load('name');
            used.load('address,values,rowCount,columnCount');
            await context.sync();
            return {
                name: sheet.name,
                address: used.address,
                values: used.values,
                rowCount: used.rowCount,
                colCount: used.columnCount,
            };
        });
    }
    async getActiveSheetName() {
        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveSheet();
            sheet.load('name');
            await context.sync();
            return sheet.name;
        });
    }
    async getAllSheetNames() {
        return Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load('items/name');
            await context.sync();
            return sheets.items.map(s => s.name);
        });
    }
    async setCellValue(address, value) {
        await Excel.run(async (context) => {
            context.workbook.worksheets.getActiveSheet().getRange(address).values = [[value]];
            await context.sync();
        });
    }
    async setRangeValues(address, values) {
        await Excel.run(async (context) => {
            context.workbook.worksheets.getActiveSheet().getRange(address).values = values;
            await context.sync();
        });
    }
    async insertFormula(address, formula) {
        await Excel.run(async (context) => {
            context.workbook.worksheets.getActiveSheet().getRange(address).formulas = [[formula]];
            await context.sync();
        });
    }
    async getActiveSheetInfo() {
        return Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            const active = sheets.getActiveSheet();
            const used = active.getUsedRange();
            sheets.load('items/name');
            active.load('name');
            used.load('address,rowCount,columnCount');
            await context.sync();
            return {
                name: active.name,
                allSheets: sheets.items.map(s => s.name),
                usedAddress: used.address,
                rowCount: used.rowCount,
                colCount: used.columnCount,
            };
        });
    }
}
export class PowerPointApp extends OfficeAppBase {
    getHostLabel() { return 'presentation'; }
    getQuickPrompts() {
        return [
            { icon: '📋', label: 'Summarise slides', prompt: 'Summarise the key points from all slides in this presentation.' },
            { icon: '✍️', label: 'Improve slide text', prompt: 'Review all slide text and suggest improvements for clarity and impact.' },
            { icon: '🎤', label: 'Talking points', prompt: 'Generate concise speaker talking points for each slide.' },
            { icon: '🌍', label: 'Translate to French', prompt: 'Translate all slide content in this presentation to French.' },
        ];
    }
    async getDocumentText() {
        try {
            return await PowerPoint.run(async (context) => {
                const slides = context.presentation.slides;
                slides.load('items');
                await context.sync();
                slides.items.forEach(s => {
                    s.shapes.load('items/name,items/textFrame/hasText,items/textFrame/textRange/text');
                });
                await context.sync();
                const parts = [];
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
        }
        catch {
            return this._getDocumentText();
        }
    }
    async getSelectedText() {
        return "";
    }
    async getAllSlidesText() {
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
                    const body = [];
                    slide.shapes.items.forEach(sh => {
                        if (!sh.textFrame?.hasText)
                            return;
                        const text = sh.textFrame.textRange.text.trim();
                        if (!text)
                            return;
                        if (!title && sh.name.toLowerCase().includes('title')) {
                            title = text;
                        }
                        else {
                            body.push(text);
                        }
                    });
                    if (!title && body.length > 0)
                        title = body.shift() ?? '';
                    return { index: i + 1, title, bodyText: body.join('\n'), speakerNotes: '' };
                });
            });
        }
        catch {
            return [];
        }
    }
    async getSlideCount() {
        try {
            return await PowerPoint.run(async (context) => {
                const slides = context.presentation.slides;
                slides.load('items');
                await context.sync();
                return slides.items.length;
            });
        }
        catch {
            return 0;
        }
    }
    async getSpeakerNotes() {
        try {
            return await PowerPoint.run(async (context) => {
                const slides = context.presentation.slides;
                slides.load('items');
                await context.sync();
                slides.items.forEach(s => s.notes.body.load('text'));
                await context.sync();
                return slides.items.map(s => s.notes?.body?.text ?? '');
            });
        }
        catch {
            return [];
        }
    }
    async setSpeakerNotes(noteTexts) {
        await PowerPoint.run(async (context) => {
            const slides = context.presentation.slides;
            slides.load('items');
            await context.sync();
            const toUpdate = slides.items.slice(0, noteTexts.length);
            toUpdate.forEach(s => s.notes.body.paragraphs.load('items/text'));
            await context.sync();
            toUpdate.forEach((s, i) => {
                const paras = s.notes.body.paragraphs.items;
                if (paras.length > 0) {
                    paras[0].text = noteTexts[i];
                }
            });
            await context.sync();
        });
    }
}
//# sourceMappingURL=OfficeApp.js.map