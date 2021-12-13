// Imports
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';

export interface Excel2LangResult {
    success: boolean;
    count: number;
    missing: Map<string, number>;
    duplicate: Set<string>;
}

export default class Excel2Lang {
    /** The excel file path */
    private filepath: string = null;
    /** The current target file */
    private currentFile: string = null;
    /** List of languages, ordered by their position as in the sheet */
    private languagesList: Array<string> = [];
    /** Language keys and values */
    private languagesContent: Map<string, Map<string, string>> = new Map();
    /** Missing values counter for each language */
    private missing: Map<string, number> = new Map();
    /** Duplicated entries */
    private duplicated: Set<string> = new Set();
    /** Number of strings entries */
    private entriesCount: number = 0;

    /**
     * Sets the languages
     * @param languages an array of languages name
     */
    private setLanguages(languages: Array<string>): void {
        // Check if already set
        if(this.languagesContent.size > 0 || this.languagesList.length > 0){
            throw new Error('Languages codes already defined (more than one "C" lines).');
        }

        // Sanitize
        if(!Array.isArray(languages) || languages.length === 0){
            throw new Error('Incorrect languages codes.');
        }

        // Check if empty languages
        languages = languages.map((language: string) => language.trim());
        if(languages.some((language: string) => language.length === 0)){
            throw new Error('Empty language code detected.');
        }

        // Set language list
        this.languagesList = languages;

        // Create empty map for each language
        languages.forEach((language: string) => {
            this.languagesContent.set(language, new Map());
            this.missing.set(language, 0);
        });
    }

    /**
     * Sets the target destination file
     * @param destination the destination filename
     */
    private setDestination(destination: string): void {
        // Check that languages have already been defined
        if(this.languagesContent.size === 0){
            throw new Error('Error: File line (F) called before Languages line (C).');
        }

        // Sanitize
        if(typeof destination !== 'string' || destination.trim().length === 0){
            throw new Error('Invalid or empty destination file.');
        }

        // Save current file if any
        if(this.currentFile !== null){
            this.saveCurrent();
        }

        // Empty languages content
        this.currentFile = destination.trim();
        for(const [lang, content] of this.languagesContent){
            content.clear();
        }
    }

    /**
     * Adds a new language translation
     * @param row the row
     */
    private addLine(row: Array<string>){
        // Check that we have a target file selected
        // This also checks that we have at least one language
        if(this.currentFile === null){
            throw new Error('Error: String line called before target file (F line).');
        }

        // Sanitize
        if(!Array.isArray(row) || row.length < 2){
            return;
        }

        // Either empty line or string line
        const stringId: string = row.shift().trim();
        if(stringId.length === 0){
            return;
        }

        // For each columns
        for(let i: number = 0; i < this.languagesList.length; ++i){
            // Get language code & content
            const languageId: string = this.languagesList[i];
            const content: string = i < row.length ? row[i] : null;

            // If the first one is empty, abord
            if(i === 0 && (typeof content !== 'string' || content.trim().length > 0)){
                throw new Error(`${stringId} has no default value (first language).`);
            }

            // Test if string exists
            if(typeof content === 'string' && content.trim().length > 0 && this.languagesContent.has(languageId)){
                if(this.languagesContent.get(languageId).get(stringId)){
                    this.duplicated.add(stringId);
                }
                else {
                    ++this.entriesCount;
                }
                this.languagesContent.get(languageId).set(stringId, content);
            }
            else {
                // Count error in missing and set default value
                this.missing.set(languageId, this.missing.get(languageId)+1);
                this.languagesContent.get(languageId).set(stringId, row[0]);
            }
        }
    }

    /**
     * Processes a single row from the sheet
     * @param row an array of string
     */
    private processRow(row: Array<string>): void {
        // Sanitize
        if(!Array.isArray(row) || row.length === 0){
            throw new Error('Incorrect row content.');
        }

        // Remove empty endings cell
        let i: number = row.length-1;
        while(i >= 0){
            if(row[i].trim().length === 0){
                row.splice(i, 1);
                --i;
            }
            else {
                break;
            }
        }

        // Skip if empty
        if(row.length === 0){
            return;
        }

        // First column
        const instr: string = row.shift();
        if(instr === '#'){
            // Comment
            return;
        }
        else if(instr == 'C'){
            // Language lines
            this.setLanguages(row.slice(1));
        }
        else if(instr == 'F' && row.length > 0){
            // Destination file line
            this.setDestination(row[0]);
        }
        else {
            // Try to add line as strings
            this.addLine(row);
        }
    }

    /**
     * Saves current strings to target files
     */
    private saveCurrent(){
        const folder: string = path.dirname(path.join(path.dirname(this.filepath), this.currentFile));
        fs.mkdirSync(folder, {recursive: true});
        for(const [language, content] of this.languagesContent){
            const file: string = path.join(path.dirname(this.filepath), this.currentFile.replace(/{lang}/g, language));
            const objectContent: any = Object.create(null);
            for(const [key, value] of content){
                objectContent[key] = value;
            }
            fs.writeFileSync(file, JSON.stringify(objectContent, null, '\t'));
        }
    }
    
    /**
     * Processes an excel file
     * @param filepath the path to the excel file
     * @returns 
     */
    public process(filepath: string): Excel2LangResult {
        // Check if source file exists
        this.filepath = path.normalize(filepath);
        if(!fs.existsSync(this.filepath)){
            throw new Error(`File ${this.filepath} was not found.`);
        }

        // Opening file, get first sheet and max range
        const workbook: any = XLSX.readFile(this.filepath);
        const sheetname: any = workbook.SheetNames[0];
        const sheet: any = workbook.Sheets[sheetname];
        const range: any = XLSX.utils.decode_range(sheet['!ref']);

        // Go through each lines
        for(let rowNum: any = range.s.r; rowNum <= range.e.r; rowNum++){
            // Get line as array, incorrect values are empty strings
            const row: Array<string> = [];
            for(let colNum = range.s.c; colNum <= range.e.c; colNum++){
                const cell: any = sheet[XLSX.utils.encode_cell({r: rowNum, c: colNum})];
                row.push(typeof cell !== 'undefined' && cell !== null ? `${cell.v}` : '');
            }

            // Process
            this.processRow(row);
        }

        // Save last target
        this.saveCurrent();

        return {
            success: true,
            count: this.entriesCount,
            missing: this.missing,
            duplicate: this.duplicated
        };
    }
}