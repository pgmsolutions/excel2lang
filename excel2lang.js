#!/usr/bin/env node

// Import
const path = require('path'),
      fs = require('fs'),
      XLSX = require('xlsx'),
      mkdirp = require('mkdirp');

class excel2lang {
    constructor(){
        this.filepath = null;
        this.currentFile = null;
        this.languages = [];
        this.lang = {};
        this.missing = {};
    }

    setDestination(dest){
        // Save current file if any
        if(this.currentFile !== null)
            this.saveCurrent();

        // Set up a new file
        this.currentFile = dest;
        for(let i = 0; i < this.languages.length; ++i){
            this.lang[this.languages[i]] = {};
            this.missing[this.languages[i]] = 0;
        }
    }

    go(){
        // Command line
        let fp = require('yargs').argv._;
        if(fp.length < 1){
            console.error('No file entered');
            return;
        }
        this.filepath = path.normalize(fp[0]);

        // Check if source file exists
        if(!fs.existsSync(this.filepath)){
            console.error('File "'+this.filepath+'" was not found!');
            return;
        }

        // Opening file
        let workbook = XLSX.readFile(this.filepath);

        // Reading sheets
        let sheetname = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetname];

        // Getting max range
        let range = XLSX.utils.decode_range(sheet['!ref']);

        // Go through lines
        for(let rowNum = range.s.r; rowNum <= range.e.r; rowNum++){
            // Get line as array
            let row = [];
            for(let colNum = range.s.c; colNum <= range.e.c; colNum++){
                let cell = sheet[XLSX.utils.encode_cell({r: rowNum, c: colNum})];
                if(typeof cell === 'undefined')
                    row.push(null);
                else
                    row.push(cell.v);
            }
        
            // Pass if empty or comment line
            if(row.length < 2 || row[1] === null) continue;

            // Each command
            let instr = row.shift();
            if(instr == '#')
                continue;
            else if(instr == 'C')
                this.languages = row.slice(1);
            else if(instr == 'F' && row.length > 0){
                if(this.languages.length === 0){
                    console.error('File line called before Code line!');
                    return;
                }
                this.setDestination(row[0]);
            }
            else if(row.length > 2){
                let stringid = row.shift();
                for(let i = 0; i < row.length; ++i){
                    if(i > this.languages.length) continue;
                    if(row[i] !== null && row[i].length > 0){
                        this.lang[this.languages[i]][stringid] = row[i];
                    }
                    else {
                        // count error in missing
                        this.missing[this.languages[i]] += 1;
                        this.lang[this.languages[i]][stringid] = row[0];
                    }
                }
            }
        }

        // Save last portion
        this.saveCurrent();

        // Show errors
        for(let i = 0; i < this.languages.length; ++i)
            if(this.missing[this.languages[i]] > 0)
                console.log(this.languages[i]+' has '+this.missing[this.languages[i]]+' missing entries.');
        
        console.log('Finished!');
    }

    saveCurrent(){
        let folder = path.dirname(path.join(path.dirname(this.filepath), this.currentFile));
        mkdirp.sync(folder);
        for(let i = 0; i < this.languages.length; ++i){
            let f = path.join(path.dirname(this.filepath), this.currentFile.replace(/{lang}/g, this.languages[i]));
            fs.writeFileSync(f, JSON.stringify(this.lang[this.languages[i]], null, '\t'));
        }
    }
};

let i = new excel2lang();
i.go();

// EoF