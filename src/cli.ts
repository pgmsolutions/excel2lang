import colors from 'colors/safe';
import Excel2Lang, { Excel2LangResult } from './excel2lang';
import {program} from 'commander';
import Table from 'cli-table';

// Check a filepath is provided
program.parse(process.argv);
if(program.args.length === 0){
    console.error(colors.red('excel2lang: please provide a filepath as argument.'));
    process.exit(1);
}

// Create instance
const instance: Excel2Lang = new Excel2Lang();
try {
    // Process
    const result: Excel2LangResult = instance.process(program.args[0]);

    // Create result table
    const tbl: Table = new Table({
        style: {head: ['blue']},
        head: ['Language', 'Count', 'Percentage']
    });
    for(const [language, missing] of result.missing){
        // Percentage translated
        const percentage: number = ((result.count-missing)/result.count);

        // Color
        let color: string = percentage > 0.98 ? 'green' : 'yellow';
        if(percentage < 0.8){
            color = 'red';
        }

        // Push string array for table
        const missingString: string = missing === 0 ? '' : ` (-${missing})`;
        tbl.push([colors[color](language), colors[color](`${result.count-missing}/${result.count}${missingString}`), colors[color](`${(percentage*100).toFixed(2)}`)]);
    }

    // Display
    console.log('excel2lang: successfully created file languages.');
    console.log(tbl.toString());

    // List duplicate
    if(result.duplicate.size > 0){
        console.log(colors.yellow('excel2lang: warning: There are some duplicate entries:'));
        for(const entry of result.duplicate){
            console.log(colors.yellow(`- ${entry}`));
        }
    }
}
catch(e: unknown){
    const error: string = e instanceof Error ? e.message : `${e}`;
    console.error(colors.red(`excel2lang: error: ${error}`));
    process.exit(1);
}

// Return OK
process.exit(0);