import jxl.*
import jxl.write.*
import jxl.format.*

/**
 * Parse arguments
 */ 
parseArgs = { args -> 
    def cli = new CliBuilder(usage: 'daily-report -[hds]')
    // Create the list of options.
    cli.with {
        h longOpt: 'help', 'Show usage information'
        d longOpt: 'date', args: 1, argName: 'date', 'Date in format dd.MM.yyyy'
        s longOpt: 'source', args: 1, argName: 'source', 'Name of data source'
    }
    
    def options = cli.parse(args)
    if (!options) {
        return
    }
    // Show usage text when -h or --help option is used.
    if (options.h) {
        cli.usage()
        System.exit(1)
    }
 
	def date
	if (options.d) {
		date = Date.parse('dd.MM.yyyy', options.d)
	} else {
		date = new Date()
	}
 
	def source
    if (options.s) {
		source = options.s
	} else {
		source = 'vis-skoly'
	}
	
	[date, source]
}

fillSetup = { sheetSetup ->

    def row = 0
    def col = 0
	sheetSetup.addCell(new Label(col++, row, "Koeficient"))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 0.5/1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 1/1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 1/1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 0.7/1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 1/1000))
}
