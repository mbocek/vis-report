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
		date = Date.parse('dd.MM.yyyy', toString)
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
	sheetSetup.addCell(new Label(col++, row, "Maso"))
	sheetSetup.addCell(new Label(col++, row, "P\u0159\u00EDloha"))
	sheetSetup.addCell(new Label(col++, row, "Om\u00E1\u010Dka"))
	sheetSetup.addCell(new Label(col++, row, "Zeleninov\u00E9 a sladk\u00E9 j\u00EDdlo"))
	sheetSetup.addCell(new Label(col++, row, "Pol\u00E9vka"))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 50 / 1000))
	sheetSetup.addCell(new Number(col++, row, 150 / 1000))
	sheetSetup.addCell(new Number(col++, row, 100 / 1000))
	sheetSetup.addCell(new Number(col++, row, 200 / 1000))
	sheetSetup.addCell(new Number(col++, row, 200 / 1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 100 / 1000))
	sheetSetup.addCell(new Number(col++, row, 250 / 1000))
	sheetSetup.addCell(new Number(col++, row, 200 / 1000))
	sheetSetup.addCell(new Number(col++, row, 400 / 1000))
	sheetSetup.addCell(new Number(col++, row, 330 / 1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 100 / 1000))
	sheetSetup.addCell(new Number(col++, row, 250 / 1000))
	sheetSetup.addCell(new Number(col++, row, 200 / 1000))
	sheetSetup.addCell(new Number(col++, row, 400 / 1000))
	sheetSetup.addCell(new Number(col++, row, 330 / 1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 70 / 1000))
	sheetSetup.addCell(new Number(col++, row, 175 / 1000))
	sheetSetup.addCell(new Number(col++, row, 150 / 1000))
	sheetSetup.addCell(new Number(col++, row, 280 / 1000))
	sheetSetup.addCell(new Number(col++, row, 250 / 1000))
	col = 0
	row++
	sheetSetup.addCell(new Number(col++, row, 100 / 1000))
	sheetSetup.addCell(new Number(col++, row, 250 / 1000))
	sheetSetup.addCell(new Number(col++, row, 200 / 1000))
	sheetSetup.addCell(new Number(col++, row, 400 / 1000))
	sheetSetup.addCell(new Number(col++, row, 330 / 1000))
}
