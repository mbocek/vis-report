import org.apache.commons.cli.Option


printHeaderIfNeeded = { scriptName ->
    /**
     * Parse arguments
     */ 
    parseArgs = { args -> 
        def cli = new CliBuilder(usage: "${scriptName} -[hsu]")
        // Create the list of options.
        cli.with {
            h longOpt: 'help', 'Show usage information'
            s longOpt: 'source', args: Option.UNLIMITED_VALUES, argName: 'source', 'Name of data source', required: true
            u longOpt: 'url', args: Option.UNLIMITED_VALUES, argName: 'url', 'Server url', required: true
        }
        
        def options = cli.parse(args)
        // Show usage text when -h or --help option is used.
        options || System.exit(1)

        if (options.h) {
            cli.usage()
            System.exit(1)
        }
     
    	def source
        if (options.s) {
    		source = options.s
    	} else {
    		source = 'vis-skoly'
    	}

    	def url
        if (options.u) {
    		url = options.u
    	} else {
    		url = 'http://localhost:8080/'
    	}
    	
    	[source,url]
    }
}
