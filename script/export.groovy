/**
 * Generating report from VIS database.
 * Database is Visual foxpro.
 */

import groovy.sql.Sql

import java.sql.Date as SqlDate
import java.sql.SQLException
import java.util.HashMap
import java.util.Map

import groovy.time.TimeCategory

import jxl.*
import jxl.write.*

def dateFileFormat = 'dd.MM.yyyy'

// dsn name in odbc 32bit windows setup
def dsn = "vis-firmy"
//def dsn = "vis-skolky"

def cli = new CliBuilder(usage: "export -[hsmr]")
// Create the list of options.
cli.with {
    h longOpt: 'help', 'Show usage information'
    s longOpt: 'source', args: 1, argName: 'source', 'Name of data source', required: true
    m longOpt: 'month', args: 1, argName: 'month', 'Report per month in shift: e.g. 0 current, -1 previouse month', required: false
    r longOpt: 'report', args: 1, argName: 'report', 'Report direcotry', required: false
}

def options = cli.parse(args)
// Show usage text when -h or --help option is used.
options || System.exit(1)

if (options.h) {
    cli.usage()
    System.exit(1)
}

if (options.s) {
    dsn = options.s
} 

def monthShift = options.m ? new Integer(options.m) : 0
def reports = options.r ? options.r : "../reports"

def calendar = Calendar.getInstance()
calendar.add(Calendar.MONTH, monthShift)
calendar.set(Calendar.DATE, 1)
calendar.set(Calendar.MILLISECOND, 0)
calendar.set(Calendar.SECOND, 0)
calendar.set(Calendar.MINUTE, 0)
calendar.set(Calendar.HOUR_OF_DAY, 0)
def dateFrom = calendar.clone().getTime()
calendar.add(Calendar.MONTH, 1)
calendar.add(Calendar.SECOND, -1)
def dateTo = calendar.clone().getTime()

println "Running report for period: ${dateFrom.format('dd.MM.yyyy')}-${dateTo.format('dd.MM.yyyy')}"
println "DSN: ${dsn}"

def from = new SqlDate(dateFrom.getTime())
def to = new SqlDate(dateTo.getTime())

// report path - can be full path or relative path 
def outputFilePath = "${reports}/report-${dsn}-${dateFrom.format('yyyy-MM-dd')}-${dateTo.format('yyyy-MM-dd')}.xls"
def ws = new WorkbookSettings()
ws.setEncoding("cp1250")
def workbook = Workbook.createWorkbook(new File(outputFilePath), ws)

// date format in excel
def customDateFormat = new DateFormat("dd.MM.yyyy"); 
dateFormat = new WritableCellFormat (customDateFormat); 

// database charset setup
def prop = new java.util.Properties();
prop.put("charSet", "cp1250");

// connection string to jdbc odbc bridge
def sql = Sql.newInstance("jdbc:odbc:${dsn}", prop, "sun.jdbc.odbc.JdbcOdbcDriver")

// query to get all persons in system
def queryAllPersons = "select * from stravnik"
// query to get all orders between period and for specified person
def queryDetail = "select ob.datum, ob.druh, ob.ev_cislo, ob.pocet, ji.nazev, " +
			"ji.cena1, ji.cena2, ji.cena3, ji.cena4, ji.cena5, ji.cena6, ji.cena7, ji.cena8, ji.cena9, ji.cena10, " + 
            "ji.cena11, ji.cena12, ji.cena13, ji.cena14, ji.cena15, ji.cena16, ji.cena17, ji.cena18, ji.cena19, ji.cena20, " + 
			"ji.cena21, ji.cena22, ji.cena23 , ji.cena24 , ji.cena25, ji.cena26, ji.cena27, ji.cena28, ji.cena29  " +
            "from objednav as ob, jidelnic as ji " + 
            "where ob.datum between :fromDate and :toDate " + 
            "and ji.datum = ob.datum " +
            "and ji.druh = ob.druh " +
            "and ob.ev_cislo = :evCislo " + 
            "order by ob.datum asc, ob.druh, ob.datcas_obj desc"
            
def sheetSummary = workbook.createSheet('Summary', -1)
def sheetSummaryRow = 0;
def total = 0.0

sql.eachRow(queryAllPersons) { stravnik ->
    def name = stravnik.jmeno
    def evCislo = stravnik.ev_cislo
    def cenovaSkupina = stravnik.cen_skup
    def sheet = workbook.createSheet(stravnik.jmeno, evCislo.intValue())
    sheet.addCell(new Label(0, 0, "Datum")) 
    sheet.addCell(new Label(1, 0, "Druh")) 
    sheet.addCell(new Label(2, 0, "J\u00EDdlo")) 
    sheet.addCell(new Label(3, 0, "Po\u010Det"))
    sheet.addCell(new Label(4, 0, "Cena"))
    sheet.addCell(new Label(5, 0, "Suma"))
    
    try {
        def subTotal = 0.0
        int row = 0
        def lastDatum = null
        def lastDruh = null
        def subSumTotal = new HashMap<Double, Integer>()
        sql.rows(queryDetail, [fromDate: from, toDate: to, evCislo: evCislo]).each {
            if (lastDatum == null || lastDruh == null || lastDatum != it.datum || lastDruh != it.druh) {
				def cena 
				switch (cenovaSkupina.trim()) {
					case "1": cena = it.cena1
							break
					case "2": cena = it.cena2
							break
					case "3": cena = it.cena3
							break
					case "4": cena = it.cena4
							break
					case "5": cena = it.cena5
							break
					case "6": cena = it.cena6
							break
					case "7": cena = it.cena7
							break
					case "8": cena = it.cena8
							break
					case "9": cena = it.cena9
							break
					case "10": cena = it.cena10
							break
					case "11": cena = it.cena11
							break
					case "12": cena = it.cena12
							break
					case "13": cena = it.cena13
							break
					case "14": cena = it.cena14
							break
					case "15": cena = it.cena15
							break
					case "16": cena = it.cena16
							break
					case "17": cena = it.cena17
							break
					case "18": cena = it.cena18
							break
					case "19": cena = it.cena19
							break
					case "20": cena = it.cena20
							break
					case "21": cena = it.cena21
							break
					case "22": cena = it.cena22
							break
					case "23": cena = it.cena23
							break
					case "24": cena = it.cena24
							break
					case "25": cena = it.cena25
							break
					case "26": cena = it.cena26
							break
					case "27": cena = it.cena27
							break
					case "28": cena = it.cena28
							break
					case "29": cena = it.cena29
							break
					default: 
							println "Nondefined price for cen_skup: $cenovaSkupina"
				}
                int col = 0
                row++
                int formulaRow = row+1
                sheet.addCell(new DateTime(col++, row, it.datum, dateFormat))
                sheet.addCell(new Label(col++, row, it.druh))
                sheet.addCell(new Label(col++, row, it.nazev))
                sheet.addCell(new Number(col++, row, it.pocet))
                sheet.addCell(new Number(col++, row, cena))
                //Create a formula for adding cells
                Formula sum = new Formula(col++, row, "D" + formulaRow + "*E" + formulaRow)
                sheet.addCell(sum);
                total += it.pocet * cena 
                subTotal += it.pocet * cena
                def count = subSumTotal.get(cena)
                count = (count == null) ? it.pocet : count + it.pocet
                subSumTotal.put(cena, count)
            }
            lastDatum = it.datum
            lastDruh = it.druh
        }
        
        def endRow = row + 1
        if (row > 0) {
            ++row
            ++row
            subSumTotal.keySet().each { cenaZaJidlo ->
                ++row
                sheet.addCell(new Label(2, row, "Suma za j\u00EDdlo"))
                sheet.addCell(new Number(3, row, subSumTotal.get(cenaZaJidlo)))
                sheet.addCell(new Number(4, row, cenaZaJidlo))
                sheet.addCell(new Number(5, row, cenaZaJidlo * subSumTotal.get(cenaZaJidlo)))
            }
            ++row
            sheet.mergeCells(0, row, 4, row);
            sheet.addCell(new Label(0, row, "Celkova suma:"))
            Formula sum = new Formula(5, row, "SUM(F2:F" + endRow + ")")
            sheet.addCell(sum)
        }
        sheetSummary.addCell(new Label(0, sheetSummaryRow, name))
        sheetSummary.addCell(new Number(1, sheetSummaryRow, subTotal))
        ++sheetSummaryRow;
    } catch (SQLException e) {
        println "No data found for ev_cislo: ${evCislo}" + e
    }
    
    for(int x = 0; x < 5; x++) {
        cell=sheet.getColumnView(x)
        cell.setAutosize(true)
        sheet.setColumnView(x, cell)
    }    
}
++sheetSummaryRow
sheetSummary.addCell(new Label(0, sheetSummaryRow, "Celkova suma:"))
sheetSummary.addCell(new Number(1, sheetSummaryRow, total))

for(int x = 0; x < 1; x++) {
    cell=sheetSummary.getColumnView(x)
    cell.setAutosize(true)
    sheetSummary.setColumnView(x, cell)
}    


workbook.write() 
workbook.close()
sql.close()


println "Report generated."
