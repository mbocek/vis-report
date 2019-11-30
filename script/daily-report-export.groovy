/**
 * Generating report from VIS database.
 * Database is Visual foxpro.
 */
import groovy.sql.Sql

import java.sql.Date as SqlDate
import java.sql.SQLException
import java.util.HashMap
import java.util.Map

import jxl.*
import jxl.write.*
import jxl.format.*

def dateFileFormat = 'dd.MM.yyyy'

def cli = new CliBuilder(usage: 'daily-report -[hds]')
// Create the list of options.
cli.with {
    h longOpt: 'help', 'Show usage information'
    d longOpt: 'date', args: 1, argName: 'date', 'Date in format dd.MM.yyyy'
    s longOpt: 'source', args: 1, argName: 'source', 'Name of data source'
}

def options = cli.parse(args)
// Show usage text when -h or --help option is used.
options || System.exit(1)

if (options.h) {
    cli.usage()
    System.exit(1)
}

def dateTo = options.d ? Date.parse('dd.MM.yyyy', options.d) : new Date()
def dsn = options.s ? options.s : 'vis-skoly'
def reports = '../reports'

println "Running report for: ${dateTo.format('dd.MM.yyyy')}"
println "DSN: ${dsn}"

def to = new SqlDate(dateTo.getTime())

// report path - can be full path or relative path 
def outputFilePath = "${reports}/daily-report-export-${dateTo.format('dd.MM.yyyy')}.xls"
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

report = { skupina, fullname ->

    // query to get all orders for date
    def queryDetail = "select ob.datum, ob.druh, sum(ob.pocet) pocet, ji.nazev, kat.skupina_no, kat.popis, kat.kateg, ji.cissur " +
                "from objednav as ob, jideln as ji, stravnik as str, kateg kat " + 
                "where ob.datum = :toDate " + 
                "and str.skupina = :skupina " +
                "and ji.datum = ob.datum " +
                "and ji.druh = ob.druh " +
                "and ob.ev_cislo = str.ev_cislo " + 
                "and str.kateg = kat.kateg " + 
                "group by ob.datum, ob.druh, kat.kateg"

    def koeficientSQL = "select skupina, druh, koeficient " +
                "from fin_lim " +
                "order by druh, skupina"
    def koeficientRows = sql.rows(koeficientSQL)

    // report sheet
    def sheet = workbook.createSheet(fullname, -1)

    def row = 0
    def col = 0

    cellFormat = new WritableCellFormat()
    format = new WritableCellFormat(cellFormat)
    format.setBackground(Colour.LIGHT_BLUE)

    sheet.addCell(new Label(col++, row, "Kategorie", format))
    sheet.addCell(new Label(col++, row, "N\u00E1zev kategorie", format))
    sheet.addCell(new Label(col++, row, "Druh j\u00EDdla", format))
    sheet.addCell(new Label(col++, row, "N\u00E1zev j\u00EDdla", format))
    sheet.addCell(new Label(col++, row, "Po\u010Det porc\u00ED", format))
    sheet.addCell(new Label(col++, row, "Koeficient", format))
    sheet.addCell(new Label(col++, row, "Polozky", format))
    sheet.addCell(new Label(col++, row, "Realny objem", format))
    sheet.addCell(new Label(col++, row, "Jenotkovy objem", format))
    sheet.addCell(new Label(col++, row, "Celkovy objem", format))

    def lastEvCislo
    def lastDruh
    def lastKateg
    colorIndex = 0

    sql.rows(queryDetail, [toDate: to, skupina: skupina]).each {
        if (lastKateg == null || lastKateg != it.kateg) {
            cellFormat = new WritableCellFormat()
            cellFormat.setBorder(Border.TOP, BorderLineStyle.THIN);
            format = new WritableCellFormat(cellFormat)
        }

        col = 0
        row++
        sheet.addCell(new Label(col++, row, it.kateg, format))
        sheet.addCell(new Label(col++, row, it.popis.trim(), format))
        sheet.addCell(new Label(col++, row, it.druh, format))
        sheet.addCell(new Label(col++, row, it.nazev.trim(), format))
        cissurData = it.cissur.split("\\r?\\n")
        cissurData.each() { cissurItem ->
            if (col == 0) {
                (0..3).each {
                    sheet.addCell(new Label(col++, row, "", format))
                }
            }
            
            def koef = 1F
            koeficientRows.each() { koefItem -> 
                if (koefItem.druh == it.druh && koefItem.skupina == it.skupina_no) {
                    koef = koefItem.koeficient
                } 
            }
            
            cissurCode = cissurItem.split(' ')[0]
            cissurResult = sql.rows("select * from suroviny where cissur = :cissur", [cissur: cissurCode])
            // always only one row
            cissurResult.each { suroviny ->
                sheet.addCell(new Number(col++, row, it.pocet, format))
                sheet.addCell(new Number(col++, row, koef, format))
                sheet.addCell(new Label(col++, row, suroviny.nazev, format))
                sheet.addCell(new Number(col++, row, suroviny.hmotnost, format))
                sheet.addCell(new Formula(col++, row, "${suroviny.hmotnost}*${koef}", format))
                sheet.addCell(new Formula(col++, row, "E${row + 1}*I${row + 1}/1000", format))
            }
            col = 0 // reset to subset
            if( cissurData.last() == cissurItem) {
                row
            } else {
                row++
                cellFormat = new WritableCellFormat()
                cellFormat.setBorder(Border.BOTTOM, BorderLineStyle.NONE);
                format = new WritableCellFormat(cellFormat)
            } 
        }
        lastDruh = it.druh
        lastKateg = it.kateg
    }

    setColumn = { column, width ->
        cellView = sheet.getColumnView(column)
        cellView.setSize(width*256)
        sheet.setColumnView(column, cellView)
    }

    col = 0
    cellView = sheet.getColumnView(col)
    cellView.setAutosize(true)
    sheet.setColumnView(col, cellView)

    col = 1
    cellView = sheet.getColumnView(col)
    cellView.setAutosize(true)
    sheet.setColumnView(col, cellView)

    setColumn(2, 5)
    setColumn(3, 30)
    setColumn(4, 9)
    setColumn(5, 9)
    setColumn(6, 20)
    setColumn(7, 15)
    setColumn(8, 15)
    setColumn(9, 15)
}

sql.rows("select uc.skupina, uc.popis from uc_skup uc").each {
    report(it.skupina, it.popis)
}

workbook.write() 
workbook.close()
sql.close()


println "Report generated."
