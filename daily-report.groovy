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

evaluate(new File("./functions/daily-report-function.groovy"))

def (Date dateTo, String dsn) = parseArgs(args)
println "Running report for: ${dateTo.format('dd.MM.yyyy')}"

def to = new SqlDate(dateTo.getTime())

// report path - can be full path or relative path 
def outputFilePath = "reports/daily-report-${dateTo.format('dd.MM.yyyy')}.xls"
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

// query to get all orders for date
def queryDetail = "select ob.datum, ob.druh, ob.datcas_obj, ob.ev_cislo, ob.pocet, ji.nazev, str.skupina_no, str.jmeno, str.kateg, ji.cissur " +
            "from objednav as ob, jideln as ji, stravnik as str " + 
            "where ob.datum = :toDate " + 
            "and ji.datum = ob.datum " +
            "and ji.druh = ob.druh " +
            "and ob.ev_cislo = str.ev_cislo " + 
            "and trim(ob.druh) IN ('1', '2', '3', '4', '5') " + 
            "order by str.kateg, str.jmeno, ob.druh, ob.datcas_obj desc"

			
// Setup sheet
def sheetSetup = workbook.createSheet('Setup', 0)
fillSetup(sheetSetup)

// report sheet
def sheet = workbook.createSheet('Report', -1)

def row = 0
def col = 0

colours = [Colour.LIGHT_GREEN, Colour.LIGHT_TURQUOISE, Colour.VERY_LIGHT_YELLOW]

cellFormat = new WritableCellFormat()
format = new WritableCellFormat(cellFormat)
format.setBackground(Colour.LIGHT_BLUE)

sheet.addCell(new Label(col++, row, "Kategorie", format))
sheet.addCell(new Label(col++, row, "Odb\u011Bratel", format))
sheet.addCell(new Label(col++, row, "Druh j\u00EDdla", format))
sheet.addCell(new Label(col++, row, "N\u00E1zev j\u00EDdla", format))
sheet.addCell(new Label(col++, row, "Po\u010Det porc\u00ED", format))
sheet.addCell(new Label(col++, row, "Suroviny", format))
sheet.addCell(new Label(col++, row, "Jenotkovy objem", format))
sheet.addCell(new Label(col++, row, "Z toho masa", format))
sheet.addCell(new Label(col++, row, "Celkovy objem", format))
sheet.addCell(new Label(col++, row, "Z toho masa", format))

def lastEvCislo
def lastDruh
def lastKateg
colorIndex = 0
sql.rows(queryDetail, [toDate: to]).each {
    if (lastKateg == null || lastKateg != it.kateg) {
        cellFormat = new WritableCellFormat()
        format = new WritableCellFormat(cellFormat)
        format.setBackground(colours[colorIndex % colours.size()])
        colorIndex++
    }
	if (lastEvCislo == null || lastDruh == null || lastEvCislo != it.ev_cislo || lastDruh != it.druh) {
		col = 0
		row++
		sheet.addCell(new Label(col++, row, it.kateg, format))
		sheet.addCell(new Label(col++, row, it.jmeno.trim(), format))
		sheet.addCell(new Label(col++, row, it.druh, format))
		sheet.addCell(new Label(col++, row, it.nazev.trim(), format))
		cissurData = it.cissur.split("\\r?\\n")
		cissurData.each() { cissurItem ->
            if (col == 0) {
                (0..3).each {
                    sheet.addCell(new Label(col++, row, "", format))
                }
            }
            
			cissurCode = cissurItem.split(' ')[0]
			cissurResult = sql.rows("select * from suroviny where cissur = :cissur", [cissur: cissurCode])
            // always only one row
            cissurResult.each { suroviny ->
                sheet.addCell(new Number(col++, row, it.pocet, format))
                sheet.addCell(new Label(col++, row, suroviny.nazev, format))
                sheet.addCell(new Number(col++, row, suroviny.hmotnost, format))
                sheet.addCell(new Number(col++, row, suroviny.hmotnost_m, format))
                sheet.addCell(new Formula(col++, row, "E${row + 1}*Setup!A${Integer.valueOf(it.skupina_no) + 1}*${suroviny.hmotnost}", format))
                sheet.addCell(new Formula(col++, row, "E${row + 1}*Setup!A${Integer.valueOf(it.skupina_no) + 1}*${suroviny.hmotnost_m}", format))
			}
            col = 0 // reset to subset
            (cissurData.last() == cissurItem) ? row : row++
		}
	}
	lastEvCislo = it.ev_cislo
	lastDruh = it.druh
    lastKateg = it.kateg
}

col = 0
cellView = sheet.getColumnView(col)
cellView.setAutosize(true)
sheet.setColumnView(col, cellView)

col = 1
cellView = sheet.getColumnView(col)
cellView.setAutosize(true)
sheet.setColumnView(col, cellView)

col = 2
cellView = sheet.getColumnView(col)
cellView.setSize(5*256)
sheet.setColumnView(col, cellView)

col = 3
cellView = sheet.getColumnView(col)
cellView.setSize(30*256)
sheet.setColumnView(col, cellView)

col = 4
cellView = sheet.getColumnView(col)
cellView.setSize(8*256)
sheet.setColumnView(col, cellView)

col = 5
cellView = sheet.getColumnView(col)
cellView.setSize(20*256)
sheet.setColumnView(col, cellView)

col = 6
cellView = sheet.getColumnView(col)
cellView.setSize(10*256)
sheet.setColumnView(col, cellView)

col = 7
cellView = sheet.getColumnView(col)
cellView.setSize(10*256)
sheet.setColumnView(col, cellView)

col = 8
cellView = sheet.getColumnView(col)
cellView.setSize(10*256)
sheet.setColumnView(col, cellView)

col = 9
cellView = sheet.getColumnView(col)
cellView.setSize(10*256)
sheet.setColumnView(col, cellView)


workbook.write() 
workbook.close()
sql.close()


println "Report generated."
