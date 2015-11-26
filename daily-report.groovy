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
def outputFilePath = "daily-report-${dateTo.format('dd.MM.yyyy')}.xls"
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
def queryDetail = "select ob.datum, ob.druh, ob.ev_cislo, ob.pocet, ji.nazev, str.skupina_no, str.jmeno " +
            "from objednav as ob, jidelnic as ji, stravnik as str " + 
            "where ob.datum = :toDate " + 
            "and ji.datum = ob.datum " +
            "and ji.druh = ob.druh " +
            "and ob.ev_cislo = str.ev_cislo " + 
            "and trim(ob.druh) IN ('1', '2', '3', '4', '5') " + 
            "order by str.jmeno, ob.druh desc"

			
// Setup sheet
def sheetSetup = workbook.createSheet('Setup', 0)
fillSetup(sheetSetup)

// report sheet
def sheet = workbook.createSheet('Report', -1)

def row = 0
def col = 0

cellFormat = new WritableCellFormat()
format = new WritableCellFormat(cellFormat)
format.setBackground(Colour.GRAY_25)

sheet.addCell(new Label(col++, row, "Odb\u011Bratel", format))
sheet.addCell(new Label(col++, row, "Druh j\u00EDdla", format))
sheet.addCell(new Label(col++, row, "N\u00E1zev j\u00EDdla", format))
sheet.addCell(new Label(col++, row, "Po\u010Det porc\u00ED", format))
sheet.addCell(new Label(col++, row, "Maso", format))
sheet.addCell(new Label(col++, row, "P\u0159\u00EDloha", format))
sheet.addCell(new Label(col++, row, "Om\u00E1\u010Dka", format))
sheet.addCell(new Label(col++, row, "Zeleninov\u00E9 a sladk\u00E9 j\u00EDdlo", format))
sheet.addCell(new Label(col++, row, "Pol\u00E9vka", format))

sql.rows(queryDetail, [toDate: to]).each {
    col = 0
	row++
	sheet.addCell(new Label(col++, row, it.jmeno))
	sheet.addCell(new Label(col++, row, it.druh))
	sheet.addCell(new Label(col++, row, it.nazev))
	sheet.addCell(new Number(col++, row, it.pocet))
	sheet.addCell(new Formula(col++, row, "D${row + 1}*Setup!A${Integer.valueOf(it.skupina_no) + 1}"))
	sheet.addCell(new Formula(col++, row, "D${row + 1}*Setup!B${Integer.valueOf(it.skupina_no) + 1}"))
	sheet.addCell(new Formula(col++, row, "D${row + 1}*Setup!C${Integer.valueOf(it.skupina_no) + 1}"))
	sheet.addCell(new Formula(col++, row, "D${row + 1}*Setup!D${Integer.valueOf(it.skupina_no) + 1}"))
	sheet.addCell(new Formula(col++, row, "D${row + 1}*Setup!E${Integer.valueOf(it.skupina_no) + 1}"))
}

for(int x = 0; x < 3; x++) {
    cellView = sheet.getColumnView(x)
    cellView.setAutosize(true)
    sheet.setColumnView(x, cellView)
}    
for(int x = 3; x < 9; x++) {
    cellView = sheet.getColumnView(x)
    cellView.setSize(8*256)
    sheet.setColumnView(x, cellView)
}    


workbook.write() 
workbook.close()
sql.close()


println "Report generated."