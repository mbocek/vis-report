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

evaluate(new File("./function/daily-report-function.groovy"))

def (Date dateTo, String dsn) = parseArgs(args)
println "Running report for: ${dateTo.format('dd.MM.yyyy')}"

def to = new SqlDate(dateTo.getTime())

// report path - can be full path or relative path 
def outputFilePath = "reports/daily-report-export-${dateTo.format('dd.MM.yyyy')}.xls"
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
def queryDetail = "select ob.datum, ob.druh, sum(ob.pocet) pocet, ji.nazev, kat.skupina_no, kat.popis, kat.kateg, ji.cissur " +
            "from objednav as ob, jideln as ji, stravnik as str, kateg kat " + 
            "where ob.datum = :toDate " + 
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
def sheet = workbook.createSheet('Report', -1)

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
        
        def koef = koeficient(it.druh, it.skupina_no, koeficientRows)

        cissurCode = cissurItem.split(' ')[0]
        cissurResult = sql.rows("select * from suroviny where cissur = :cissur", [cissur: cissurCode])
        // always only one row
        cissurResult.each { suroviny ->
            sheet.addCell(new Number(col++, row, it.pocet, format))
            sheet.addCell(new Label(col++, row, suroviny.nazev, format))
            sheet.addCell(new Formula(col++, row, "${suroviny.hmotnost}*${koef}", format))
            sheet.addCell(new Formula(col++, row, "${suroviny.hmotnost_m}*${koef}" , format))
            sheet.addCell(new Formula(col++, row, "E${row + 1}*G${row + 1}/1000", format))
            sheet.addCell(new Formula(col++, row, "E${row + 1}*H${row + 1}/1000", format))
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
