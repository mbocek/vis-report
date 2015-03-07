/**
 * Generating report from VIS database.
 * Database is Visual foxpro.
 */

import groovy.sql.Sql

import java.sql.Date as SqlDate
import java.sql.SQLException

import jxl.*
import jxl.write.*

// Parameters to report. Limits for time period.
def fromString = '2015-01-01'
def toString = '2015-01-31'
def dateFrom = Date.parse('yyyy-MM-dd', fromString)
def dateTo = Date.parse('yyyy-MM-dd', toString)

println "Running report for period: ${dateFrom.format('dd.MM.yyyy')}-${dateTo.format('dd.MM.yyyy')}"

// dsn name in odbc 32bit windows setup
def dsn = "vis-firmy"

def from = new SqlDate(dateFrom.getTime())
def to = new SqlDate(dateTo.getTime())

// report path - can be full path or relative path 
def outputFilePath = "report-${fromString}-${toString}.xls"
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
def queryDetail = "select ob.datum, ob.druh, ob.ev_cislo, ob.pocet, ji.nazev, ji.cena1 " + 
            "from objednav as ob, jidelnic as ji " + 
            "where ob.datum between :fromDate and :toDate " + 
            "and ji.datum = ob.datum " +
            "and ji.druh = ob.druh " +
            "and ob.ev_cislo = :evCislo " + 
            "order by ob.datum asc, ob.druh, ob.datcas_obj desc"
// query to get summary for orders between period and for specified person
def querySum = "select ji.cena1, sum(ob.pocet) as pocet " + 
            "from objednav as ob, jidelnic as ji " + 
            "where ob.datum between :fromDate and :toDate " + 
            "and ji.datum = ob.datum " +
            "and ji.druh = ob.druh " +
            "and ob.ev_cislo = :evCislo " + 
            "group by ji.cena1"
            
def sheetSummary = workbook.createSheet('Summary', -1)
def sheetSummaryRow = 0;
def total = 0.0

sql.eachRow(queryAllPersons) { stravnik ->
    def name = stravnik.jmeno
    def evCislo = stravnik.ev_cislo
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
        sql.rows(queryDetail, [fromDate: from, toDate: to, evCislo: evCislo]).each {
            if (lastDatum == null || lastDruh == null || lastDatum != it.datum || lastDruh != it.druh) {
                int col = 0
                row++
                int formulaRow = row+1
                sheet.addCell(new DateTime(col++, row, it.datum, dateFormat))
                sheet.addCell(new Label(col++, row, it.druh))
                sheet.addCell(new Label(col++, row, it.nazev))
                sheet.addCell(new Number(col++, row, it.pocet))
                sheet.addCell(new Number(col++, row, it.cena1))
                //Create a formula for adding cells
                Formula sum = new Formula(col++, row, "D" + formulaRow + "*E" + formulaRow);
                sheet.addCell(sum);
                total += it.pocet * it.cena1 
                subTotal += it.pocet * it.cena1 
            }
            lastDatum = it.datum
            lastDruh = it.druh
        }
        
        if (row > 0) {
            ++row
            sheet.mergeCells(0, row, 4, row);
            sheet.addCell(new Label(0, row, "Celkova suma:"))
            Formula sum = new Formula(5, row, "SUM(F2:F" + row + ")")
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