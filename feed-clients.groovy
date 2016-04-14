/**
 * Feed server with client data.
 */
@Grab('org.codehaus.groovy.modules.http-builder:http-builder:0.7')

import groovy.sql.Sql

import java.sql.Date as SqlDate
import java.sql.SQLException
import java.util.HashMap
import java.util.Map

import groovy.json.JsonBuilder

import groovyx.net.http.RESTClient
import static groovyx.net.http.ContentType.*

evaluate(new File("./functions/feed-clients-function.groovy"))

def (String dsn, String url) = parseArgs(args)
println "Feeding clients from dsn: ${dsn} in url: ${url}"

def json = new JsonBuilder()

// database charset setup
def prop = new java.util.Properties();
prop.put("charSet", "cp1250");

// connection string to jdbc odbc bridge
def sql = Sql.newInstance("jdbc:odbc:${dsn}", prop, "sun.jdbc.odbc.JdbcOdbcDriver")

// query to get all persons in system
def queryAllPersons = "select * from stravnik"
            
List result = []

sql.eachRow(queryAllPersons) { stravnik ->
    def name = stravnik.jmeno
    def code = stravnik.ev_cislo
    def group = stravnik.cen_skup
    def category = stravnik.kateg

    result << ["name": name, "code": code, "groupId": group, "category": category]
}

json(result)
//println json.toPrettyString()

def client = new RESTClient(url)
def response = client.post(
    contentType: JSON,
    requestContentType: JSON,
    path: '/store/client',
    body: json.toString())

assert response.status == 200

println "Data imported."