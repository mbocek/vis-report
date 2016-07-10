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

evaluate(new File("./functions/FeedFunction.groovy"))
printHeaderIfNeeded("feed-materials")

def (String dsn, String url) = parseArgs(args)
println "Feeding materials from dsn: ${dsn} in url: ${url}"

def json = new JsonBuilder()

// database charset setup
def prop = new java.util.Properties();
prop.put("charSet", "cp1250");

// connection string to jdbc odbc bridge
def sql = Sql.newInstance("jdbc:odbc:${dsn}", prop, "sun.jdbc.odbc.JdbcOdbcDriver")

// query to get all persons in system
def queryAllMaterials = "select * from suroviny"
            
List result = []

sql.eachRow(queryAllMaterials) { suroviny ->
    def code = suroviny.cissur
    def name = suroviny.nazev
    def totalWeight = suroviny.hmotnost
    def meatWeight = suroviny.hmotnost_m

    result << ["name": name, "code": code, "totalWeight": totalWeight, "meatWeight": meatWeight]
}

json(result)
//println json.toPrettyString()

def client = new RESTClient(url)
def response = client.post(
    contentType: JSON,
    requestContentType: JSON,
    path: '/store/material',
    body: json.toString())

assert response.status == 200

println "Data imported."