import java.net.*
import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential
import com.google.gdata.client.spreadsheet.*
import com.google.gdata.data.spreadsheet.*
import com.google.gdata.util.*

// Define CONST
APPLICATION_NAME = "google-spreadsheet-api-trial"
API_JSON_FILE = new File("src/main/resources/gapi.json")
SCOPES = [
    "https://spreadsheets.google.com/feeds"
]
SPREADSHEET_URL = "https://spreadsheets.google.com/feeds/spreadsheets/private/full"
SPREADSHEET_FEED_URL = null

try {
    SPREADSHEET_FEED_URL = new URL(SPREADSHEET_URL)
} catch (MalformedURLException e) {
    throw new RuntimeException(e)
}

// Utility Function
Credential authorize() {
    // println "authorize in"
    FileInputStream jsonKeyFile = new FileInputStream(API_JSON_FILE)
    GoogleCredential credential = GoogleCredential.fromStream(jsonKeyFile)
        .createScoped(SCOPES)

    boolean ret = credential.refreshToken()
    // debug dump
    // println "refreshToken: ${ret}"
    if (credential != null) {
        // println "AccessToken: ${credential.getAccessToken()}"
    }

    // println "authorize out"

    return credential
}

SpreadsheetService getService() {
    // println "service in"
    SpreadsheetService service = new SpreadsheetService(APPLICATION_NAME)
    service.setProtocolVersion(SpreadsheetService.Versions.V3)

    Credential credential = authorize()
    service.setOAuth2Credentials(credential)

    // debug dump
    // println "Schema: ${service.getSchema().toString()}"
    // println "Protocol: ${service.getProtocolVersion().getVersionString()}"
    // println "ServiceVersion: ${service.getServiceVersion()}"

    // println "service out"
    return service
}

List<SpreadsheetEntry> findAllSpreadsheets(SpreadsheetService service) {
    // println "findAllSpreadsheets in"

    SpreadsheetFeed feed = service.getFeed(SPREADSHEET_FEED_URL, SpreadsheetFeed.class)

    List<SpreadsheetEntry> spreadsheets = feed.getEntries()

    for (SpreadsheetEntry spreadsheet : spreadsheets) {
        // println "title: ${spreadsheet.getTitle().getPlainText()}"
    }

    // println "findAllSpreadsheets out"
    return spreadsheets
}

SpreadsheetEntry findAllSpreadsheetByName(SpreadsheetService service, String spreadsheetName) {
    // println "findAllSpreadsheetByName in"
    SpreadsheetQuery sheetQuery = new SpreadsheetQuery(SPREADSHEET_FEED_URL)
    sheetQuery.setTitleQuery(spreadsheetName)
    SpreadsheetFeed feed = service.query(sheetQuery, SpreadsheetFeed.class)
    SpreadsheetEntry ssEntry = null;
    if (feed.getEntries().size() > 0) {
        ssEntry = feed.getEntries().get(0)
    }
    // println "findAllSpreadsheetByName out"
    return ssEntry
}

WorksheetEntry findWorksheetByName(SpreadsheetService service, SpreadsheetEntry ssEntry, String sheetName) {
    // println "findWorksheetByName in"
    WorksheetQuery worksheetQuery = new WorksheetQuery(ssEntry.getWorksheetFeedUrl())
    worksheetQuery.setTitleQuery(sheetName)
    WorksheetFeed feed = service.query(worksheetQuery, WorksheetFeed.class)
    WorksheetEntry wsEntry = null
    if (feed.getEntries().size() > 0) {
        wsEntry = feed.getEntries().get(0)
    }
    // println "findWorksheetByName out"
    return wsEntry
}

def findAllElements(SpreadsheetService service, WorksheetEntry wsEntry) {
    URL listFeedUrl = wsEntry.getListFeedUrl()
    ListFeed listFeed = service.getFeed(listFeedUrl, ListFeed.class)
    for (ListEntry row : listFeed.getEntries()) {
        print "title:${row.getTitle().getPlainText()}\t"
        for (String tag: row.getCustomElements().getTags()) {
            print "${tag}:${row.getCustomElements().getValue(tag)}\t"
        }
        println()
    }
}


//main
def service = getService()

// findAllSpreadsheets(service)
def ssEntry = findAllSpreadsheetByName(service, "テストスプレッドシート")
// println ssEntry.getTitle().getPlainText()
def wsEntry = findWorksheetByName(service, ssEntry, "テストシート")
// println wsEntry.getTitle().getPlainText()

findAllElements(service, wsEntry)
