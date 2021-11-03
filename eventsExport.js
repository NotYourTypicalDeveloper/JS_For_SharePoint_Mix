// MCN EVENTS
//::::::::::::::::::::::  DOCUMENT READY STARTS ::::::::::::::::::::::
$(document).ready(function () {
    console.log('eventExport.js is connected');
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', Getdata);

});

// added this function because sp.js wouldn't load
function Getdata() {
    try {
        clientContext = new SP.ClientContext.get_current();
    } catch (err) {
        alert(err);
    }
}


// _-_-_-_-_-_-_ Export to Word _-_-_-_-_-_-_
function export2Word() {
    alert('Export Word clicked!');
    // grab values
    getSearchValues();

    // if no search is performed before pressing the button throw error
    if (searchInputVal === "" && eventStartDateVal === "" && eventEndDateVal === "" && catDispVal === "" && regionDispVal === "" && countyDispVal === "" && countyDispVal === "" && venueDispVal === "" && RIDispVal === "" && CIDispVal === "") {
        alert('Please perform a search before trying to export to PDF');
    } else {
        // if all OK, then do query and creates Excel doc
        retrieveListItems();
    }
}

// ::::::::::::::::::::::: Get all the current values of the search :::::::::::::::::::::::
function getSearchValues() {

    //get the values of the search input and the refiners
    try {
        searchInputVal = $("input[placeholder='Enter your search terms...']").val();
    } catch (error) {
        searchInputVal = ''; // set it to empty string to avoid 'undefined' errors later
        console.log('no value in the search box');
    }

    // Event date values
    try {
        eventStartDateVal = $('input[id*="DatePicker"]:first').val();
        var date = new Date(eventStartDateVal);
        eventStartDateFmt = date.toISOString().split("T")[0] + 'T00:01:00Z'; // convert it to a CAML friendly format
    } catch (error) {
        eventStartDateVal = '';
        console.log('no start date');
    }

    try {
        eventEndDateVal = $('input[id*="DatePicker"]:last').val();
        var date2 = new Date(eventEndDateVal);
        eventEndDateFmt = date2.toISOString();
    } catch (error) {
        eventEndDateVal = '';
        console.log('no end date');
    }

    // Category value
    try {
        catDispVal = $("div.ellipsis_e61d9da1:contains('Category')").text().slice(11, -2);
    } catch (error) {
        catDispVal = '';
        console.log('no category refiner value');
    }

    // Region value
    try {
        regionDispVal = $("div.ellipsis_e61d9da1:contains('Region')").text().slice(9, -2);
    } catch (error) {
        regionDispVal = '';
        console.log('no region refiner value');
    }

    // County value
    try {
        countyDispVal = $("div.ellipsis_e61d9da1:contains('County')").text().slice(9, -2);
    } catch (error) {
        countyDispVal = '';
        console.log('no county refiner value');
    }

    // Venue value
    try {
        venueDispVal = $("div.ellipsis_e61d9da1:contains('Venue')").text().slice(8, -2);
    } catch (error) {
        venueDispVal = '';
        console.log('no venue refiner value');
    }

    // Recurring info
    try {
        RIDispVal = $("div.ellipsis_e61d9da1:contains('Recurring')").text().slice(17, -2);
    } catch (error) {
        RIDispVal = '';
        console.log('no Recurring info refiner value');
    }

    //Contact info value
    try {
        CIDispVal = $("div.ellipsis_e61d9da1:contains('Contact')").text().slice(15, -2);
    } catch (error) {
        CIDispVal = '';
        console.log('no Contact info value');
    }

}


// :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
// ::::::::::::::::::::  Query to retrieve the list items ::::::::::::::::::::: 

var clientContext;
var website;
var collListem;

function retrieveListItems() {
    //returns the context information about the current web application
    var clientContext = new SP.ClientContext('https://bauer.sharepoint.com/sites/UK-Article-Library');
    oList = clientContext.get_web().get_lists().getByTitle('MCNEvents2');

    var camlQuery = new SP.CamlQuery();

    var filtersList = [];
    var camlSyntax = '';

    // Individual CAML queries for each filter, to add up
    // will search the term in all the columns
    if (searchInputVal !== '') {
        filtersList.push(
            "<Or><Contains> <FieldRef Name='Title' /> <Value Type='Text'>" + searchInputVal + "</Value> </Contains>" +
            //"<Or> <Eq> <FieldRef Name='Category' /> <Value Type='Choice'>" + searchInputVal + "</Value> </Eq>" + 
            "<Or> <Contains> <FieldRef Name='Region' /> <Value Type='Text'>" + searchInputVal + "</Value> </Contains>" +
            "<Or> <Contains> <FieldRef Name='Venue' /> <Value Type='Text'>" + searchInputVal + "</Value> </Contains>" +
            "<Or> <Contains> <FieldRef Name='County' /> <Value Type='Text'>" + searchInputVal + "</Value> </Contains>" +
            "<Or> <Contains> <FieldRef Name='RecurringInfo' /> <Value Type='Text'>" + searchInputVal + "</Value> </Contains>" +
            "<Contains> <FieldRef Name='ContactInfo' /> <Value Type='Text'>" + searchInputVal + "</Value> </Contains>" +
            "</Or> </Or> </Or> </Or> </Or>"
            //</Or>
        )

    }

    // search in Event date column
    if (eventStartDateVal !== "") {
        filtersList.push("<Geq><FieldRef Name='EventDate'/><Value Type='DateTime' IncludeTimeValue='TRUE'>" + eventStartDateFmt + "</Value> </Geq>")
    }

    if (eventEndDateVal !== "") {
        filtersList.push("<Leq> <FieldRef Name='EventDate'/><Value Type='DateTime' IncludeTimeValue='TRUE'>" + eventEndDateFmt + "</Value> </Leq>")
    }

    // search in Category 
    //if(catDispVal !== '') {
    // filtersList.push("<Eq><FieldRef Name='Category'/> <Value Type='Choice'>" + catDispVal + "</Value></Eq>")
    //}

    // search in Region 
    if (regionDispVal !== '') {
        filtersList.push("<Contains><FieldRef Name='Region'/> <Value Type='Text'>" + regionDispVal + "</Value></Contains>")
    }

    // search in County 
    if (countyDispVal !== '') {
        filtersList.push("<Contains><FieldRef Name='County'/> <Value Type='Text'>" + countyDispVal + "</Value></Contains>")
    }

    // search in Venue 
    if (venueDispVal !== '') {
        filtersList.push("<Contains><FieldRef Name='Venue'/> <Value Type='Text'>" + venueDispVal + "</Value></Contains>")
    }

    // search in Recurring Info 
    if (RIDispVal !== '') {
        filtersList.push("<Contains><FieldRef Name='RecurringInfo'/> <Value Type='Text'>" + RIDispVal + "</Value></Contains>")
    }

    // search in Contact Info
    if (CIDispVal !== '') {
        filtersList.push("<Contains><FieldRef Name='ContactInfo'/> <Value Type='Text'>" + CIDispVal + "</Value></Contains>")
    }

    console.log('filtersList contains: ', filtersList);

    // if CAML query has 1 search term
    if (filtersList.length === 1) {
        camlSyntax = "<Where>" + filtersList[0] + "</Where>";
    }
    //  if CAML query has 2 search terms
    else if (filtersList.length === 2) {
        camlSyntax = "<Where><And>" + filtersList[0] + filtersList[1] + "</And></Where>";
    }
    // if CAML query has 3 search terms
    else if (filtersList.length === 3) {
        camlSyntax = "<Where><And>" + filtersList[0] + "<And>" + filtersList[1] + filtersList[2] + "</And></And></Where>";
    }

    // if CAML query has 4 search terms
    if (filtersList.length === 4) {
        query.Query = "<Where><And>" + filtersList[0] + "<And>" + filtersList[1] + "<And>" + filtersList[2] + filtersList[3] + "</And></And></And></Where>";
    }

    // if CAML query has 5 search terms
    if (filtersList.length === 5) {
        query.Query = "<Where><And>" + filtersList[0] + "<And>" + filtersList[1] + "<And>" + filtersList[2] + "<And>" + filtersList[3] + filtersList[4] + "</And></And></And></And></Where>";
    }

    // if CAML query has 6 search terms
    if (filtersList.length === 6) {
        query.Query = "<Where><And>" + filtersList[0] + "<And>" + filtersList[1] + "<And>" + filtersList[2] + "<And>" + filtersList[3] + "<And>" + filtersList[4] + filtersList[5] + "</And></And></And></And></And></Where>";
    }

    // if CAML query has 7 search terms
    if (filtersList.length === 7) {
        query.Query = "<Where><And>" + filtersList[0] + "<And>" + filtersList[1] + "<And>" + filtersList[2] + "<And>" + filtersList[3] + "<And>" + filtersList[4] + "<And>" + filtersList[5] + filtersList[6] + "</And></And></And></And></And></And></Where>";
    }

    // if CAML query has 8 search terms
    if (filtersList.length === 8) {
        query.Query = "<Where><And>" + filtersList[0] + "<And>" + filtersList[1] + "<And>" + filtersList[2] + "<And>" + filtersList[3] + "<And>" + filtersList[4] + "<And>" + filtersList[5] + "<And>" + filtersList[6] + filtersList[7] + "</And></And></And></And></And></And></And></Where>";
    }

    // Set CAML query
    camlQuery.set_viewXml("<View><Query>" + camlSyntax + "<OrderBy><FieldRef Name='EventDate' Ascending='TRUE'/><FieldRef Name='Category' Ascending='TRUE'/><FieldRef Name='Region' Ascending='TRUE'/></OrderBy><RowLimit>5000</RowLimit></Query></View>");

    // retrieve the items from the CAML query
    this.collListItem = oList.getItems(camlQuery);

    clientContext.load(collListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, this.onExportQuerySucceeded),
        Function.createDelegate(this, this.onExportQueryFailed)
    );

}

// if query succeeds
function onExportQuerySucceeded(sender, args) {
    var listItemInfo = '';
    var listItemEnumerator = collListItem.getEnumerator();
    var currentCategory;
    var prevCategory;
    var currentDate;
    var prevDate;
    var pd;
    var cd;
    var count = 0;
    var array = [];

    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
        count++;

        currentCategory = oListItem.get_item('Category');
        currentDate = oListItem.get_item('EventDate');
        currentTime = oListItem.get_item('EventTime');
        title = oListItem.get_item('Title');
        venue = oListItem.get_item('Venue');
        county = oListItem.get_item('County');
        country = oListItem.get_item('Country');
        price = oListItem.get_item('Price');
        contact = oListItem.get_item('ContactInfo');

        var region = oListItem.get_item('Region');


        // if current date event is same as the previous event date, print nothing
        try {
            pd = prevDate.getDate();
        } catch (ex) {
            pd = 0;
        }
        try {
            cd = currentDate.getDate();
        } catch (ex) {
            cd = 0;
        }
        if (cd === pd) {
            listItemInfo += '\n';
        }

        // if new date, print 
        else {
            var dateFormatted = oListItem.get_item('EventDate').format('dddd dd MMMM yyyy');
            listItemInfo += '\n' + '\n' + '\n' + dateFormatted + '\n';
            prevDate = currentDate;
        }

        // check if the category of the item is the same as the previous category, print nothing
        if (currentCategory === prevCategory) {
            listItemInfo += '\n';
        }
        // if new category, print
        else {
            listItemInfo += '\n' + '\n' + oListItem.get_item('Category') + '\n =============================================';
            prevCategory = currentCategory;
        }

        //___________ print the rest of data________
        listItemInfo += '\n' + oListItem.get_item('Title') + '\n';

        // add different signs before REGION depending on which region the event is
        if (region != null && region.toUpperCase() === 'SOUTH') {
            listItemInfo += '\n$';
        }

        if (region != null && region.toUpperCase() === 'MIDLANDS') {
            listItemInfo += '\n%';
        }

        if ((region != null && region.toUpperCase() === 'NORTH') || (region != null && region.toUpperCase() === 'SCOTLAND')) {
            listItemInfo += '\n*';
        }

        if ((region != null && region.toUpperCase() === 'SOUTH EAST') || (region != null && region.toUpperCase() === 'SOUTH WEST') || (region != null && region.toUpperCase() === 'EAST')) {
            listItemInfo += '\n^';
        }

        if (title != null) {
            listItemInfo += ' ' + title;
        }

        if (venue != null) {
            listItemInfo += ' ' + venue;
        }

        if (county != null) {
            listItemInfo += ' ' + county;
        }

        if (country != null) {
            listItemInfo += ' ' + country;
        }

        if (price != null) {
            listItemInfo += ' ' + price;
        }

        if (currentTime != null) {
            listItemInfo += ' ' + currentTime;
        }

        if (contact != null) {
            listItemInfo += ' ' + contact;
        }

        // add a separator
        listItemInfo += '\n--------------------------------';

    };

    listItemInfo = listItemInfo.replace(/ï¿½/g, '£'); // clean special unknown charac and replace with £ sign

    console.log('retrieving ' + count + ' events...' + listItemInfo.toString());
    createWordDoc(listItemInfo); // create the word document with the generated string
}


function onExportQueryFailed(sender, args) {
    alert('Export query failed.' + args.get_message() + '\n' + args.get_stackTrace());
}


// Export the events list to a word document (will be called in the exportEvents function)
function createWordDoc(text) {
    var link, blob, url;
    blob = new Blob(['\ufeff', text], {
        type: 'application/msword'
    });
    /* url = URL.createObjectURL(blob);
    link = document.createElement('A');
    link.href = url.replace('https://bauertest.sharepoint.com/sites/NewSandbox/', '');
    link.download = 'MCN_Events_report'; //default name without extension
    document.body.appendChild(link);
    if (navigator.msSaveOrOpenBlob) navigator.msSaveOrOpenBlob(blob, 'MCN_Events_report.docx'); //IE10-11
    else link.click(); //otherbrowsers
    document.body.removeChild(link); */

    saveAs(blob, "MCNEVENTS.doc");
};