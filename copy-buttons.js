// :::::::::::::::::::::::::::::: MCN Display Form ::::::::::::::::::::::::::::::

$(document).ready(function () {
  console.log("UAT MCNEventsDisplayForm.js is running!");
});

//================ Create a copy of the current displayed item ==========================
function copyEvent() {
  // selects the values of all fields of the item
  var titleVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(1) > td.ms-formbody").innerText;
  var eventDateVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(2) > td.ms-formbody").innerText;
  var eventTimeVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(3) > td.ms-formbody").innerText;
  var categoryVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(4) > td.ms-formbody").innerText;
  var venueVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(5) > td.ms-formbody").innerText;
  var priceVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(6) > td.ms-formbody").innerText;
  var contactVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(7) > td.ms-formbody").innerText;
  var recurVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(8) > td.ms-formbody").innerText;
  var townVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(9) > td.ms-formbody").innerText;
  var countyVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(10) > td.ms-formbody").innerText;
  var regionVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(11) > td.ms-formbody").innerText;
  var countryVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(12) > td.ms-formbody").innerText;
  var infoVal = document.querySelector("#part1 > table.ms-formtable > tbody > tr:nth-child(13) > td.ms-formbody").innerText;

  // returns the context information about the current web application
  var siteUrl = "https://bauer.sharepoint.com/sites/Sandbox/UK-Article-Library";
  var clientContext = new SP.ClientContext(siteUrl);

  var oList = clientContext.get_web().get_lists().getByTitle("MCN Events"); // display name of the list

  var itemCreateInfo = new SP.ListItemCreationInformation();

  //creates and sets the new item with the copied values________________
  this.oListItem = oList.addItem(itemCreateInfo);

  oListItem.set_item("Title", titleVal);

  if (!!eventDateVal) {
    // check which button is on focus
    var add1DayFocused = document.activeElement === document.querySelector("#copy-item-btn1");
    var add1WeekFocused = document.activeElement === document.querySelector("#copy-item-btn2");
    var add1MonthFocused = document.activeElement === document.querySelector("#copy-item-btn3");
    var add4WeeksFocused = document.activeElement === document.querySelector("#copy-item-btn4");

    var newDate;

    var dateParts = eventDateVal.split("/");
    var dateObj = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);

    // if "+1 day" or "+1week", etc .. button is clicked, add x, y, z number of days/week/month to the current date
    if (add1DayFocused) {
      newDate = moment(dateObj).add(1, "d");
    }
    if (add1WeekFocused) {
      newDate = moment(dateObj).add(1, "w");
    }
    if (add1MonthFocused) {
      newDate = moment(dateObj).add(1, "M");
    }
    if (add4WeeksFocused) {
      newDate = moment(dateObj).add(4, "w");
    }

    dateObj = new Date(newDate); // re-convert it to another object

    oListItem.set_item("EventDate", dateObj);
  }

  if (!!eventTimeVal) {
    oListItem.set_item("EventTime", eventTimeVal);
  }
  if (!!categoryVal) {
    oListItem.set_item("Category", categoryVal);
  }
  if (!!venueVal) {
    oListItem.set_item("Venue", venueVal);
  }
  if (!!priceVal) {
    oListItem.set_item("Price", priceVal);
  }
  if (!!contactVal) {
    oListItem.set_item("ContactInfo", contactVal);
  }
  if (!!recurVal) {
    oListItem.set_item("RecurringInfo", recurVal);
  }
  if (!!townVal) {
    oListItem.set_item("TownArea", townVal);
  }
  if (!!countyVal) {
    oListItem.set_item("County", countyVal);
  }
  if (!!regionVal) {
    oListItem.set_item("Region", regionVal);
  }
  if (!!countryVal) {
    oListItem.set_item("Country", countryVal);
  }
  if (!!infoVal) {
    oListItem.set_item("Info", infoVal);
  }

  // updating the item_____________________
  oListItem.update();

  clientContext.load(oListItem);

  clientContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded), Function.createDelegate(this, this.onQueryFailed));
}
function onQuerySucceeded() {
  var itemID = oListItem.get_id();

  alert("Copy of the item created: " + itemID);
  // redirect to the list (use internal name)
  window.location.href = "https://bauer.sharepoint.com/sites/Sandbox/UK-Article-Library/Lists/MCNEvents2/DispForm.aspx?ID=" + itemID;
}

function onQueryFailed(sender, args) {
  alert("Request failed. " + args.get_message() + "\n" + args.get_stackTrace());
}
