// snippet of code to generate values for a dynamic dropdown getting the data from a sharepoint API

// ____________________ Set COMPANY CODE dropdown menu ____________________
function getCompanyCodes() {
  executeJson("https://bauer.sharepoint.com/sites/MDS/_api/contextinfo", "POST")
    .done(function (requestDigestValue) {
      const requestHeaders = {
        "X-RequestDigest": requestDigestValue.d.GetContextWebInformation.FormDigestValue,
      };
      const caml = "<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>CompanyCodes</Value></Eq></Where>";
      getListItems("https://bauer.sharepoint.com/sites/MDS", "MDSData", caml, requestHeaders)
        .done(function (data) {
          if (data && data.d && data.d.results && data.d.results.length > 0) {
            onGetCompCodesSucceeded(data.d.results);
          } else {
            onGetCompCodesFailed();
          }
        })
        .fail(function (error) {
          console.log(JSON.stringify(error));
        });
    })
    .fail(function (error) {
      console.log(JSON.stringify(error));
    });
}
function onGetCompCodesSucceeded(data) {
  const jsonCompCodes = JSON.parse(data[0].JSON);
  //create an array with all the company codes
  const compCodesArr = jsonCompCodes.map((elem) => elem.CompanyCode);

  // remove duplicates
  const uniqueCompCodes = [...new Set(compCodesArr)];
  const len = uniqueCompCodes.length;
  const compCodeInput = document.getElementById("company-code");
  const compCodeOptions = document.querySelectorAll("#company-code option");

  // append those comp codes to the related dropdown
  for (i = 0; i < len; i++) {
    const option = document.createElement("option");
    // set the value attribute to the ID of each list item
    option.value = option.text = uniqueCompCodes[i];
    compCodeInput.add(option);
  }

  sortOptions(compCodeOptions);

  return console.log("Company codes generated successfully..");
}
function onGetCompCodesFailed(sender, args) {
  alert("Failed to retrieve company codes data " + args.get_message() + "\n" + args.get_stackTrace());
}

// doing the request

function executeJson(url, method, headers, payload) {
  method = method || "GET";
  headers = headers || {};
  headers["Accept"] = "application/json;odata=verbose";
  if (method == "POST" && headers["X-RequestDigest"] === undefined) {
    headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
  }
  var ajaxOptions = {
    url: url,
    type: method,
    contentType: "application/json;odata=verbose",
    headers: headers,
  };
  if (typeof payload != "undefined") {
    ajaxOptions.data = JSON.stringify(payload);
  }
  return $.ajax(ajaxOptions);
}
