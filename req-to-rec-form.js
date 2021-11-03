// check current URL parameters values
function getQueryVariable(variable) {
  var query = window.location.search.substring(1);
  var vars = query.split('&');
  for (var i = 0; i < vars.length; i++) {
    var pair = vars[i].split('=');
    if (pair[0] == variable) {
      return pair[1];
    }
  }
  return (false);
}
userName = ''
userID = ''
admin = '';
creatingRequest = false;
/////////////////////////////////////////// Document ready starts ///////////////////////////////////////////


$(document).ready(function() {
  console.log("req-to-rec-form.js is connected!");

  console.log('getting current user..');
  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit");
  var web = clientContext.get_web();
  var user = web.get_currentUser(); //must load this to access info.
  clientContext.load(user);
  clientContext.executeQueryAsync(function() {
    //alert("User is: " + user.get_email()); //there is also id, email, so this is pretty useful.
    userName = user.get_title();
    userID = user.get_id();
    userEmail = user.get_email();
  }, function() {
    alert(":(");
  });

  var url = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/_api/web/currentuser/groups';
  $.getJSON(url, function(data) {
    $.each(data.value, function(key, value) {
      if ((value.Title == 'Request to Recruit Owners') || (value.Title == 'Request to Recruit Members')) {
        admin = true;
      }
    });
  })


  // ::::::::::::::::::::::::::::::: VARIABLES ::::::::::::::::::::::::::::::::::::: 
  // Input selectors
  $divisionInput = $('#division-input');
  $regionInput = $('#region-input');
  $areaInput = $('#area-input');
  $recManagerInput = $('#rec-manager-input');
  //$lineManagerInput = $('#line-manager-input');

  $jobRoleInput = $('#job-role-input');
  $locationInput = $('#location-input');
  $baseInput = $('#base-input');
  $startDateInput = $('#start-date-input');
  $existingRoleInput = $('#exist-role-input');
  $methodInput = $('#method-input');
  $specMORInput = $('#specMOR-input');
  //$FPTimeInput = $('#FP-time-input');
  $permTempInput = $('#perm-temp-input');
  $maternityInput = $('#maternity-input');
  $endDateInput = $('#end-date-input');
  $budgCostInput = $('#budg-cost-input');
  
  $costCodeInput = $('#cost-code-input');
  $usingAgencyInput = $('#using-agency-input');
  $agencyRateInput = $('#agency-rate-input');
  
  $bizCaseInput = $('#bizcase-input');
  $fileInput = $('#fileinput'); // job description upload

  $leaverNameInput = $('#leaver-name-input');
  $leaverJobInput = $('#leaver-JT-input');
  $leaverSalaryInput = $('#leaver-salary-input');
  $leaverReasonInput = $('#leaving-reason-input');

  $salaryInput = $('#annual-salary-input');
  $pensionInput = $('#pension-input');
  $commissionInput = $('#commission-input');
  $bonusInput = $('#bonus-input');
  $bonusPercentInput = $('#bonus-percent-input');
  $carInput = $('#car-input');
  $carAllowInput = $('#car-allow-input');
  $addRenumInput = $('#add-renum-input');
  //$medicalInput = $('#medical-input');
  $runtimeInput = $('#runtime-input');
  $specADInput = $('#specAD-input');
  $flexibleWorkingInput = $('#flexibleWorking-input');
  $shortlistingQuestion1Input = $('#shortlistingQuestion1-input');
  $shortlistingQuestion2Input = $('#shortlistingQuestion2-input');
  $shortlistingQuestion3Input = $('#shortlistingQuestion3-input');
  $shortlistingQuestion4Input = $('#shortlistingQuestion4-input');
  $ourTeamInput = $('#our-team-input');
  $roleFocusInput = $('#role-focus-input');
  $whatYouDoInput = $('#what-you-do-input');
  $whatYouBringInput = $('#what-you-bring-input');
  $howApplyInput = $('#how-apply-input');
  $itwDeadlineInput = $('#itw_deadline-input');
  $closingDateInput = $('#appli-closing-input');


  // sections
  $materSection = $('#mater-wrapper');
  $endDateSection = $('#enddate-wrapper');
  $leaverSection = $('#leaver-section');
  $bonusPercentSection = $('#bonus-percent-wrapper');
  $confirmMsg = $('#thank-you-msg');
  $mainForm = $('#recruit-form');
  $spinner = $('div.spinner-border');

  
  R2RID = getQueryVariable('r2rid');
  stage = getQueryVariable('stage');
  recruitingManagerEmail = '';
  //lineManagerEmail = '';
  setDivisions();

  if (!(stage)) {

    // validation jQuery
    $mainForm.validate({
      rules: {
        division: "required",
        region: "required",
        area: "required",
        recmanager: "required",
        //linemanager: "required",
        jobrole: "required",
        locationF: "required",
        baseF: "required",
        startdate: "required",
        existingrole: "required",
        //recmethod: "required",
        //FPTime: "required",
        permtemp: "required",
        bizcase: "required",
        fileinput: "required",
        leavername: "required",
        leaverjob: "required",
        leaversalary: "required",
        reason: "required",
        annualsalary: "required",
        companycar: "required",
        //medicalinsurance: "required",
        runtime: "required",
        flexibleWorking: "required",
       // shortlistingQuestion1: "required",
      //  shortlistingQuestion2: "required",
      //  shortlistingQuestion3: "required",
      //  shortlistingQuestion4: "required",
        ourteam: "required",
        rolefocus: "required",
        whatyoudo: "required",
        whatyoubring: "required",
        itw_deadline: "required",
        closingdate: "required",
			specMOR: "required",
			specAD: "required",
        costcode: "required",
        usingagency: "required",
        agencyrate: "required",
      },

      messages: {
        division: "Please Select a Division.",
        region: "Please Select a Region.",
        area: "Please Select an Area",
        recmanager: "Please enter a recruiting manager name",
        //linemanager: "Please Select a Management director",
        jobrole: "Please enter a job title",
        locationF: "Please Select a location",
        baseF: "Please Select where this role is based",
        startdate: "Please Select a date in the calendar",
        existingrole: "Please confirm if the hire type",
        //recmethod: "Please Select a recruitment method",
        //FPTime: "Please confirm if the contract is full or part time",
        permtemp: "Please confirm if the contract is permanent or temporary",
        bizcase: "Please provide a business case",
        fileinput: "Please attach a job description",
        leavername: "Please enter the name of the person who left",
        leaverjob: "Please enter the leaver's job title",
        leaversalary: "Please enter the leaver's current salary",
        reason: "Please enter a reason for leaving",
        annualsalary: "Please enter the annual salary offered for the position",
        //medicalinsurance: "Please Select the type of medical insurance",
        runtime: "Please Select the advert runtime",
        flexibleWorking: "Please Select if this role allows flexible working",
      //  shortlistingQuestion1: "Please enter a shortlisting question",
      //  shortlistingQuestion2: "Please enter a shortlisting question",
      //  shortlistingQuestion3: "Please enter a shortlisting question",
      //  shortlistingQuestion4: "Please enter a shortlisting question",
        ourteam: "Please enter a text about your team",
        rolefocus: "Please describe the role",
        whatyoudo: "Please describe the duties",
        whatyoubring: "Please enter some text",
		itw_deadline: "Please select a decision date.",
		closingdate: "Please select a closing date",
			specMOR: "Please provide some details about the external recuitment",
			specAD: "Please provide some details about the advert runtime",
        costcode: "Please enter the cost code",
        usingagency: "please specify if you are using an agency",
        agencyrate: "Please speify the Agency rate",
      },

      submitHandler: function() {
        checkRecMethod();
        if (checkRecMethod() != false) {
          submitForm();
        }
      },

      invalidHandler: function(event, validator) {
        var errorDiv = $('#error-msg');
        var errors = validator.numberOfInvalids();
        if (errors) {
              errorDiv.html('PLEASE CHECK THE ' + errors + ' ERROR(S) HIGHLIGHTED IN RED').show();
        } else {
          errorDiv.hide();
        }
      }
    });
    // load regions on division change
    $divisionInput.change(function() {
      setRegions();
    });

    // load areas on region change
    $regionInput.change(function() {
      setAreas();
    });

    hidePageLoadingSpinner();


  }

  // enable Bootstrap tooltips
  $('[data-toggle="tooltip"]').tooltip({
    placement: 'bottom'
  });

  // set typeahead for Rec  
  setTypeAhead();
  
  /* set dropdown menu for Managing Directors -NOT NEEDED ANYMORE
  setMDDropdown(); */

  //startCascade();

  //READ ONLY / View
  if (stage === "view") {

    setRegionsAll();
    setAreasAll();
    setTimeout(function() {
      populateForm(R2RID);
      hidePageLoadingSpinner();

    }, 1000);

    // DISABLE/READONLY all inputs -once all the info is retrieved to prevent Approval Manager from modifying
    disableFields();
    //hide submit & save draft buttons
    hideButtons();
    $('#fileinput-label-wrapper').hide();
  }

  //Edit
  if (stage === "edit") {

    setRegionsAll();
    setAreasAll();

    setTimeout(function() {
      populateForm(R2RID);
    }, 1000);

    setTimeout(function() {
      // load regions on division change
      $divisionInput.change(function() {
        setRegions();
      });

      // load areas on region change
      $regionInput.change(function() {
        setAreas();
      });

      hidePageLoadingSpinner();

    }, 4000);
  }

  //Manager Approval
  if (stage === "MA") {

    setRegionsAll();
    setAreasAll();
    setTimeout(function() {
      populateForm(R2RID);
    }, 1000);

    // DISABLE/READONLY all inputs -once all the info is retrieved to prevent Approval Manager from modifying
    disableFields();

    //hide submit & save draft buttons
    hideButtons();
    $('#fileinput-label-wrapper').hide();

    //show approve / reject buttons for manager
    $('#approval-banner, #man-label, #man-comments, #man-approve, #man-reject, #ba-label, #ba-comments').show();
    document.getElementById('approval-banner').innerHTML = "Managing Director Approval";
    setTimeout(function() {
      $('#man-comments, #man-approve, #man-reject').removeAttr("disabled");
      hidePageLoadingSpinner();
    }, 5000);

  }


  if (stage === "BA") {

    setRegionsAll();
    setAreasAll();
    setTimeout(function() {
      populateForm(R2RID);
    }, 1000);

    // DISABLE/READONLY all inputs -once all the info is retrieved to prevent Approval Manager from modifying
    disableFields();

    //hide submit & save draft buttons
    hideButtons();
    $('#fileinput-label-wrapper').hide();

    //show approve / reject buttons for ba
    document.getElementById('approval-banner').innerHTML = "BA Approval";
    $('#approval-banner, #ba-label, #ba-comments, #ba-approve, #ba-reject').show();

    setTimeout(function() {
      $('#ba-comments, #ba-approve, #ba-reject').removeAttr("disabled");
      hidePageLoadingSpinner();
    }, 5000);

  }
  
  
  if (stage === "FD") {

    setRegionsAll();
    setAreasAll();
    setTimeout(function() {
      populateForm(R2RID);
    }, 1000);

    // DISABLE/READONLY all inputs -once all the info is retrieved to prevent Approval Manager from modifying
    disableFields();

    //hide submit & save draft buttons
    hideButtons();
    $('#fileinput-label-wrapper').hide();

    //show approve / reject buttons for fd
    document.getElementById('approval-banner').innerHTML = "FD Approval";
    $('#approval-banner, #man-label, #man-comments, #ba-label, #ba-comments,  #fd-label, #fd-comments, #fd-approve, #fd-reject').show();

    setTimeout(function() {
      $('#fd-comments, #fd-approve, #fd-reject').removeAttr("disabled");
      hidePageLoadingSpinner();
    }, 5000);

  }
  
  
  if (stage === "RA") {

    setRegionsAll();
    setAreasAll();
    setTimeout(function() {
      populateForm(R2RID);
    }, 1000);

    // DISABLE/READONLY all inputs -once all the info is retrieved to prevent Approval Manager from modifying
    disableFields();

    //hide submit & save draft buttons
    hideButtons();
    $('#fileinput-label-wrapper').hide();

    //show approve / reject buttons for fd
    document.getElementById('approval-banner').innerHTML = "Radio Team Approval";
	document.getElementById('fd-label').innerHTML = "Radio Team Comments";
    $('#approval-banner, #man-label, #man-comments, #ba-label, #ba-comments,  #fd-label, #fd-comments, #fd-approve, #fd-reject').show();

    setTimeout(function() {
      $('#fd-comments, #fd-approve, #fd-reject').removeAttr("disabled");
      hidePageLoadingSpinner();
    }, 5000);

  }
  
    if (stage === "TM") {

    setRegionsAll();
    setAreasAll();
    setTimeout(function() {
      populateForm(R2RID);
    }, 1000);

    // DISABLE/READONLY all inputs -once all the info is retrieved to prevent Approval Manager from modifying
    //disableFields();

    //hide submit & save draft buttons
    hideButtons();
    $('#fileinput-label-wrapper').hide();

    //show approve / reject buttons for tm
    document.getElementById('approval-banner').innerHTML = "TM Approval";
    $('#approval-banner, #man-label, #man-comments, #ba-label, #ba-comments,  #fd-label, #fd-comments, #tm-label, #tm-comments, #tm-approve, #tm-reject').show();

    setTimeout(function() {
      $('#tm-comments, #tm-approve, #tm-reject').removeAttr("disabled");
      hidePageLoadingSpinner();
    }, 5000);

  }
  
  
  // _____________-_-_EVENT LISTENERS_-_-______________



  // toggle "Specify Advert Runtime" question
  $runtimeInput.change(function() {
    selectedOption = $('#runtime-input option:selected');
    if (selectedOption[0].innerText.indexOf("Other")>-1) {
		$('#specADdiv').show();
	}
	else { $('#specADdiv').hide(); }
  });


  // "salary input", calculate "pension" field
  $salaryInput.change(function() {
    calculatePension();
  });

  // On change event on "HIRE TYPE", toggle "Leaver" section when hire type 'Leaver Replacement', toggle the "End date" field if "Maternity/Project Cover/Secondment/Intern" are selected
  $existingRoleInput.change(function() {
    selectedOption = $('#exist-role-input option:selected');
    selectedOptionTxt = $('#exist-role-input option:selected').text();
    toggleSections($leaverSection, selectedOption, "Leaver Replacement");
    //toggleSections($endDateSection, selectedOption, "Maternity Cover");
    if ((selectedOptionTxt === "Maternity Cover") || (selectedOptionTxt === "Project Cover") || (selectedOptionTxt === "Secondment") || (selectedOptionTxt === "Intern")) {
    	$endDateSection.show();    
    } else {
    	$endDateSection.hide();    
    }
  });

  // toggle "Maternity cover" & "End date" section
  $permTempInput.change(function() {
    selectedOption = $('#perm-temp-input option:selected');
    //toggleSections($materSection, selectedOption, "Temporary");
    toggleSections($endDateSection, selectedOption, "Temporary");
  });

  // toggle "Bonus %" section, when Bonus is YES
  $bonusInput.change(function() {
    selectedOption = $('#bonus-input option:selected');
    toggleSections($bonusPercentSection, selectedOption, "Yes")
  });

  // round up to integer if user types decimal number - Budgeted recruitment + Annual salary input
  $budgCostInput.change(function() {
    roundINT($budgCostInput);
  });
  $salaryInput.change(function() {
    roundINT($salaryInput);
  });
  $leaverSalaryInput.change(function() {
    roundINT($leaverSalaryInput);
  });

  // maximum 2 decimals - Bonus % input
  $bonusPercentInput.change(function() {
    round2NumberDecimal($bonusPercentInput);
  });


  // maximum 2 decimals - Commission % input
  $commissionInput.change(function() {
    round2NumberDecimal($commissionInput);
  });

  // on change, display file name as inner text in the "upload" button
  $fileInput.change(function() {
    displayFileNames();
  });

  // on click, for DRAFT button
  $('#save-btn').click(function() {
    submitForm();
  })

  //multiple select checkboxes for method of recruitment
  $methodInput.dropdownchecklist({
    forceMultiple: true
  });

  if ((!(stage) || (stage === 'edit'))) {} else {
    $methodInput.prop("disabled", true);
    $methodInput.hide();
    document.getElementsByClassName("ui-dropdownchecklist")[0].style.display = "none"
  }

  // styling the tickbox dropdown
  document.querySelector("#ddcl-method-input > span > span").innerHTML = "Please Select...";
  $('#ddcl-method-input').addClass('form-control');
  $('.ui-dropdownchecklist-selector').removeAttr("style");


  

  // toggle "Specify Method of Recruitment cover" question
$('.ui-dropdownchecklist-text').on('DOMSubtreeModified',function(){
   selectedOptions = $('.ui-dropdownchecklist-text')[0].innerText;
    if (selectedOptions.indexOf("External")>-1) {
		$('#specMORdiv').show();
	}
	else { $('#specMORdiv').hide(); }
})
  
  
  // toggle "Agency Rate" question
  $usingAgencyInput.change(function() {
    selectedOption = $('#using-agency-input option:selected');
     if (selectedOption.val() == "Yes") {
		$('#agency-rate-wrapper').show();
	}
	else { $('#agency-rate-wrapper').hide(); }
  });

  //hide radio
  setTimeout(function() {
document.getElementById("division-input").remove(5);
  }, 3000);
});

///////////////////////////////////////////document ready ends ///////////////////////////////////////////


// hide SAVE/ SUBMIT buttons
function hideButtons() {
  $('#save-btn, #submit-btn').hide('slow');
}

// make fields READ-ONLY
function disableFields() {
  setTimeout(function() {
    $(':input').attr("disabled", true);
  }, 3000);
}

// hide Page Loading spinner
function hidePageLoadingSpinner() {
  $("div.spanner, div.overlay").css('visibility', 'hidden');
}


// calculate pension contribution
function calculatePension() {
  var salary = $salaryInput.val();
  if (!!salary && $.isNumeric(salary)) {
    var contribution = (salary * 0.03).toFixed(2);
    $pensionInput.val(contribution);
  }
}

// reusable function hide/show
function toggleSections(section, selection, res) {
  var choice = selection.text();
  return (choice === res ? section.show() : section.hide());
}


//// 365 Typeahead /////////////////////////////////////////////
function setTypeAhead() {
  console.log("setting typeahead...");
  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/mds/");
  var oList = clientContext.get_web().get_lists().getByTitle('MDSData');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query><Where><Eq><FieldRef Name=\'Title\'/>' +
    '<Value Type=\'Text\'>ADUsers</Value></Eq></Where></Query></View>'
  );
  collListItem = oList.getItems(camlQuery);
  clientContext.load(collListItem);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, typeAheadonQuerySucceeded),
    Function.createDelegate(this, typeAheadonQueryFailed)
  );
}

function typeAheadonQuerySucceeded(sender, args, targetField) {
  var listItemEnumerator = collListItem.getEnumerator();
  while (listItemEnumerator.moveNext()) {
    var oListItem = listItemEnumerator.get_current();
    employeesJSONString = oListItem.get_item('JSON');
    employeesJSON = JSON.parse(employeesJSONString);

    //Enable typeahead on Recruiting Manager name field
    EnableEmployeeTypeAhead($recManagerInput);

    console.log("Typeahead enabled successfully... yay");
  }
}

function typeAheadonQueryFailed(sender, args) {
  alert('Failed to get employees ' + args.get_message() + '\n' + args.get_stackTrace());
}

function EnableEmployeeTypeAhead(target) {
  var currentEmployeeID = '';
  var currentEmployeeName = '';

  // Constructing the suggestion engine
  employees = new Bloodhound({
    datumTokenizer: function(d) {
      return Bloodhound.tokenizers.whitespace(d.DisplayName);
    },
    queryTokenizer: Bloodhound.tokenizers.whitespace,
    local: employeesJSON
  });

  employees.initialize();

  // Initializing the employees typeahead
  target.typeahead(null, {
    display: 'DisplayName',
    source: employees,
    limit: 30,
    templates: {
      empty: [
        '<div style="background-color: #fdf9dd; font-weight: bold" class="empty-message">',
        'No Employees Found',
        '</div>'
      ].join('\n'),
      suggestion: Handlebars.compile('<div><strong>{{DisplayName}}</strong> – {{EmployeeID}}<br />{{EmailAddress}}</div>')
    }
  }).blur(function() {});

  target.on('typeahead:selected typeahead:autocompleted', function(obj, data) {
    target.val(data['DisplayName']);
    recruitingManagerEmail = data['EmailAddress'];
    managerString = data['Manager']; // grab name of the Manager
    var i = managerString.indexOf(',OU=');
    var managerFormatted = managerString.slice(3, i); // return name and firstname, remove extra junk
    var j = managerFormatted.lastIndexOf(' ');
    managerSurname = managerFormatted.substring(j + 1);
    managerForename = managerFormatted.substring(0, j);
    //lineManagerInput.val(managerSurname + ', ' + managerForename); // set field
    //lineManagerInput.typeahead('val',managerSurname + ', ' + managerForename).blur();
    //$lineManagerInput.val(managerSurname + ', ' + managerForename) //.focus().trigger('keyup')
  });
}


/* Set MD for "Managing Directors" dropdown - NOT NEEDED ANYMORE

function setMDDropdown() {
	console.log("populating Managing directors list...");
    var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit");
    MDList = clientContext.get_web().get_lists().getByTitle('Managing Directors');
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        '<View><Query></Query></View>'
    );
    this.allMDItems = MDList.getItems(camlQuery);
    clientContext.load(allMDItems);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, this.onSetMDDropdownSucceeded), 
        Function.createDelegate(this, this.onSetMDDropdownFailed)
    ); 
}

function onSetMDDropdownSucceeded(sender, args) {
	var eachMDName = "";
	var listItemEnumerator = allMDItems.getEnumerator();
	var MDInput = document.getElementById('line-manager-input');
	$("#line-manager-input").empty();

    while (listItemEnumerator.moveNext()) {
		var MDListItem = listItemEnumerator.get_current();
		eachMDName = MDListItem.get_item('MD');
		var option = document.createElement("option");
		option.value = eachMDName.get_email();
		option.text = eachMDName.get_lookupValue();

		MDInput.add(option);
	}
		
	var my_options = $("#line-manager-input option");
	my_options.sort(function(a, b) {
		if (a.text > b.text) return 1;
		else if (a.text < b.text) return -1;
		else return 0
	})
	
	$("#line-manager-input").empty().append(my_options);
	$("#line-manager-input").prepend("<option selected value=''>Please Select...</option>");				
		
	
}

function onSetMDDropdownFailed(sender, args) {
	alert('Managing directors request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}
*/

// Round number for numeric inputs - on key up
function roundINT(input) {
  var val = input.val();
  var valToInteger = parseInt(val).toFixed(0); // remove any decimals
  input.val(valToInteger);
  if (valToInteger === "NaN") {
    input.val(0);
  }
}

function round2NumberDecimal(input) {
  var val = input.val();
  var val2decimals = parseFloat(val).toFixed(2); // set input maximum 2 decimals
  input.val(val2decimals);
}

// Display name of uploaded files on Upload 
function displayFileNames() {
  var fileInputVal = $fileInput.val();
  var updatedStr = fileInputVal.replace("C:\\fakepath\\", "");
  $('#fileinput-label > span').html(updatedStr); // update upload btn inner text
}

// ============================ EDIT / VIEW mode - POPULATE FORM - retrieve the request info from 'Requests' list ==============================

// validation before submission for 'Method of recruitment' (doesn't work with the jQuery validation plugin)
function checkRecMethod() {
  var methodRecTxt = document.querySelector("#ddcl-method-input > span > span").innerHTML;
  // if no option is selected for 'Method of recruitment', display error message
  
  /*
  if ((methodRecTxt === "&nbsp;") || (methodRecTxt === "Please Select...")) {
    $('#rec-meth-error').text("Please select a recruitment method.");
    return false;
  } else {
    return true;
  }
  */
  return true;
}

//   POPULATE FORM 
function populateForm(itemID) {
  console.log("populating form...");

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/");
  targetList = clientContext.get_web().get_lists().getByTitle('Requests');
  targetListItem = targetList.getItemById(itemID);
  clientContext.load(targetListItem);

  clientContext.executeQueryAsync(Function.createDelegate(this, this.requestOnQuerySucceeded), Function.createDelegate(this, this.requestOnQueryFailed));
}

function requestOnQuerySucceeded(sender, args) {
  // RETRIEVE attachment file link of the list item
    console.log("getting attached file...");
  $().SPServices({
    operation: "GetAttachmentCollection",
    listName: "Requests",
    ID: R2RID,
    completefunc: function(xData, Status) {

      var output = "";

      $(xData.responseXML).find("Attachments > Attachment").each(function(i, el) {
        var $node = $(this),
          filePath = $node.text(),
          arrString = filePath.split("/"),
          fileName = arrString[arrString.length - 1];

        output += "<a href='" + filePath + "' target='_blank'>" + fileName + "</a><br />";
      });
      //   $('#fileinput-label-wrapper').hide();
      $('#fileName').append(output); // insert link and name of attachment

      if (output !== "") {

console.log("attachment found!");
        $('#fileinput-label-wrapper').hide();
        // validation jQuery
        $mainForm.validate({
          rules: {
            division: "required",
            region: "required",
            area: "required",
            recmanager: "required",
            //linemanager: "required",
            jobrole: "required",
            locationF: "required",
            baseF: "required",
            startdate: "required",
            existingrole: "required",
            //recmethod: "required",
            //FPTime: "required",
            permtemp: "required",
            enddate: "required",
            bizcase: "required",
            leavername: "required",
            leaverjob: "required",
            leaversalary: "required",
            reason: "required",
            annualsalary: "required",
            companycar: "required",
            //medicalinsurance: "required",
			runtime: "required",
			flexibleWorking: "required",
       // shortlistingQuestion1: "required",
       // shortlistingQuestion2: "required",
       // shortlistingQuestion3: "required",
       // shortlistingQuestion4: "required",
            ourteam: "required",
            rolefocus: "required",
            whatyoudo: "required",
            whatyoubring: "required",
			itw_deadline: "required",
			closingdate: "required",
			specMOR: "required",
			specAD: "required",
        costcode: "required",
        usingagency: "required",
        agencyrate: "required",
			},

          messages: {
            division: "Please Select a Division.",
            region: "Please Select a Region.",
            area: "Please Select an Area",
            recmanager: "Please enter a recruiting manager name",
            //linemanager: "Please enter a line manager name",
            jobrole: "Please enter a job title",
            locationF: "Please Select a location",
            baseF: "Please Select where this role is based",
            startdate: "Please Select a date in the calendar",
            existingrole: "Please confirm the hire type.",
            //recmethod: "Please Select a recruitment method",
           // FPTime: "Please confirm if the contract is full or part time",
            permtemp: "Please confirm if the contract is permanent or temporary",
            enddate: "Please enter the expected end date",
            bizcase: "Please provide a business case",
            leavername: "Please enter the name of the person who left",
            leaverjob: "Please enter the leaver's job title",
            leaversalary: "Please enter the leaver's current salary",
            reason: "Please enter a reason for leaving",
            annualsalary: "Please enter the annual salary offered for the position",
            //medicalinsurance: "Please Select the type of medical insurance",
			runtime: "Please Select the advert runtime",
			flexibleWorking: "Please Select if this role allows flexible working",
      //  shortlistingQuestion1: "Please enter a shortlisting question",
       // shortlistingQuestion2: "Please enter a shortlisting question",
      //  shortlistingQuestion3: "Please enter a shortlisting question",
     //   shortlistingQuestion4: "Please enter a shortlisting question",
            ourteam: "Please enter a text about your team",
            rolefocus: "Please describe the role",
            whatyoudo: "Please describe the duties",
            whatyoubring: "Please enter some text",
			itw_deadline: "Please select a decision date.",
			closingdate: "Please select a closing date",
			specMOR: "Please provide some details about the external recuitment",
			specAD: "Please provide some details about the advert runtime",
        costcode: "Please enter the cost code",
        usingagency: "please specify if you are using an agency",
        agencyrate: "Please speify the Agency rate",
          },

          submitHandler: function() {
            checkRecMethod();
            if (checkRecMethod() != false) {
              submitForm();
            }
          },

          invalidHandler: function(event, validator) {
            var errorDiv = $('#error-msg');
            var errors = validator.numberOfInvalids();
            if (errors) {
              errorDiv.html('PLEASE CHECK THE ' + errors + ' ERROR(S) HIGHLIGHTED IN RED').show();
            } else {
              errorDiv.hide();
            }
          }
        });

      } else {
		  
		  
console.log("no attchment found...");
       // $('#fileinput-label-wrapper').hide();
		  
        // validation jQuery
        $mainForm.validate({
          rules: {
            division: "required",
            region: "required",
            area: "required",
            recmanager: "required",
            //linemanager: "required",
            jobrole: "required",
            locationF: "required",
            baseF: "required",
            startdate: "required",
            existingrole: "required",
            //recmethod: "required",
          //  FPTime: "required",
            permtemp: "required",
            enddate: "required",
            bizcase: "required",
            fileinput: "required",
            leavername: "required",
            leaverjob: "required",
            leaversalary: "required",
            reason: "required",
            annualsalary: "required",
            companycar: "required",
            //medicalinsurance: "required",
			runtime: "required",
			flexibleWorking: "required",
      //  shortlistingQuestion1: "required",
      //  shortlistingQuestion2: "required",
      //  shortlistingQuestion3: "required",
      //  shortlistingQuestion4: "required",
            ourteam: "required",
            rolefocus: "required",
            whatyoudo: "required",
            whatyoubring: "required",
			itw_deadline: "required",
			closingdate: "required",
			specMOR: "required",
			specAD: "required",
        costcode: "required",
        usingagency: "required",
        agencyrate: "required",
          },

          messages: {
            division: "Please Select a Division.",
            region: "Please Select a Region.",
            area: "Please Select an Area",
            recmanager: "Please enter a recruiting manager name",
            //linemanager: "Please enter a line manager name",
            jobrole: "Please enter a job title",
            locationF: "Please Select a location",
            baseF: "Please Select where this role is based",
            startdate: "Please Select a date in the calendar",
            existingrole: "Please confirm the hire type.",
            //recmethod: "Please Select a recruitment method",
           // FPTime: "Please confirm if the contract is full or part time",
            permtemp: "Please confirm if the contract is permanent or temporary",
            enddate: "Please enter the expected end date",
            bizcase: "Please provide a business case",
            fileinput: "Please attach a job description",
            leavername: "Please enter the name of the person who left",
            leaverjob: "Please enter the leaver's job title",
            leaversalary: "Please enter the leaver's current salary",
            reason: "Please enter a reason for leaving",
            annualsalary: "Please enter the annual salary offered for the position",
            //medicalinsurance: "Please Select the type of medical insurance",
			runtime: "Please Select the advert runtime",
			flexibleWorking: "Please Select if this role allows flexible working",
      //  shortlistingQuestion1: "Please enter a shortlisting question",
      //  shortlistingQuestion2: "Please enter a shortlisting question",
      //  shortlistingQuestion3: "Please enter a shortlisting question",
      //  shortlistingQuestion4: "Please enter a shortlisting question",
            ourteam: "Please enter a text about your team",
            rolefocus: "Please describe the role",
            whatyoudo: "Please describe the duties",
            whatyoubring: "Please enter some text",
			itw_deadline: "Please select a decision date.",
			closingdate: "Please select a closing date",
			specMOR: "Please provide some details about the external recuitment",
			specAD: "Please provide some details about the advert runtime",
        costcode: "Please enter the cost code",
        usingagency: "please specify if you are using an agency",
        agencyrate: "Please speify the Agency rate",
          },

          submitHandler: function() {
            checkRecMethod();
            if (checkRecMethod() != false) {
              submitForm();
            }
          },

          invalidHandler: function(event, validator) {
            var errorDiv = $('#error-msg');
            var errors = validator.numberOfInvalids();
            if (errors) {
              errorDiv.html('PLEASE CHECK THE ' + errors + ' ERROR(S) HIGHLIGHTED IN RED').show();
            } else {
              errorDiv.hide();
            }
          }
        });
      }
    }
  });

  // Grab existing values of the related item 
  var division = targetListItem.get_item('Division');
  var region = targetListItem.get_item('Region');
  var area = targetListItem.get_item('Area');
  var recManager = targetListItem.get_item('RecruitingManagerText');
 // var lineManager = targetListItem.get_item('LineManager').get_email(); 
  var BAApprover = targetListItem.get_item('BAApprover');
  var FDApprover = targetListItem.get_item('FDApprover');
  var TMApprover = targetListItem.get_item('TMApprover');

  var jobRole = targetListItem.get_item('JobRole');
  var locationF = targetListItem.get_item('Location');
  var baseF = targetListItem.get_item('Base');
  var startDate = targetListItem.get_item('AnticipatedStartDate');
  var startDateF = String.format('{0:yyyy}-{0:MM}-{0:dd}', new Date(startDate)); // convert date to the correct format 
  var existRole = targetListItem.get_item('ExistingRole');
  var recMeth = targetListItem.get_item('MethodOfRecruitment');
  var specRecMeth = targetListItem.get_item('SpecifyMethodOfRecruitment');

  //var FPTime = targetListItem.get_item('FullTimePartTime');
  var permTemp = targetListItem.get_item('PermanentTemporary');
  //var materCov = targetListItem.get_item('MaternityCover');
  var endDate = targetListItem.get_item('ExpectedEndDate');
  var endDateF = String.format('{0:yyyy}-{0:MM}-{0:dd}', new Date(endDate)); // convert date

  var budgCost = targetListItem.get_item('BudgetedCostOfRecruitment');
  
  
  var costCode = targetListItem.get_item('CostCode');
  var usingAgency = targetListItem.get_item('UsingAgency');
  var agencyRate = targetListItem.get_item('AgencyRate');
  
  var bizCase = targetListItem.get_item('BusinessCase');

  var leaverName = targetListItem.get_item('NameOfLeaver');
  var leaverJob = targetListItem.get_item('LeaverJobTitle');
  var leaverSal = targetListItem.get_item('LeaversSalary');
  var leaverReason = targetListItem.get_item('ReasonForLeaving');

  var annualSalary = targetListItem.get_item('AnnualSalary');
  var pension = targetListItem.get_item('PensionContribution');
  var commission = targetListItem.get_item('Commission');
  var bonus = targetListItem.get_item('Bonus');
  var bonusPercent = targetListItem.get_item('Bonus_x0025_');
  var compCar = targetListItem.get_item('CompanyCar');
  var carAllow = targetListItem.get_item('CarAllowance');
  var addRenum = targetListItem.get_item('AdditionalRenumeration');
  //var medical = targetListItem.get_item('CompayPaidMedicalInsurance');
  var runtime = targetListItem.get_item('Runtime');
  var specRuntime = targetListItem.get_item('SpecifyRuntime');
  var flexibleWorking = targetListItem.get_item('FlexibleWorking');
  var shortlistingQuestion1 = targetListItem.get_item('shortlistingQuestion1');
  var shortlistingQuestion2 = targetListItem.get_item('shortlistingQuestion2');
  var shortlistingQuestion3 = targetListItem.get_item('shortlistingQuestion3');
  var shortlistingQuestion4 = targetListItem.get_item('shortlistingQuestion4');
  var ourTeam = targetListItem.get_item('OurTeam');
  var roleFocus = targetListItem.get_item('RoleFocus');
  var whatYouDo = targetListItem.get_item('WhatYouDo');
  var whatBring = targetListItem.get_item('WhatYouBring');
  var howApply = targetListItem.get_item('HowToApply');
  var decisionItvwDate = targetListItem.get_item('DecisionInterviewDate');
  var decisionItvwDateF = String.format('{0:yyyy}-{0:MM}-{0:dd}', new Date(decisionItvwDate)); // convert date

  var closingDate = targetListItem.get_item('ApplicationClosingDate');
  var closingDateF = String.format('{0:yyyy}-{0:MM}-{0:dd}', new Date(closingDate)); // convert date

	var R2RStatus = targetListItem.get_item('Status');
	
  try {
    if (stage === "BA") {
		if (R2RStatus != 'BA Approval') { 
			alert("This approval has already been actioned. Now taking you back to the dashboard...");
			window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
		}
    }
  } catch (ex) {console.log(ex);}
  
  try {
    if (stage === "MD") {
		if (R2RStatus != 'Managerial Approval') { 
			alert("This approval has already been actioned. Now taking you back to the dashboard...");
			window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
		}
      console.log('Setting BA comments...');
      document.getElementById('ba-comments').innerHTML = targetListItem.get_item('BAComments');
    }
  } catch (ex) {console.log(ex);}
  
  try {
    if (stage === "FD") {
		if (R2RStatus != 'FD Approval') { 
			alert("This approval has already been actioned. Now taking you back to the dashboard...");
			window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
		}
      console.log('Setting manager and BA comments...');
      document.getElementById('man-comments').innerHTML = targetListItem.get_item('ManagerComments');
      document.getElementById('ba-comments').innerHTML = targetListItem.get_item('BAComments');
    }
  } catch (ex) {console.log(ex);}
  
  
  try {
    if (stage === "TM") {
		if (R2RStatus != 'Talent Manager Approval') { 
			alert("This approval has already been actioned. Now taking you back to the dashboard...");
			window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
		}
      console.log('Setting manager, BA and FD comments...');
      document.getElementById('man-comments').innerHTML = targetListItem.get_item('ManagerComments');
      document.getElementById('ba-comments').innerHTML = targetListItem.get_item('BAComments');
	  if (targetListItem.get_item('FDComments') == '' || targetListItem.get_item('FDComments') == undefined) 
		{ $('#fd-label, #fd-comments').hide(); }
	  else {  document.getElementById('fd-comments').innerHTML = targetListItem.get_item('FDComments'); }
    }
  } catch (ex) {console.log(ex);}
  
  
  //::::::::::::: set fields with these values ::::::::::::: 

  $divisionInput.val(division);
  setTimeout(function() {
    $areaInput.val(area);
  }, 1000)
  $recManagerInput.val(recManager);
  //$lineManagerInput.val(lineManager).change(); this field is not needed anymore

  $jobRoleInput.val(jobRole);
  $locationInput.val(locationF);
  $baseInput.val(baseF);
  $startDateInput.val(startDateF)

  // if EXISTING ROLE was ticked, set value to 'Yes' + toggle "leaver section"
  if (existRole == "Leaver Replacement") {
    $existingRoleInput.val('Leaver Replacement');
    $leaverSection.show();
  } else {
    $existingRoleInput.val(existRole);
  }

  // retrieve RECRUITMENT methods
  try {if (recMeth.includes('External')) {
	   if (recMeth.includes('External - Other')) {
    $("input[value='External - Other']").prop("checked", true);
    $("input[value='External - Other']").change();
	   }
	   if (recMeth.includes('External - Social Media')) {
    $("input[value='External - Social Media]").prop("checked", true);
    $("input[value='External - Socila Media']").change();
	   }
	   if (recMeth.includes('External - Standard')) {
    $("input[value='External - Standard]").prop("checked", true);
    $("input[value='External - Standard']").change();
	   }
  }

  if (recMeth.includes('Internal')) {
    $("input[value='Internal']").prop("checked", true);
    $("input[value='Internal']").change();
	   }
  
  if (recMeth.includes('Agency')) {
    $("input[value='Agency']").prop("checked", true);
    $("input[value='Agency']").change();
	   }
  } catch (ex) {}
  
  try {
    // update RECRUITMENT METHODS input with the name of the ticked values     
    $("#ddcl-method-input > span > span").html(recMeth.join(', '));
  } catch (ex) {}

  if ((!(stage) || (stage === 'edit'))) {  } else {
	  try {
    document.querySelector("#job-info-section > div:nth-child(2) > div > label").innerHTML += "<br><br>"+ recMeth.join(', ').replace('Please Select...', '');
	  } catch (ex) {}
	
  }

  try {$specMORInput.val(specRecMeth);} catch (ex) {}
 // $FPTimeInput.val(FPTime);

  $permTempInput.val(permTemp);
  
  // if PERMANENT/TEMPORARY is set to 'Temporary', toggle the "maternity cover" section
  /* if (permTemp === 'Temporary') {
    $('#mater-wrapper, #enddate-wrapper').show();
  } */

  //if CONTRACT is 'Temporary', set the fields "Maternity Cover" & "End Date"  
  
  if (existRole = "Maternity Cover") {$maternityInput.val('Yes') ;} else { $maternityInput.val('No');} // if maternity cover was ticked,returns TRUE, set value to 'Yes'

  $endDateInput.val(endDateF);
  $budgCostInput.val(budgCost);
  
  $costCodeInput.val(costCode);
  
  if (usingAgency) {  $usingAgencyInput.val("Yes"); }
  else {  $usingAgencyInput.val("No"); }
  
  $agencyRateInput.val(agencyRate);
  
  // if using agency is set to 'Yes', toggle the "agency rate" question
   if (usingAgency) {
    $('#agency-rate-wrapper').show();
  } else {
	$('#agency-rate-wrapper').hide();
	}
  
  $specADInput.val(specRuntime);
  
  $bizCaseInput.val(bizCase);

  $leaverNameInput.val(leaverName);
  $leaverJobInput.val(leaverJob);
  $leaverSalaryInput.val(leaverSal);
  $leaverReasonInput.val(leaverReason);

  $salaryInput.val(annualSalary);
  $pensionInput.val(pension);
  $commissionInput.val(commission);

  // if BONUS was ticked, set it to 'Yes' + toggle the "bonus percent" section
  if (bonus) {
    $bonusInput.val('Yes');
    $bonusPercentSection.show();
  } else {
    $bonusInput.val('No');
  }

  $bonusPercentInput.val(bonusPercent);
  compCar ? $carInput.val('Yes') : $carInput.val('No');
  $carAllowInput.val(carAllow);
  $addRenumInput.val(addRenum);
 // $medicalInput.val(medical);
  $runtimeInput.val(runtime);
  
  // if Advert Runtime id 'other', toggle "leaver section"
  if (runtime = "Other") {
    $('#specADdiv').show();
  } else {
		$('#specADdiv').hide();
  }

	$specADInput.val(specRuntime);
    $flexibleWorkingInput.val(flexibleWorking);
  
  $shortlistingQuestion1Input.val(shortlistingQuestion1);
  $shortlistingQuestion2Input.val(shortlistingQuestion2);
  $shortlistingQuestion3Input.val(shortlistingQuestion3);
  $shortlistingQuestion4Input.val(shortlistingQuestion4);

  $ourTeamInput.val(ourTeam);
  $roleFocusInput.val(roleFocus);
  $whatYouDoInput.val(whatYouDo);
  $whatYouBringInput.val(whatBring);
  $howApplyInput.val(howApply);
  $itwDeadlineInput.val(decisionItvwDateF);
  $closingDateInput.val(closingDateF);

  // execute after division has loaded
  setTimeout(function() {
    $regionInput.val(region);
  }, 1000);

}

function requestOnQueryFailed(sender, args) {
  alert('Failed to populate form data. ' + args.get_message() + '\n' + args.get_stackTrace());
}


// ============================  🆂🆄🅱🅼🅸🆃 🅵🅾🆁🅼 ===========================

function submitForm() {

  if (!creatingRequest) {
    console.log('____-_-_Submitting form starting_-_-___');
    creatingRequest = true;
    draftMode = false;
    submitMode = false;

    var submitBtn = document.getElementById('submit-btn');
    var draftBtn = document.getElementById('save-btn');

    // check for focus
    var submitIsFocused = (document.activeElement === submitBtn);
    var draftIsFocused = (document.activeElement === draftBtn);

    if (submitIsFocused) {
      console.log('submit btn is focussed');
      submitMode = true;
      draftMode = false;
    }
    if (draftIsFocused) {
      console.log('draft btn is focused');
      draftMode = true;
      submitMode = false;
    }

    // hide form 
    $mainForm.hide();
    // creating item
    createListItem();

  }
}


// ========================== 🅲🆁🅴🅰🆃🅴 🅽🅴🆆 🅸🆃🅴🅼 ==========================
siteUrl = '/sites/UK-HR-RequestToRecruit/';
oListItem = '';

function createListItem() {
  console.log('createListItem function is firing!');
  var clientContext = new SP.ClientContext(siteUrl);
  var oList = clientContext.get_web().get_lists().getByTitle('Requests'); // targets "Requests" list    


  if (stage === 'edit') {
    oListItem = oList.getItemById(R2RID);
  } else {
    var itemCreateInfo = new SP.ListItemCreationInformation();
    oListItem = oList.addItem(itemCreateInfo);
  }


  try {
    $spinner.show();

    divisionVal = $('#division-input option:selected').val(); // grab value of the selected option
    regionVal = $('#region-input option:selected').val();
    areaVal = $('#area-input option:selected').val();
    recManagerVal = $recManagerInput.val();
    //lineManagerVal = $lineManagerInput.val();
    jobRoleVal = $jobRoleInput.val();
    locationVal = $('#location-input option:selected').val();
    baseVal = $('#base-input option:selected').val();
    startDateVal = $startDateInput.val();
    existingRoleVal = $('#exist-role-input option:selected').val();
    //methodRecVal = $('#method-input option:selected').val();
    //methodRecVal = document.querySelector("#ddcl-method-input > span > span").innerText;
	//specRecMethVal = $specMORInput.val();
    //specRecMethVal = $('#specMOR-input').val();
   // FPTimeVal = $('#FP-time-input option:selected').val();
    contractTypeVal = $('#perm-temp-input option:selected').val();
    //materVal = $('#maternity-input option:selected').val();
    endDateVal = $('#end-date-input').val();
    budgCostVal = $budgCostInput.val();
	
	costCodeVal = $costCodeInput.val();
	usingAgencyVal = $usingAgencyInput.val();
	agencyRateVal = $agencyRateInput.val();
	
    bizCaseVal = $('#bizcase-input').val();

    salaryVal = $salaryInput.val();
    pensionVal = $('#pension-input').val();
    commissionVal = $('#commission-input').val();
    bonusVal = $('#bonus-input option:selected').val();
    bonusPercentVal = $('#bonus-percent-input').val();
    carVal = $('#car-input option:selected').val();
    carAllowVal = $('#car-allow-input').val();
    addRenumVal = $('#add-renum-input').val();
    //medicalVal = $('#medical-input option:selected').val();

    leaverVal = $('#leaver-name-input').val();
    leaverJobVal = $('#leaver-JT-input').val();
    leaverSalaryVal = $leaverSalaryInput.val();
    leaverReasonVal = $('#leaving-reason-input').val();
/*
    runtimeVal = $('#runtime-input option:selected').val();
	specRuntimeVal = $specADInput.val();
    flexibleWorkingVal = $('#flexibleWorking-input option:selected').val();
*/

    shortlistingQuestion1Val = $('#shortlistingQuestion1-input').val();
    shortlistingQuestion2Val = $('#shortlistingQuestion2-input').val();
    shortlistingQuestion3Val = $('#shortlistingQuestion3-input').val();
    shortlistingQuestion4Val = $('#shortlistingQuestion4-input').val();
/*
    ourTeamVal = $('#our-team-input').val();
    roleFocusVal = $('#role-focus-input').val();
    whatYouDoVal = $('#what-you-do-input').val();
    whatYouBringVal = $('#what-you-bring-input').val();
    how2ApplyVal = $('#how-apply-input').val();
    itwDeadlineVal = $('#itw_deadline-input').val();
    applyDateVal = $('#appli-closing-input').val();
*/
  } catch (err) {
    console.log(err);
  }

  // set item columns with related IF fields are not empty
  try {

    if (!!divisionVal) {
      oListItem.set_item('Division', divisionVal)
    }

    if (!!regionVal) {
      oListItem.set_item('Region', regionVal)
    };
    if (!!areaVal) {
      oListItem.set_item('Area', areaVal)
    };

    if (!!recManagerVal) {
      oListItem.set_item('RecruitingManagerText', recManagerVal)
    };

    if (!!recruitingManagerEmail) {
      spUser = SP.FieldUserValue.fromUser(recruitingManagerEmail);
      oListItem.set_item('RecruitingManager', spUser);
    };

    /*if (!!lineManagerVal) {
      oListItem.set_item('LineManagerText', lineManagerVal)
    };

    if (!!lineManagerEmail) {
      spUser = SP.FieldUserValue.fromUser(lineManagerEmail);
      oListItem.set_item('LineManager', spUser);
    }; */

    if (!!jobRoleVal) {
      oListItem.set_item('JobRole', jobRoleVal)
    };

    if (!!locationVal) {
      oListItem.set_item('Location', locationVal)
    };

    if (!!baseVal) {
      oListItem.set_item('Base', baseVal)
    };
	
    if (!!startDateVal) {
      // Converting end date input value to a date object 
      //var dateParts = startDateVal.split("-");
      // month is 0-based, that's why we need dataParts[1] - 1
      //var startDateObj = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
      //oListItem.set_item('AnticipatedStartDate', startDateObj);

      oListItem.set_item('AnticipatedStartDate', startDateVal);
    }

    if (!!existingRoleVal) {
      oListItem.set_item('ExistingRole', existingRoleVal)
    };
	/*
    if (!!methodRecVal) {
      var methodRec = methodRecVal.split(', ')
      oListItem.set_item('MethodOfRecruitment', methodRec);
    }; */
	
    /*if (!!specRecMethVal) {
      oListItem.set_item('SpecifyMethodOfRecruitment', specRecMethVal)
    }; */
	
   /* if (!!FPTimeVal) {
      oListItem.set_item('FullTimePartTime', FPTimeVal)
    }; */

    if (!!contractTypeVal) {
      oListItem.set_item('PermanentTemporary', contractTypeVal)
    };

   /* if (!!materVal) {
      oListItem.set_item('MaternityCover', materVal);
    }
*/
    if (!!endDateVal) {
      // Converting end date input value to a date object 
      //var dateParts = endDateVal.split("-");
      //var endDateObj = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
      //oListItem.set_item('ExpectedEndDate', endDateObj);

      oListItem.set_item('ExpectedEndDate', endDateVal);
    }

    if (!!budgCostVal) {
      oListItem.set_item('BudgetedCostOfRecruitment', budgCostVal);
    }

    if (!!costCodeVal) {
      oListItem.set_item('CostCode', costCodeVal);
    }

    if (!!usingAgencyVal) {
      oListItem.set_item('UsingAgency', usingAgencyVal);
    }

    if (!!agencyRateVal) {
      oListItem.set_item('AgencyRate', agencyRateVal);
    }

    if (!!bizCaseVal) {
      oListItem.set_item('BusinessCase', bizCaseVal);
    }

    if (!!salaryVal) {
      oListItem.set_item('AnnualSalary', salaryVal);
    }

    if (!!pensionVal) {
      oListItem.set_item('PensionContribution', pensionVal);
    }

    if (!!commissionVal) {
      oListItem.set_item('Commission', commissionVal);
    }
	
    if (!!bonusVal) {
      oListItem.set_item('Bonus', bonusVal);
    }

    if (!!bonusPercentVal) {
      oListItem.set_item('Bonus_x0025_', bonusPercentVal);
    }

    if (!!carVal) {
      oListItem.set_item('CompanyCar', carVal);
    }

    if (!!carAllowVal) {
      oListItem.set_item('CarAllowance', carAllowVal);
    }

    if (!!addRenumVal) {
      oListItem.set_item('AdditionalRenumeration', addRenumVal);
    }
	
    /*if (!!medicalVal) {
      oListItem.set_item('CompayPaidMedicalInsurance', medicalVal);
    } */
	
    if (!!leaverVal) {
      oListItem.set_item('NameOfLeaver', leaverVal);
    }

    if (!!leaverJobVal) {
      oListItem.set_item('LeaverJobTitle', leaverJobVal);
    }

    if (!!leaverSalaryVal) {
      oListItem.set_item('LeaversSalary', leaverSalaryVal);
    }

    if (!!leaverReasonVal) {
      oListItem.set_item('ReasonForLeaving', leaverReasonVal);
    }

  /*  if (!!runtimeVal) {
      oListItem.set_item('Runtime', runtimeVal);
    }
	
    if (!!specRuntimeVal) {
      oListItem.set_item('SpecifyRuntime', specRuntimeVal)
    }; 
	
	
    if (!!flexibleWorkingVal) {
      oListItem.set_item('FlexibleWorking', flexibleWorkingVal);
    }
	
    if (!!ourTeamVal) {
      oListItem.set_item('OurTeam', ourTeamVal);
    }

    if (!!roleFocusVal) {
      oListItem.set_item('RoleFocus', roleFocusVal);
    }

    if (!!whatYouDoVal) {
      oListItem.set_item('WhatYouDo', whatYouDoVal);
    }

    if (!!whatYouBringVal) {
      oListItem.set_item('WhatYouBring', whatYouBringVal);
    }

    if (!!how2ApplyVal) {
      oListItem.set_item('HowToApply', how2ApplyVal);
    }

    if (!!itwDeadlineVal) {
      oListItem.set_item('DecisionInterviewDate', itwDeadlineVal);
    }

    if (!!applyDateVal) {
      oListItem.set_item('ApplicationClosingDate', applyDateVal);
    }
	*/
	
    if (!!shortlistingQuestion1Val) {
      oListItem.set_item('shortlistingQuestion1', shortlistingQuestion1Val);
    }
	
    if (!!shortlistingQuestion2Val) {
      oListItem.set_item('shortlistingQuestion2', shortlistingQuestion2Val);
    }
	
    if (!!shortlistingQuestion3Val) {
      oListItem.set_item('shortlistingQuestion3', shortlistingQuestion3Val);
    }
	
    if (!!shortlistingQuestion4Val) {
      oListItem.set_item('shortlistingQuestion4', shortlistingQuestion4Val);
    }
	

    // set status column depending on which button was pressed
    if (draftMode === true) {
      oListItem.set_item('Status', 'Draft');
    }

    if (submitMode === true) {
      oListItem.set_item('Status', 'Submitted');
    }

  } catch (error) {
    $spinner.hide();
    console.log('could not set field values');
    alert('createListItem error is: ' + error);
    $mainForm.show();
  }

  oListItem.update();
  clientContext.load(oListItem);
  clientContext.executeQueryAsync(Function.createDelegate(this, this.createonQuerySucceeded), Function.createDelegate(this, this.createonQueryFailed));
}

// upload the attached file after the item is created
function createonQuerySucceeded(sender, args) {
  currentID = oListItem.get_id();
  try {
    // if there is a file uploaded, upload the file
    if (document.getElementById("fileinput").value !== "") {
      uploadTheFile();
    } else {

      // show relevant confirmation message
      if (submitMode === true) {
        $confirmMsg.addClass('alert alert-success').html('Thank you for submitting the request. There will be a slight delay of up to 3 mins before the request will be shown on the dashboard. You will now be taken to the dashboard.').show();
      }
      if (draftMode === true) {
        $confirmMsg.addClass('alert alert-warning').html('Your request has been saved as a Draft. There will be a slight delay of up to 3 mins before the request will be shown on the dashboard. You will now be taken to the dashboard..').show();
      }

      setTimeout(function() {
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
      }, 4000); // redirects user to "DASHBOARD" 

    }

  } catch (error) {
    console.log('no attached file');
    alert('createonQuerySucceeded error is: ' + error);
    $mainForm.show();
  }

}

function createonQueryFailed(sender, args) {
  $spinner.hide();
  alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
  $mainForm.show();
}

// ================================== ATTACH FILE =================================

function uploadTheFile() {
  var listTitle = 'Requests';
  var itemId = currentID;
  var fileInputElem = document.getElementById("fileinput");
  var file = fileInputElem.files[0];


  processUpload(file, listTitle, itemId,
    function() {
      $spinner.hide();
      //alert('Item created and file uploaded successfully.');
      if (submitMode === true) {
		submitRequestFlow();
        $confirmMsg.addClass('alert alert-success').html('Thank you for submitting the request. There will be a slight delay of up to 3 mins before the request will be shown on the dashboard. You will now be taken to the dashboard.').show();
      }
      if (draftMode === true) {
        $confirmMsg.addClass('alert alert-warning').html('Your request has been saved as a Draft. There will be a slight delay of up to 3 mins before the request will be shown on the dashboard. You will now be taken to the dashboard.').show();
		
      setTimeout(function() {
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
      }, 5000); // redirects user to "DASHBOARD" 
      }

    },
    function(sender, args) {
      $spinner.hide();
      //alert('Item created, but file not uploaded. Please try again or contact an administrator. \n' + + args.get_message() + '\n' + args.get_stackTrace());
    }
  );
}


function processUpload(fileInputElem, listTitle, itemId, success, error) {
  try {
    var reader = new FileReader();
    reader.onload = function(result) {
      var fileContent = new Uint8Array(result.target.result);
      performAttachmentUpload(listTitle, fileInputElem.name, itemId, fileContent, success, error);
    };
    reader.readAsArrayBuffer(fileInputElem);
  } catch (error) {
    $spinner.hide();
    alert('File not uploaded. Please try again or contact an administrator');
  }

}


function performAttachmentUpload(listTitle, fileName, itemId, fileContent, success, error) {
  ensureAttachmentFolder(listTitle, itemId,
    function(folder) {
      var attachmentFolderUrl = folder.get_serverRelativeUrl();
      uploadFile(attachmentFolderUrl, fileName, fileContent, success, error);
    },
    error);
}


function ensureAttachmentFolder(listTitle, itemId, success, error) {
  try {
    var ctx = SP.ClientContext.get_current();
    var web = ctx.get_web();
    var list = web.get_lists().getByTitle(listTitle);
    ctx.load(list, 'RootFolder');
    var item = list.getItemById(itemId);
    ctx.load(item);
    ctx.executeQueryAsync(
      function() {
        var attachmentsFolder;
        if (!item.get_fieldValues()['Attachments']) {
          /* Attachments folder exists? */
          var attachmentRootFolderUrl = String.format('{0}/Attachments', list.get_rootFolder().get_serverRelativeUrl());
          var attachmentsRootFolder = ctx.get_web().getFolderByServerRelativeUrl(attachmentRootFolderUrl);
          //Note: Here is a tricky part. 
          //Since SharePoint prevents the creation of folder with name that corresponds to item id, we are going to:   
          //1)create a folder with name in the following format '_<itemid>'
          //2)rename a folder from '_<itemid>'' into '<itemid>'
          //This allow to bypass the limitation of creating attachment folders
          attachmentsFolder = attachmentsRootFolder.get_folders().add('_' + itemId);
          attachmentsFolder.moveTo(attachmentRootFolderUrl + '/' + itemId);
          ctx.load(attachmentsFolder);
        } else {
          var attachmentFolderUrl = String.format('{0}/Attachments/{1}', list.get_rootFolder().get_serverRelativeUrl(), itemId);
          attachmentsFolder = ctx.get_web().getFolderByServerRelativeUrl(attachmentFolderUrl);
          ctx.load(attachmentsFolder);
        }
        ctx.executeQueryAsync(
          function() {
            success(attachmentsFolder);
          },
          error);
      },
      error);
  } catch (e) {
    console.log("Error occurred in ensureAttachmentFolder: " + e.message);
  }
}


function uploadFile(folderUrl, fileName, fileContent, success, error) {
  try {
    var ctx = SP.ClientContext.get_current();
    var folder = ctx.get_web().getFolderByServerRelativeUrl(folderUrl);
    var encContent = new SP.Base64EncodedByteArray();
    for (var b = 0; b < fileContent.length; b++) {
      encContent.append(fileContent[b]);
    }

    varbase64 = btoa(encContent); //encContent works but cannot be read correctly by the workflow, lets try this

    var createInfo = new SP.FileCreationInformation();
    createInfo.set_content(encContent); //passing varbase64 or encContent
    createInfo.set_url(fileName);
    folder.get_files().add(createInfo);
    ctx.executeQueryAsync(success, error);
  } catch (e) {
    console.log('error in uploadFile function: ', e.message);

  }
}

// Display name of uploaded files
function displayFileNames() {
  var fileInputVal = $fileInput.val();
  var updatedStr = fileInputVal.replace("C:\\fakepath\\", "");
  $('#fileinput-label > span').html(updatedStr); // update upload btn inner text
}


// ======================= Set division dropdown menu ====================

function setDivisions() {
  console.log("populating division dropdown...");
  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  divisionsList = clientContext.get_web().get_lists().getByTitle('Division');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query></Query></View>'
  );
  this.allListItems = divisionsList.getItems(camlQuery);
  clientContext.load(allListItems);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetDivisionsQuerySucceeded),
    Function.createDelegate(this, this.onSetDivisionsQueryFailed)
  );
}

function onSetDivisionsQuerySucceeded(sender, args) {
  var eachDivisionName = "";
  var listItemEnumerator = allListItems.getEnumerator();

  divisionInput_ = document.getElementById('division-input');
  $("#division-input").empty()

  while (listItemEnumerator.moveNext()) {
    var divisionListItem = listItemEnumerator.get_current();
    eachDivisionName = divisionListItem.get_item('Title');
    var option = document.createElement("option");
    option.value = divisionListItem.get_item('Title');
    option.appendChild(document.createTextNode(eachDivisionName));
    //option.text = eachDivisionName;

    divisionInput_.appendChild(option);
  }

  var my_options = $("#division-input option");
  my_options.sort(function(a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $("#division-input").empty().append(my_options);
  $("#division-input").prepend("<option selected>Please Select...</option>");

  setOffices();
}

function onSetDivisionsQueryFailed() {
  alert('Division Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

// ======================= Set region dropdown menu ====================

function setRegions() {
  console.log("populating region dropdown...");
  divisionVal = $('#division-input option:selected').val();

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  authorizationsList = clientContext.get_web().get_lists().getByTitle('Authorisation Table');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query><Where><Eq><FieldRef Name=\'Division\'/>' +
    '<Value Type=\'Lookup\'>' + divisionVal + '</Value></Eq></Where></Query></View>'
  );
  this.allListItems2 = authorizationsList.getItems(camlQuery);
  clientContext.load(allListItems2);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetRegionsQuerySucceeded),
    Function.createDelegate(this, this.onSetRegionsQueryFailed)
  );

}

function setRegionsAll() {
  console.log("populating region dropdown with all...");
  divisionVal = $('#division-input option:selected').val();

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  regionsList = clientContext.get_web().get_lists().getByTitle('Region');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query></Query></View>'
  );
  this.allListItems2 = regionsList.getItems(camlQuery);
  clientContext.load(allListItems2);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetRegionsAllQuerySucceeded),
    Function.createDelegate(this, this.onSetRegionsQueryFailed)
  );

}

function onSetRegionsQuerySucceeded(sender, args) {
  var eachRegionName = "";
  var listItemEnumerator = allListItems2.getEnumerator();

  regionInput = document.getElementById('region-input');
  $regionInput.empty();

  while (listItemEnumerator.moveNext()) {
    var regionListItem = listItemEnumerator.get_current();

    // eachRegionName variable not found, so add it as an option
    eachRegionName = regionListItem.get_item('Region').get_lookupValue();
    var option = document.createElement("option");
    option.value = regionListItem.get_item('Region').get_lookupValue();
    option.appendChild(document.createTextNode(eachRegionName));

    opts = regionInput.options;
    var present = false;

    for (i = 0; i < opts.length; i++) {
      if (opts[i].value == eachRegionName) {
        present = true;
      }
    }

    // add the option if it s not already in the options array
    if (!(present)) {
      regionInput.appendChild(option);
    }
  }

  var my_options = $("#region-input option");
  my_options.sort(function(a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  });

  $regionInput.empty().append(my_options);
  $regionInput.prepend("<option selected>Please Select...</option>");

}



function onSetRegionsAllQuerySucceeded(sender, args) {
  var eachRegionName = "";
  var listItemEnumerator = allListItems2.getEnumerator();

  regionInput = document.getElementById('region-input');
  $regionInput.empty();

  while (listItemEnumerator.moveNext()) {
    var regionListItem = listItemEnumerator.get_current();
    eachRegionName = regionListItem.get_item('Title');
    var option = document.createElement("option");
    option.value = regionListItem.get_item('Title');
    option.appendChild(document.createTextNode(eachRegionName));

    regionInput.appendChild(option);
  }
  var my_options = $("#region-input option");
  my_options.sort(function(a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $regionInput.empty().append(my_options);
  $regionInput.prepend("<option selected>Please Select...</option>");

}

function onSetRegionsQueryFailed(sender, args) {
  alert('Region Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

// ======================= Set area dropdown menu ====================

function setAreas() {
  console.log("populating area dropdown...");
  regionVal = $('#region-input option:selected').val();

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  authList = clientContext.get_web().get_lists().getByTitle('Authorisation Table');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query><Where><And><Eq><FieldRef Name=\'Division\'/>' +
    '<Value Type=\'Lookup\'>' + divisionVal + '</Value></Eq><Eq><FieldRef Name=\'Region\'/>' +
    '<Value Type=\'Lookup\'>' + regionVal + '</Value></Eq></Where></And></Query></View>'
  );
  this.allListItems3 = authList.getItems(camlQuery);
  clientContext.load(allListItems3);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetAreasQuerySucceeded),
    Function.createDelegate(this, this.onSetAreasQueryFailed)
  );

}

function setAreasAll() {
  console.log("populating area dropdown with all...");

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  areasList = clientContext.get_web().get_lists().getByTitle('Area');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query></Query></View>'
  );
  this.allListItems3 = areasList.getItems(camlQuery);
  clientContext.load(allListItems3);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetAreasAllQuerySucceeded),
    Function.createDelegate(this, this.onAreasAllQueryFailed)
  );

}

function onSetAreasQuerySucceeded(sender, args) {
  var eachAreaName = "";
  var listItemEnumerator = allListItems3.getEnumerator();

  areaInput = document.getElementById('area-input');
  $("#area-input").empty()

  while (listItemEnumerator.moveNext()) {
    var areaListItem = listItemEnumerator.get_current();
    eachAreaName = areaListItem.get_item('Area').get_lookupValue();
    var option = document.createElement("option");
    option.value = areaListItem.get_item('Area').get_lookupValue();
    option.appendChild(document.createTextNode(eachAreaName));

    areaInput.appendChild(option);
  }

  var my_options = $("#area-input option");
  my_options.sort(function(a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $("#area-input").empty().append(my_options);
  $("#area-input").prepend("<option selected>Please Select...</option>");

}

function onSetAreasAllQuerySucceeded(sender, args) {
  var eachAreaName = "";
  var listItemEnumerator = allListItems3.getEnumerator();

  areaInput = document.getElementById('area-input');
  $("#area-input").empty()

  while (listItemEnumerator.moveNext()) {
    var areaListItem = listItemEnumerator.get_current();
    eachAreaName = areaListItem.get_item('Title');
    var option = document.createElement("option");
    option.value = areaListItem.get_item('Title');
    option.appendChild(document.createTextNode(eachAreaName));

    areaInput.appendChild(option);
  }

  var my_options = $("#area-input option");
  my_options.sort(function(a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $("#area-input").empty().append(my_options);
  $("#area-input").prepend("<option selected>Please Select...</option>");

}

function onSetAreasQueryFailed(sender, args) {
  alert('Area Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

// ======================= Set location dropdown menu ====================

function setOffices() {
  console.log("populating locations dropdown...");
  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit");
  officesList = clientContext.get_web().get_lists().getByTitle('Offices');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query></Query></View>'
  );
  this.allListItems4 = officesList.getItems(camlQuery);
  clientContext.load(allListItems4);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetOfficesQuerySucceeded),
    Function.createDelegate(this, this.onSetOfficesQueryFailed)
  );

}

function onSetOfficesQuerySucceeded(sender, args) {
  var eachOfficeName = "";
  var listItemEnumerator = allListItems4.getEnumerator();

  officeInput = document.getElementById('location-input');
  $("#location-input").empty();

  while (listItemEnumerator.moveNext()) {
    var officesListItem = listItemEnumerator.get_current();
    eachOfficeName = officesListItem.get_item('Title');
    var option = document.createElement("option");
    option.text = eachOfficeName;

    officeInput.add(option);
  }

  var my_options = $("#location-input option");
  my_options.sort(function(a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $("#location-input").empty().append(my_options);
  $("#location-input").prepend("<option selected>Please Select...</option>");
}

function onSetOfficesQueryFailed() {
  alert('Offices Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
  
}



//==================== triggers the Microsoft workflow to submit request ==================== 
function submitRequestFlow() {
	
try {
      managerComments = $('#man-comments').val();
      dataTemplate = '{ "id":' + R2RID + '" }'
      httpPostUrl = "https://prod-32.westeurope.logic.azure.com:443/workflows/16b4e8acfbe6435293b3078b09457a1c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9rwDXn9ZUyrej3QBVZBq_eGFy9np1EymNU83LknkbvU";

      //Call FormatRow function and replace with the values supplied in input controls. 
      var settings = {
        "async": true,
        "crossDomain": true,
        "url": httpPostUrl,
        "method": "POST",
        "headers": {
          "content-type": "application/json",
          "cache-control": "no-cache"
        },
        "processData": false,
        "data": dataTemplate
      }

      $.ajax(settings).done(function(response) {
        console.log("Successfully triggered the Microsoft Flow: R2R - Request Submitted");
      });

      setTimeout(function() {
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx';
      }, 2000); // redirects user to "DASHBOARD" 
	  
    } catch (e) {
      console.log("Error occurred in submitRequestFlow: " + e.message);
    }
	

}


//==================== triggers the Microsoft workflow to send e-mail ==================== 
function managerApprovalFlow(response) {

  if (admin === false) {
    $mainForm.hide();
    $confirmMsg.addClass('alert alert-warning').html('You are not authorised to approve this request.').show();
  } else {

    try {
	  d = new Date;
      managerComments = "From " + userName + " at " + d.toLocaleString() + ": " + $('#man-comments').val();
      dataTemplate = '{ "id":' + R2RID + ', "response":"' + response + '", "comments":"' + managerComments + '" }'
      httpPostUrl = "https://prod-89.westeurope.logic.azure.com:443/workflows/17d09a862e224e9b9886dfdb9e59036b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dKI3eMY9DyyWQW8-YODAkHGwpCXOSHapb3nvbK59Tg8";

      //Call FormatRow function and replace with the values supplied in input controls. 
      var settings = {
        "async": true,
        "crossDomain": true,
        "url": httpPostUrl,
        "method": "POST",
        "headers": {
          "content-type": "application/json",
          "cache-control": "no-cache"
        },
        "processData": false,
        "data": dataTemplate
      }

      $.ajax(settings).done(function(response) {
        console.log("Successfully triggered the Microsoft Flow: R2R - Manager Approval");
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx'; // redirects user to "DASHBOARD" 
      });

      // hide form 
      $mainForm.hide();
      $confirmMsg.addClass('alert alert-warning').html('Thank you. Your approval has been saved. There will be a slight delay of up to 3 mins before the booking will be shown on the dashboard. You will now be taken to the dashboard.').show();

    } catch (e) {
      console.log("Error occurred in managerApprovalFlow: " + e.message);
    }
  }
}

function baApprovalFlow(response) {
  if ((admin === false) && (!BAApprover.includes(userTitle))) {
    $mainForm.hide();
    $confirmMsg.addClass('alert alert-warning').html('You are not authorised to approve this request.').show();
  } else {
    try {
	  d = new Date;
      baComments = "From " + userName + " at " + d.toLocaleString() + ": " + $('#ba-comments').val();
      dataTemplate = '{ "id":' + R2RID + ', "response":"' + response + '", "comments":"' + baComments + '" }'
      httpPostUrl = "https://prod-108.westeurope.logic.azure.com:443/workflows/ca6aab21ba8c401e965ff18d63c3812f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TqSHooynmPTKUU2TgthLy_ypqAfwgPHZI7MVoYvkfLg";

      //Call FormatRow function and replace with the values supplied in input controls. 
      var settings = {
        "async": true,
        "crossDomain": true,
        "url": httpPostUrl,
        "method": "POST",
        "headers": {
          "content-type": "application/json",
          "cache-control": "no-cache"
        },
        "processData": false,
        "data": dataTemplate
      }

      $.ajax(settings).done(function(response) {
        console.log("Successfully triggered the Microsoft Flow: R2R - BA Approval");
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx'; // redirects user to "DASHBOARD" 
      });

      // hide form 
      $mainForm.hide();
      $confirmMsg.addClass('alert alert-warning').html('Thank you. Your approval has been saved. There will be a slight delay of up to 3 mins before your changes will be reflected on the dashboard. You will now be taken to the dashboard.').show();
    } catch (e) {
      console.log("Error occurred in baApprovalFlow: " + e.message);
    }
  }
}


function fdApprovalFlow(response) {
  if ((admin === false) && (!FDApprover.includes(userTitle))) {
	  if (stage != "RA") {
    $mainForm.hide();
	  $confirmMsg.addClass('alert alert-warning').html('You are not authorised to approve this request.').show();}
  } else {
    try {
	  d = new Date;
      fdComments = "From " + userName + " at " + d.toLocaleString() + ": " + $('#fd-comments').val();
      dataTemplate = '{ "id":' + R2RID + ', "response":"' + response + '", "comments":"' + fdComments + '" }'
      httpPostUrl = "https://prod-52.westeurope.logic.azure.com:443/workflows/cbfaa7366a1f4c10a0833d47b17a6850/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=XsmRauH2dKCCFj5iKI2jrhh5ukkCkJeiJs80w4cNhek";

      //Call FormatRow function and replace with the values supplied in input controls. 
      var settings = {
        "async": true,
        "crossDomain": true,
        "url": httpPostUrl,
        "method": "POST",
        "headers": {
          "content-type": "application/json",
          "cache-control": "no-cache"
        },
        "processData": false,
        "data": dataTemplate
      }

      $.ajax(settings).done(function(response) {
        console.log("Successfully triggered the Microsoft Flow: R2R - FD Approval");
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx'; // redirects user to "DASHBOARD" 
      });

      // hide form 
      $mainForm.hide();
      $confirmMsg.addClass('alert alert-warning').html('Thank you. Your approval has been saved. There will be a slight delay of up to 3 mins before your changes will be reflected on the dashboard. You will now be taken to the dashboard.').show();
    } catch (e) {
      console.log("Error occurred in fdApprovalFlow: " + e.message);
    }
  }
}


function tmApprovalFlow(response) {
  if ((admin === false) && (!TMApprover.includes(userTitle))) {
    $mainForm.hide();
    $confirmMsg.addClass('alert alert-warning').html('You are not authorised to approve this request.').show();
  } else {
    try {
	  d = new Date;
      tmComments = "From " + userName + " at " + d.toLocaleString() + ": " + $('#tm-comments').val();
      dataTemplate = '{ "id":' + R2RID + ', "response":"' + response + '", "comments":"' + tmComments + '" }'
      httpPostUrl = "https://prod-31.westeurope.logic.azure.com:443/workflows/ba13f335c0144aedb167ef1f14a69956/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tGu7YhxF_4xU7otnSt8Issm-nedhPsYWz-TWok46-0E";

      //Call FormatRow function and replace with the values supplied in input controls. 
      var settings = {
        "async": true,
        "crossDomain": true,
        "url": httpPostUrl,
        "method": "POST",
        "headers": {
          "content-type": "application/json",
          "cache-control": "no-cache"
        },
        "processData": false,
        "data": dataTemplate
      }

      $.ajax(settings).done(function(response) {
        console.log("Successfully triggered the Microsoft Flow: R2R - TM Approval");
        window.location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Pages/dashboard.aspx'; // redirects user to "DASHBOARD" 
      });

      // hide form 
      $mainForm.hide();
      $confirmMsg.addClass('alert alert-warning').html('Thank you. Your approval has been saved. There will be a slight delay of up to 3 mins before your changes will be reflected on the dashboard. You will now be taken to the dashboard.').show();
    } catch (e) {
      console.log("Error occurred in tmApprovalFlow: " + e.message);
    }
  }
}



function startCascade() {
  var cascadeArray = new Array();

  cascadeArray.push({
    parentFormField: "division", //Display name on form of field from parent list
    childList: "Auth Table", //List name of child list
    childLookupField: "Region", //Internal field name in Child List used in lookup
    childFormField: "region", //Display name on form of the child field
    parentFieldInChildList: "Division", //Internal field name in Child List of the parent field
    firstOptionText: "Please Select..."
  });

  /*
  cascadeArray.push({
  	parentFormField: "region", //Display name on form of field from parent list
  	childList: "Auth Table", //List name of child list
  	childLookupField: "Area", //Internal field name in Child List used in lookup
  	childFormField: "area", //Display name on form of the child field
  	parentFieldInChildList: "Region", //Internal field name in Child List of the parent field
  	firstOptionText: "Please Select..."
  });
  */
  //Set Divsion > Region > Area cascading drop down
  HillbillyCascade(cascadeArray[0]);
  //HillbillyCascade(cascadeArray[1]);


}


function HillbillyCascade(params) {

  var parent = $("select[Title='" + params.parentFormField + "'], select[Title='" + params.parentFormField + "']");

  $(parent).change(function() {
    DoHillbillyCascade(this.value, params);
  });

  var currentParent = $(parent).val();
  if (currentParent != 0) {
    DoHillbillyCascade(currentParent, params);
  }
}


function DoHillbillyCascade(parentID, params) {

  var child = $("select[Title='" + params.childFormField + "'], select[Title='" + params.childFormField + "']," + "select[Title='" + params.childFormField + "']");

  $(child).empty();

  var options = "";

  var call = $.ajax({
    url: "https://bauer.sharepoint.com/sites/MDS/_api/Lists/GetByTitle('Authorisation%20Table')/items?$select=Id,Division/Id,Region/Id&$expand=Division/Id,Region/Id&$filter=Division/Id%20eq%20%276%27",
    type: "GET",
    dataType: "json",
    headers: {
      Accept: "application/json;odata=verbose"
    }

  });
  call.done(function(data, textStatus, jqXHR) {

    for (index in data.d.results) {
      options += "<option value='" + data.d.results[index].Id + "'>" +
        data.d.results[index][params.childLookupField] + "</option>";
    }
    $(child).append(options);

  });
  call.fail(function(jqXHR, textStatus, errorThrown) {
    alert("Error retrieving information from list: " + params.childList + jqXHR.responseText);
    $(child).append(options);
  });

}