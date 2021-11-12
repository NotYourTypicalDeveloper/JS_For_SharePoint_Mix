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

/////////////////////////////////////////// document ready /////////////////////////////////////////////

$(document).ready(function () {
  console.log("req-to-rec-form.js is connected!");

  // ::::::::::::::::::::::::::::::: VARIABLES ::::::::::::::::::::::::::::::::::::: 
  // Input selectors
  $divisionInput = $('#division-input');
  $regionInput = $('#region-input');
  $areaInput = $('#area-input');
  $recManagerInput = $('#rec-manager-input');
  $lineManagerInput = $('#line-manager-input');

  $jobRoleInput = $('#job-role-input');
  $locationInput = $('#location-input');
  $startDateInput = $('#start-date-input');
  $existingRoleInput = $('#exist-role-input');
  $methodInput = $('#method-input');
  $FPTimeInput = $('#FP-time-input');
  $permTempInput = $('#perm-temp-input');
  $maternityInput = $('#maternity-input');
  $endDateInput = ('#end-date-input');
  $budgCostInput = $('#budg-cost-input');
  $bizCaseInput = $('#bizcase-input');
  $fileInput = $('#fileinput'); // job description upload

  $leaverNameInput = $('#leaver-name-input');
  $leaverJobInput = $('#leaver-JT-input');
  $leaverSalaryInput = $('#leaver-salary-input');
  $leaverReasonInput = $('#leaving-reason-input');

  $salaryInput = $('#annual-salary-input');
  $pensionInput = $('#pension-input');
  $bonusInput = $('#bonus-input');
  $bonusPercentInput = $('#bonus-percent-input');
  $carInput = $('#car-input');
  $carAllowInput = $('#car-allow-input');
  $medicalInput = $('#medical-input');
  $ourTeamInput = $('#our-team-input');
  $roleFocusInput = $('#role-focus-input');
  $whatYouDoInput = $('#what-you-do-input');
  $whatYouBringInput = $('#what-you-bring-input');
  $howApplyInput = $('#how-apply-input');
  $itwDeadlineInput = $('#itw-deadline-input');
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
  lineManagerEmail = '';
  setDivisions();

  if (!(stage)) {

    // load regions on division change
    $divisionInput.change(function () {
      setRegions();
    });

    // load areas on region change
    $divisionInput.change(function () {
      setAreas();
    });

  }


  // enable Bootstrap tooltips
  $('[data-toggle="tooltip"]').tooltip({
    placement: 'bottom'
  });

  // set typeahead for Rec & Line Managers 
  setTypeAhead();

  // validation jQuery
  $mainForm.validate({
    rules: {
      division: "required",
      region: "required",
      area: "required",
      recmanager: "required",
      linemanager: "required",
      jobrole: "required",
      location: "required",
      startdate: "required",
      existingrole: "required",
      recmethod: "required",
      FPTime: "required",
      permtemp: "required",
      matercov: "required",
      enddate: "required",
      budgetcost: "required",
      bizcase: "required",
      fileinput: "required",
      leavername: "required",
      leaverjob: "required",
      leaversalary: "required",
      reason: "required",
      annualsalary: "required",
      companycar: "required",
      medicalinsurance: "required",
      ourteam: "required",
      rolefocus: "required",
      whatyoudo: "required",
      whatyoubring: "required",
    },

    messages: {
      division: "Please Select... a Division.",
      region: "Please Select... a Region.",
      area: "Please Select... an Area",
      recmanager: "Please enter a recruiting manager name",
      linemanager: "Please enter a line manager name",
      jobrole: "Please enter a job role",
      location: "Please Select... a location",
      startdate: "Please Select... a date in the calendar",
      existingrole: "Please confirm if the role is already existing",
      recmethod: "Please Select... a recruitment method",
      FPTime: "Please confirm if the contract is full or part time",
      permtemp: "Please confirm if the contract is permanent or temporary",
      matercov: "Please confirm if it is a maternity cover",
      enddate: "Please enter the expected end date",
      budgetcost: "Please enter the budgeted cost of recruitment",
      bizcase: "Please provide a business case",
      fileinput: "Please attach a job description",
      leavername: "Please enter the name of the person who left",
      leaverjob: "Please enter the leaver's job title",
      leaversalary: "Please enter the leaver's current salary",
      reason: "Please enter a reason for leaving",
      annualsalary: "Please enter the annual salary offered for the position",
      medicalinsurance: "Please Select... the type of medical insurance",
      ourteam: "Please enter a text about your team",
      rolefocus: "Please describe the role",
      whatyoudo: "Please describe the duties",
      whatyoubring: "Please enter a text",
    },

    submitHandler: function () {
      submitForm();
    },

    invalidHandler: function (event, validator) {
      var errorDiv = $('#error-msg');
      var errors = validator.numberOfInvalids();
      if (errors) {
        errorDiv.html('Please check errors').show();
      } else {
        errorDiv.hide();
      }
    }
  });




  //startCascade();

  //READ ONLY / View
  if (stage === "view") {
    populateForm(R2RID);
    setTimeout(function () {
      $(':input').attr('readonly', 'readonly')
    }, 1000);
    //hide submit & save draft buttons
    $('#save-btn').hide();
    $('#submit-btn').hide();
  }

  //Edit
  if (stage === "edit") {
    setTimeout(function () {
      populateForm(R2RID);


    }, 1000);

  }

  //Manager Approval
  if (stage === "MA") {

    populateForm(R2RID);
    // load regions on division change
    $divisionInput.change(function () {
      setRegions();
    });

    // load areas on region change
    $divisionInput.change(function () {
      setAreas();
    });

    setTimeout(function () {
      $(':input').attr('readonly', 'readonly')
    }, 1000);

    //hide submit & save draft buttons
    $('#save-btn').hide();
    $('#submit-btn').hide();

    //show approve / reject buttons for manager
    $('#approval-banner').show();
    document.getElementById('approval-banner').innerHTML = "Manager Approval"
    $('#man-label').show();
    $('#man-comments').show();
    $('#man-approve').show();
    $('#man-reject').show();
    setTimeout(function () {
      $('#man-comments').removeAttr("readonly");
    }, 1500);
  }


  //BA Approval
  if (stage === "BA") {
    populateForm(R2RID);

    // load regions on division change
    $divisionInput.change(function () {
      setRegions();
    });

    // load areas on region change
    $divisionInput.change(function () {
      setAreas();
    });

    setTimeout(function () {
      $(':input').attr('readonly', 'readonly')
    }, 1000);

    //hide submit & save draft buttons
    $('#save-btn').hide();
    $('#submit-btn').hide();

    //show approve / reject buttons for ba
    document.getElementById('approval-banner').innerHTML = "BA Approval"
    $('#approval-banner').show();
    $('#man-label').show();
    $('#man-comments').show();
    $('#ba-label').show();
    $('#ba-comments').show();
    $('#ba-approve').show();
    $('#ba-reject').show();
    setTimeout(function () {
      $('#ba-comments').removeAttr("readonly");
    }, 3000);
  }

  // _____________-_-_EVENT LISTENERS_-_-______________

  // "salary input", calculate "pension" field
  $salaryInput.change(function () {
    calculatePension();
  });

  // toggle "Leaver" section when select 'YES'
  $existingRoleInput.change(function () {
    selectedOption = $('#exist-role-input option:selected');
    toggleSections($leaverSection, selectedOption, "Yes");
  });

  // toggle "Maternity cover" & "End date" section
  $permTempInput.change(function () {
    selectedOption = $('#perm-temp-input option:selected');
    toggleSections($materSection, selectedOption, "Temporary");
    toggleSections($endDateSection, selectedOption, "Temporary")
  });

  // toggle "Bonus %" section, when Bonus is YES
  $bonusInput.change(function () {
    selectedOption = $('#bonus-input option:selected');
    toggleSections($bonusPercentSection, selectedOption, "Yes")
  });

  // round up to integer if user types decimal number - Budgeted recruitment + Annual salary input
  $budgCostInput.change(function () {
    roundINT($budgCostInput);
  });
  $salaryInput.change(function () {
    roundINT($salaryInput);
  });
  $leaverSalaryInput.change(function () {
    roundINT($leaverSalaryInput);
  });

  // maximum 2 decimals - Bonus % input
  $bonusPercentInput.change(function () {
    round2NumberDecimal($bonusPercentInput);
  });

  // on change, display file name as inner text in the "upload" button
  $fileInput.change(function () {
    displayFileNames();
  });

  //multiple select checkboxes for method of recruitment
  $methodInput.dropdownchecklist({
    forceMultiple: true
  });
  document.querySelector("#ddcl-method-input > span > span").innerHTML = "Please Select...";
  document.querySelector("#ddcl-method-input-ddw > div").style.height = "auto";
  document.querySelector("#method-input").classList.add("form-control");
  document.querySelector("#ddcl-method-input > span").style.width = "200px";


});
///////////////////////////////////////////document ready ends ///////////////////////////////////////////


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
    EnableManagerypeAhead($lineManagerInput);

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
    datumTokenizer: function (d) {
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
  }).blur(function () {});

  target.on('typeahead:selected typeahead:autocompleted', function (obj, data) {
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
    $lineManagerInput.val(managerSurname + ', ' + managerForename) //.focus().trigger('keyup')
  });
}

function EnableManagerypeAhead(target) {
  var currentEmployeeID = '';
  var currentEmployeeName = '';

  // Constructing the suggestion engine
  employees = new Bloodhound({
    datumTokenizer: function (d) {
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
  }).blur(function () {});

  target.on('typeahead:selected typeahead:autocompleted', function (obj, data) {
    //target.val(data['DisplayName']);
    lineManagerEmail = data['EmailAddress'];
  });
}


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
  var fileInputVal = fileInput.val();
  var updatedStr = fileInputVal.replace("C:\\fakepath\\", "");
  $('#fileinput-label > span').html(updatedStr); // update upload btn inner text
}

// ============================ POPULATE FORM - retrieve the request info from 'Requests' list ==============================

function populateForm(itemID) {
  console.log("populating form...");

  setRegionsAll();
  setAreasAll();

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/");
  targetList = clientContext.get_web().get_lists().getByTitle('Requests');
  targetListItem = targetList.getItemById(itemID);
  clientContext.load(targetListItem);

  clientContext.executeQueryAsync(Function.createDelegate(this, this.requestOnQuerySucceeded), Function.createDelegate(this, this.requestOnQueryFailed));
}

function requestOnQuerySucceeded(sender, args) {

  // Grab existing values of the related item 
  var division = targetListItem.get_item('Division');
  var region = targetListItem.get_item('Region');
  var area = targetListItem.get_item('Area');
  var recManager = targetListItem.get_item('RecruitingManagerText');
  var lineManager = targetListItem.get_item('LineManagerText');

  var jobRole = targetListItem.get_item('JobRole');
  var location = targetListItem.get_item('Location');
  var startDate = targetListItem.get_item('AnticipatedStartDate');
  var existRole = targetListItem.get_item('ExistingRole');
  var recMeth = targetListItem.get_item('MethodOfRecruitment');
  var FPTime = targetListItem.get_item('FullTimePartTime');
  var permTemp = targetListItem.get_item('PermanentTemporary');
  var materCov = targetListItem.get_item('MaternityCover');
  var endDate = targetListItem.get_item('ExpectedEndDate');
  var budgCost = targetListItem.get_item('BudgetedCostOfRecruitment');
  var bizCase = targetListItem.get_item('BusinessCase');

  var leaverName = targetListItem.get_item('NameOfLeaver');
  var leaverJob = targetListItem.get_item('LeaverJobTitle');
  var leaverSal = targetListItem.get_item('LeaversSalary');
  var leaverReason = targetListItem.get_item('ReasonForLeaving');

  var annualSalary = targetListItem.get_item('AnnualSalary');
  var pension = targetListItem.get_item('PensionContribution');
  var bonus = targetListItem.get_item('Bonus');
  var bonusPercent = targetListItem.get_item('Bonus_x0025_');
  var compCar = targetListItem.get_item('CompanyCar');
  var carAllow = targetListItem.get_item('CarAllowance');
  var medical = targetListItem.get_item('CompayPaidMedicalInsurance');
  var ourTeam = targetListItem.get_item('OurTeam');
  var roleFocus = targetListItem.get_item('RoleFocus');
  var whatYouDo = targetListItem.get_item('WhatYouDo');
  var whatBring = targetListItem.get_item('WhatYouBring');
  var howApply = targetListItem.get_item('HowToApply');
  var decisionItvwDate = targetListItem.get_item('DecisionInterviewDate');
  var closingDate = targetListItem.get_item('ApplicationClosingDate');

  // convert date formats to 'yyyy-MM-dd"


  // set fields with these values
  $divisionInput.val(division);
  $areaInput.val(area);
  $recManagerInput.val(recManager);
  $lineManagerInput.val(lineManager);

  $jobRoleInput.val(jobRole);
  $locationInput.val(location);
  $startDateInput.val(startDate)
  console.log('start date ===> ', startDate, typeof (startDate));

  if (existRole) {
    $existingRoleInput.val('Yes');
    $leaverSection.show();
  } else {
    $existingRoleInput.val('No');
  }

  $methodInput.val(recMeth);
  $FPTimeInput.val(FPTime);
  $permTempInput.val(permTemp);
  $budgCostInput.val(budgCost);
  $bizCaseInput.val(bizCase);

  $leaverNameInput.val(leaverName);
  $leaverJobInput.val(leaverJob);
  $leaverSalaryInput.val(leaverSal);
  $leaverReasonInput.val(leaverReason);

  $salaryInput.val(annualSalary);
  $pensionInput.val(pension);
  bonus ? $bonusInput.val('Yes') : $bonusInput.val('No');
  $bonusPercentInput.val(bonusPercent);
  compCar ? $carInput.val('Yes') : $carInput.val('No');
  $carAllowInput.val(carAllow);
  $medicalInput.val(medical);

  $ourTeamInput.val(ourTeam);
  $roleFocusInput.val(roleFocus);
  $whatYouDoInput.val(whatYouDo);
  $whatYouBringInput.val(whatBring);
  $howApplyInput.val(howApply);
  $itwDeadlineInput.val(decisionItvwDate);
  $closingDateInput.val(closingDate);

  console.log('bonus ===> ', bonus);

  // execute after division has loaded
  setTimeout(function () {
    $regionInput.val(region);
  }, 1000);


  /*       document.getElementById('division-input').value = targetListItem.get_item('Division');
     $('#division-input').change();
       document.getElementById('region-input').value = targetListItem.get_item('Region');
     $('#region-input').change();
       document.getElementById('area-input').value = targetListItem.get_item('Area');
     $('#area-input').change();
     $('#rec-manager-input').val(targetListItem.get_item('RecruitingManagerText'));
       $('#line-manager-input').val(targetListItem.get_item('LineManagerText'));
       
     $('#job-role-input').val(targetListItem.get_item('JobRole'));
       document.getElementById('location-input').value = targetListItem.get_item('Location');
     $('#location-input').change();
     $('#start-date-input').val(targetListItem.get_item('AnticipatedStartDate')); // "yyyy-MM-dd"
       document.getElementById('exist-role-input').value = targetListItem.get_item('ExistingRole');
       $('#exist-role-input').change();
       document.getElementById('method-input').value = targetListItem.get_item('MethodOfRecruitment');
       $('#method-input').change();
       
     document.getElementById('FP-time-input').value = targetListItem.get_item('FullTimePartTime');
     $('#FP-time-input').change();
     document.getElementById('perm-temp-input').value = targetListItem.get_item('PermanentTemporary');
       $('#perm-temp-input').change();
       
       $('#budg-cost-input').val(targetListItem.get_item('BudgetedCostOfRecruitment'));
                 
       $('#bizcase-input').val(targetListItem.get_item('BusinessCase'));
       
     $('#annual-salary-input').val(targetListItem.get_item('RecruitingManagerText')); */




}


function requestOnQueryFailed(sender, args) {
  alert('Failed to populate form data. ' + args.get_message() + '\n' + args.get_stackTrace());
}


// ============================  🆂🆄🅱🅼🅸🆃 🅵🅾🆁🅼 ===========================

function submitForm() {
  console.log('____-_-_Submitting form starting_-_-___');

  draftMode = true;
  submitMode = true;

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


// ========================== 🅲🆁🅴🅰🆃🅴 🅽🅴🆆 🅸🆃🅴🅼 ==========================
siteUrl = '/sites/UK-HR-RequestToRecruit/';

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

  // remove an readonly attributes that could block item creation
  // if ($('input').is('[readonly]')) { $('input').removeAttr('readonly'); }
  // $('input').removeAttr('type');

  try {
    $spinner.show();

    divisionVal = $('#division-input option:selected').val(); // grab value of the selected option
    regionVal = $('#region-input option:selected').val();
    areaVal = $('#area-input option:selected').val();
    recManagerVal = $recManagerInput.val();
    lineManagerVal = $lineManagerInput.val();

    jobRoleVal = $jobRoleInput.val();
    locationVal = $('#location-input option:selected').val();
    startDateVal = startDateInput.val();
    existingRoleVal = $('#exist-role-input option:selected').val();
    //methodRecVal = $('#method-input option:selected').val();
    methodRecVal = document.querySelector("#ddcl-method-input > span > span").innerText

    FPTimeVal = $('#FP-time-input option:selected').val();
    contractTypeVal = $('#perm-temp-input option:selected').val();
    materVal = $('#maternity-input option:selected').val();
    endDateVal = $('#end-date-input').val();
    budgCostVal = budgCostInput.val();
    bizCaseVal = $('#bizcase-input').val();

    salaryVal = salaryInput.val();
    pensionVal = $('#pension-input').val();
    bonusVal = $('#bonus-input option:selected').val();
    bonusPercentVal = $('#bonus-percent-input').val();
    carVal = $('#car-input option:selected').val();
    carAllowVal = $('#car-allow-input').val();
    medicalVal = $('#medical-input option:selected').val();

    leaverVal = $('#leaver-name-input').val();
    leaverJobVal = $('#leaver-JT-input').val();
    leaverSalaryVal = leaverSalaryInput.val();
    leaverReasonVal = $('#leaving-reason-input').val();

    ourTeamVal = $('#our-team-input').val();
    roleFocusVal = $('#role-focus-input').val();
    whatYouDoVal = $('#what-you-do-input').val();
    whatYouBringVal = $('#what-you-bring-input').val();
    how2ApplyVal = $('#how-apply-input').val();
    applyDateVal = $('#appli-closing-input').val();

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

    if (!!lineManagerVal) {
      oListItem.set_item('LineManagerText', lineManagerVal)
    };

    if (!!lineManagerEmail) {
      spUser = SP.FieldUserValue.fromUser(lineManagerEmail);
      oListItem.set_item('LineManager', spUser);
    };

    if (!!jobRoleVal) {
      oListItem.set_item('JobRole', jobRoleVal)
    };

    if (!!locationVal) {
      oListItem.set_item('Location', locationVal)
    };

    if (!!startDateVal) {
      // Converting end date input value to a date object 
      var dateParts = startDateVal.split("-");
      // month is 0-based, that's why we need dataParts[1] - 1
      var startDateObj = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
      oListItem.set_item('AnticipatedStartDate', startDateObj);
    }

    if (!!existingRoleVal) {
      oListItem.set_item('ExistingRole', existingRoleVal)
    };
    if (!!methodRecVal) {
      var methodRec = methodRecVal.split(', ')
      oListItem.set_item('MethodOfRecruitment', methodRec);
    };
    if (!!FPTimeVal) {
      oListItem.set_item('FullTimePartTime', FPTimeVal)
    };

    if (!!contractTypeVal) {
      oListItem.set_item('PermanentTemporary', contractTypeVal)
    };

    if (!!materVal) {
      oListItem.set_item('MaternityCover', materVal);
    }

    if (!!endDateVal) {
      // Converting end date input value to a date object 
      var dateParts = endDateVal.split("-");
      var endDateObj = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
      oListItem.set_item('ExpectedEndDate', endDateObj);
    }

    if (!!budgCostVal) {
      oListItem.set_item('BudgetedCostOfRecruitment', budgCostVal);
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

    if (!!medicalVal) {
      oListItem.set_item('CompayPaidMedicalInsurance', medicalVal);
    }

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

    if (!!applyDateVal) {
      oListItem.set_item('ApplicationClosingDate', applyDateVal);
    }
    console.log('draft status', draftMode);
    console.log('submit status', submitMode);

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
  if (stage === 'edit') {} else {
    clientContext.load(oListItem);
  }
  clientContext.executeQueryAsync(Function.createDelegate(this, this.createonQuerySucceeded), Function.createDelegate(this, this.createonQueryFailed));
}

// upload the attached file after the item is created
function createonQuerySucceeded() {
  try {
    // if there is a file uploaded, upload the file
    if (!!fileInput) {
      uploadTheFile();
    } else {

      // show relevant confirmation message
      if (submitMode === true) {
        $confirmMsg.addClass('alert alert-success').html('Thank you for submitting the request. You can now close your browser.').show();
      }
      if (draftMode === true) {
        $confirmMsg.addClass('alert alert-warning').html('Your request has been saved as a Draft. You can now close your browser.').show();
      }

      //alert('Item created');
      //location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Lists/Requests/'; // redirects user to "REQUESTS" list
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
  currentID = oListItem.get_id();
  var listTitle = 'Requests';
  var itemId = currentID;
  var fileInputElem = document.getElementById("fileinput");
  var file = fileInputElem.files[0];


  processUpload(file, listTitle, itemId,
    function () {
      $spinner.hide();
      alert('Item created and file uploaded successfully.');
      //location.href = 'https://bauer.sharepoint.com/sites/UK-HR-RequestToRecruit/Lists/Requests/';// redirects user to "REQUESTS" list
      if (submitMode === true) {
        $confirmMsg.addClass('alert alert-success').html('Thank you for submitting the request. You can now close your browser.').show();
      }
      if (draftMode === true) {
        $confirmMsg.addClass('alert alert-warning').html('Your request has been saved as a Draft. You can now close your browser.').show();
      }
    },
    function (sender, args) {
      $spinner.hide();
      alert('Item created, but file not uploaded. Please try again or contact an administrator');
    }
  );
}


function processUpload(fileInputElem, listTitle, itemId, success, error) {
  try {
    var reader = new FileReader();
    reader.onload = function (result) {
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
    function (folder) {
      var attachmentFolderUrl = folder.get_serverRelativeUrl();
      uploadFile(attachmentFolderUrl, fileName, fileContent, success, error);
    },
    error);
}


function ensureAttachmentFolder(listTitle, itemId, success, error) {
  var ctx = SP.ClientContext.get_current();
  var web = ctx.get_web();
  var list = web.get_lists().getByTitle(listTitle);
  ctx.load(list, 'RootFolder');
  var item = list.getItemById(itemId);
  ctx.load(item);
  ctx.executeQueryAsync(
    function () {
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
        function () {
          success(attachmentsFolder);
        },
        error);
    },
    error);
}


function uploadFile(folderUrl, fileName, fileContent, success, error) {
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
  divisionsList = clientContext.get_web().get_lists().getByTitle('Divisions');
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
  my_options.sort(function (a, b) {
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
  divisionsList = clientContext.get_web().get_lists().getByTitle('Authorisation Table');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query><Where><Eq><FieldRef Name=\'Division\'/>' +
    '<Value Type=\'Lookup\'>' + divisionVal + '</Value></Eq></Where></Query></View>'
  );
  this.allListItems2 = divisionsList.getItems(camlQuery);
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
  divisionsList = clientContext.get_web().get_lists().getByTitle('Regions');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query></Query></View>'
  );
  this.allListItems2 = divisionsList.getItems(camlQuery);
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
    eachRegionName = regionListItem.get_item('Region').get_lookupValue();
    var option = document.createElement("option");
    option.value = regionListItem.get_item('Region').get_lookupValue();
    option.appendChild(document.createTextNode(eachRegionName));

    $divisionInput.appendChild(option);
  }
  var my_options = $("#region-input option");
  my_options.sort(function (a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
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
  my_options.sort(function (a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $regionInput.empty().append(my_options);
  $regionInput.prepend("<option selected>Please Select...</option>");

}

function onSetRegionsQueryFailed() {
  alert('Region Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

// ======================= Set area dropdown menu ====================

function setAreas() {
  console.log("populating area dropdown...");
  regionVal = $('#region-input option:selected').val();

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  divisionsList = clientContext.get_web().get_lists().getByTitle('Authorisation Table');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query><Where><And><Eq><FieldRef Name=\'Division\'/>' +
    '<Value Type=\'Lookup\'>' + divisionVal + '</Value></Eq><Eq><FieldRef Name=\'Region\'/>' +
    '<Value Type=\'Lookup\'>' + regionVal + '</Value></Eq></Where></And></Query></View>'
  );
  this.allListItems3 = divisionsList.getItems(camlQuery);
  clientContext.load(allListItems3);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetAreasQuerySucceeded),
    Function.createDelegate(this, this.onSetAreasQueryFailed)
  );

}

function setAreasAll() {
  console.log("populating area dropdown with all...");

  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/MDS");
  divisionsList = clientContext.get_web().get_lists().getByTitle('Area');
  var camlQuery = new SP.CamlQuery();
  camlQuery.set_viewXml(
    '<View><Query></Query></View>'
  );
  this.allListItems3 = divisionsList.getItems(camlQuery);
  clientContext.load(allListItems3);
  clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onSetAreasAllQuerySucceeded),
    Function.createDelegate(this, this.onAreasAllQueryFailed)
  );

}

function onSetAreasQuerySucceeded(sender, args) {
  var eachAreasName = "";
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
  my_options.sort(function (a, b) {
    if (a.text > b.text) return 1;
    else if (a.text < b.text) return -1;
    else return 0
  })
  $("#area-input").empty().append(my_options);
  $("#area-input").prepend("<option selected>Please Select...</option>");

}

function onSetAreasAllQuerySucceeded(sender, args) {
  var eachAreasName = "";
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
  my_options.sort(function (a, b) {
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
  var clientContext = new SP.ClientContext("https://bauer.sharepoint.com/sites/IT/Equipment");
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
  my_options.sort(function (a, b) {
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




//==================== triggers the Microsoft workflow to send e-mail ==================== 
function managerApprovalFlow(response) {
  try {
    managerComments = $('#man-comments').val();
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

    $.ajax(settings).done(function (response) {
      console.log("Successfully triggered the Microsoft Flow: R2R - Manager Approval");
    });

    // hide form 
    $mainForm.hide();
    $confirmMsg.addClass('alert alert-warning').html('Thank you. Your approval has been saved. You can now close your browser.').show();

  } catch (e) {
    console.log("Error occurred in managerApprovalFlow: " + e.message);
  }

}

function baApprovalFlow(response) {
  try {
    baComments = $('#ba-comments').val();
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

    $.ajax(settings).done(function (response) {
      console.log("Successfully triggered the Microsoft Flow: R2R - BA Approval");
    });

    // hide form 
    $mainForm.hide();
    $confirmMsg.addClass('alert alert-warning').html('Thank you. Your approval has been saved. You can now close your browser.').show();
  } catch (e) {
    console.log("Error occurred in baApprovalFlow: " + e.message);
  }

}

function startCascade() {
  var cascadeArray = new Array();

  cascadeArray.push({
    parentFormField: "division", //Display name on form of field from parent list
    childList: "Authorisation Table", //List name of child list
    childLookupField: "Region", //Internal field name in Child List used in lookup
    childFormField: "region", //Display name on form of the child field
    parentFieldInChildList: "Division", //Internal field name in Child List of the parent field
    firstOptionText: "Please Select..."
  });


  //Set Divsion > Region > Area cascading drop down
  HillbillyCascade(cascadeArray[0]);


}


function HillbillyCascade(params) {

  var parent = $("select[Title='" + params.parentFormField + "'], select[Title='" + params.parentFormField + "']");

  $(parent).change(function () {
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
  call.done(function (data, textStatus, jqXHR) {

    for (index in data.d.results) {
      options += "<option value='" + data.d.results[index].Id + "'>" +
        data.d.results[index][params.childLookupField] + "</option>";
    }
    $(child).append(options);

  });
  call.fail(function (jqXHR, textStatus, errorThrown) {
    alert("Error retrieving information from list: " + params.childList + jqXHR.responseText);
    $(child).append(options);
  });

}