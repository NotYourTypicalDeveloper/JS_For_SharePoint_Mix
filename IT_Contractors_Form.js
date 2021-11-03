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
$(document).ready(function() {

    console.log('IT_Contractors_Form.js connected!');

    var terminDateInput = $('#terminationdate');
    infoHeader = $('#headerinfo'); // info about the contractor
    section1 = $('#section1'); // CONTINUE - renewal form
    section2 = $('#section2'); // TERMINATE - confirmation
    section3 = $('#section3'); // CONTINUE - confirmation 
    newDayRateInput = $('#newdayrate');
    endDateInputTermination = $('#contractenddate2'); // Termination screen - End date input



    URLResponseVal = getQueryVariable('response'); // checks value of parameter 'response' in URL
    URLIDVal = getQueryVariable('contractorid');

    var currUrl = window.location.href;

    // if URL contains response=continue, show the form
    if (URLResponseVal === 'Continue') {
        section2.hide(); // hides Termination messages
        infoHeader.add(section1).show();
    }

    if (URLResponseVal === 'Terminate') {
        infoHeader.add(section1).hide(); // hide header and continuation form
        section2.show(); // show termination confirm
    }

    // if(currUrl.includes('?')){ 
    //use of indexof for IE compatibility
    if (currUrl.indexOf('?')>1) {
        retrieveContractorsInfo();
    }


    // Calendar date picker for "Termination date" input - 'CONTINUE' form, when user clicks CONTINUE
    terminDateInput.daterangepicker({
        opens: 'left',
        autoUpdateInput: false,
        singleDatePicker: true,
        showDropdowns: true,
        locale: {
            format: 'DD-MM-YYYY',
            cancelLabel: 'Cancel'
        },
        minDate: new Date()
    },
    function(start) {
        terminDateFormatted = start.format('YYYY-MM-DD');
        console.log('Termination date: ' + terminDateFormatted);
        var nextRevDate = start.subtract(42, 'days').format('DD-MM-YYYY');
        var txt = 'Next review date is: ' + nextRevDate;
        $('#nextrevdatemsg').html(txt);
    });

    terminDateInput.on('apply.daterangepicker', function(ev, picker) {
        $(this).val(picker.startDate.format('DD-MM-YYYY'));
    });

    // validation form, if "DAY RATE" is not a number, show error
    function checkForm() {
        var newDayRateVal = $('#newdayrate').val();
        var newDayRateMsg = $('#newdayratemsg');
        if (isNaN(newDayRateVal)) {
            newDayRateInput.css('border', '2px solid red'); 
            newDayRateMsg.html("Please enter a number.");
        } else {
            newDayRateInput.css('border', '1px solid #e66b29');
            newDayRateMsg.html("");            
        }
    }
    newDayRateInput.change(checkForm);
    
});
///////////////////////////////////////////document ready ends ///////////////////////////////////////////

// ============== Retrieve Contractor's info from IT contractors List ============== 
    var targetListItem;

    // returns the context information about the current web application
    siteUrl = 'https://bauer.sharepoint.com/sites/Legal/ITContractors';
    clientContext = new SP.ClientContext(siteUrl);
    targetList = clientContext.get_web().get_lists().getByTitle('IT Contractors'); // our contractors list

    function retrieveContractorsInfo() {   
        targetListItem = targetList.getItemById(URLIDVal);
        clientContext.load(targetListItem);

        clientContext.executeQueryAsync(Function.createDelegate(this, this.retrOnQuerySucceeded), Function.createDelegate(this, this.retrOnQueryFailed));
    }
 
    function retrOnQuerySucceeded() {
        
       var contractorsName = targetListItem.get_item('Title');
       prevDayRate = targetListItem.get_item('Day_x0020_Rate');
       prevReviewDate = targetListItem.get_item('Review_x0020_Date');
       var prevReviewDateFormatted = String.format('{0:dd}-{0:MM}-{0:yyyy}',new Date(prevReviewDate));
       endDate = targetListItem.get_item('End_x0020_Date');
       var endDateFormatted = String.format('{0:dd}-{0:MM}-{0:yyyy}',new Date(endDate));
       var pageTitle = $("div[title='Contractors review']");
       var prevProjMoves = targetListItem.get_item('ProjectMoves');
       var prevJobDesc = targetListItem.get_item('JobDescription');

       // auto-populate the contractors info in title,header,form inputs,termination section

       var title =  "IT Contractor Review - " +  contractorsName;
       pageTitle.html(title); 
       $('#currdayrate').html(prevDayRate); 
       $('#lastrevrate').html(prevReviewDateFormatted);
       $('#prevenddate, #contractenddate').html(endDateFormatted);
       $('#contractorsname').html('&nbsp;' + contractorsName); // termination section
       $('#projectmoves').html(prevProjMoves);
       newDayRateInput.val(prevDayRate);
       $('#JDchanges').html(prevJobDesc);
        endDateInputTermination.val(endDateFormatted); // set "End date" input with current "End date" value (on termination screen)

        // Populates Calendar date picker for "Termination date" - 'TERMINATE' form. User can't select dates later than current "End Date", can't select dates before today.
        endDateInputTermination.daterangepicker({
            opens: 'left',
            autoUpdateInput: false,
            singleDatePicker: true,
            showDropdowns: true,
            locale: {
                format: 'DD-MM-YYYY',
                cancelLabel: 'Cancel'
            },
            minDate: new Date(),
            maxDate: new Date(endDate)
        },

        function(start) {
            terminDateFormatted = start.format('YYYY-MM-DD');
            console.log('new confirmed Termination date: ' + terminDateFormatted);

        });

        endDateInputTermination.on('apply.daterangepicker', function(ev, picker) {
            $(this).val(picker.startDate.format('DD-MM-YYYY'));
        });

    }
 
    function retrOnQueryFailed(sender, args) {
      alert('Request failed. \nError: ' + args.get_message() + '\nStackTrace: ' + args.get_stackTrace());
    }
    

// ============== IF RESPONSE = 'CONTINUE' / Update List Item once user clicks 'SUBMIT' ==============
function updateListItem() {

    console.log('updateListItem function is firing!');
    var newProjMovesVal = $('#projectmoves').val();
    newDayRateVal = $('#newdayrate').val();
    var newEndDateVal = $('#terminationdate').val();
    var jobDescVal = $('#JDchanges').val();
    var newNoticeVal = $('#newnoticeperiod option:selected').text();

    this.oListItem = targetList.getItemById(URLIDVal);

    // setting 'previous end date', 'previous rate', 'previous review date' columns 
    if (newProjMovesVal !== "") { oListItem.set_item('ProjectMoves', newProjMovesVal) };
    oListItem.set_item('PreviousEndDate', endDate);
    oListItem.set_item('PreviousReviewDate', prevReviewDate);
  
    // Setting the other columns with the new data from the form
    if(prevDayRate !== newDayRateVal) { oListItem.set_item('PreviousDayRate', prevDayRate) };
    if (newDayRateVal !== "" && $.isNumeric(newDayRateVal)) { oListItem.set_item('Day_x0020_Rate', newDayRateVal) }; // check that new day rate input is a number
    if(newEndDateVal !== "") {   	
        // Converting end date input value to a date object 
        var dateParts = newEndDateVal.split("-");
        // month is 0-based, that's why we need dataParts[1] - 1
        var newEndDateObject = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]); 
        oListItem.set_item('End_x0020_Date', newEndDateObject );
    }
    if (jobDescVal !== "") {oListItem.set_item('JobDescription', jobDescVal)};
    if (newNoticeVal !== "please select") {oListItem.set_item('NoticePeriod', newNoticeVal)};
    if (URLResponseVal === "Continue") {oListItem.set_item('ReviewOutcome', "Extended")};

    oListItem.update();

    clientContext.executeQueryAsync(Function.createDelegate(this, this.updateOnQuerySucceeded), Function.createDelegate(this, this.updateOnQueryFailed));
}

function updateOnQuerySucceeded() {
    console.log("Contactor info updated!");
    displaySuccessMsg();
    startMSFlow();
}

function updateOnQueryFailed(sender, args) {
    alert('Could not update contractor info. failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

// once user has submitted form, display 'Thank You' message
function displaySuccessMsg() {
    infoHeader.add(section1).hide();
    section3.fadeIn(2000).css('display', 'flex');
}


// ============== IF RESPONSE ='TERMINATE' - update columns ==============

function confirmEndDate() {
    $('#confirmEndDateDiv').hide();
    updTerminated(); // update item and starts MS workflow
    $('#thankyoumsg').fadeIn(1000).css('display', 'flex');
}

function updTerminated() {
    this.oListItem = targetList.getItemById(URLIDVal); // get ID of the item
    var newEndDateVal2 = $('#contractenddate2').val(); // grab the latest End Date that user has confirmed
    // Converting end date input value to a date object 
    var dateParts = newEndDateVal2.split("-");
    var newEndDateObject2 = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]); 

    // get current end date from the list to later set "Previous end date" column
    endDate = targetListItem.get_item('End_x0020_Date');

    // get today's date 
    var today = new Date();
    var todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), today.getHours(), today.getMinutes(), today.getSeconds(), today.getMilliseconds()); 

    // display the latest termination date in "Thank You" message 
    $('#contractenddate2msg').html('&nbsp;' + newEndDateVal2);

    // setting list columns
    oListItem.set_item('End_x0020_Date', newEndDateObject2 ); // set the new end date 
    oListItem.set_item('ReviewOutcome', "Terminated"); 
    oListItem.set_item('PreviousEndDate', endDate); // replace value of "previous end date" with the current value of "end date"
    oListItem.set_item('PreviousReviewDate', todayDate);
    // oListItem.set_item('Review_x0020_Date', null);

    // update columns
    oListItem.update();

    clientContext.executeQueryAsync(Function.createDelegate(this, this.updTerminatedOnQuerySucceeded), Function.createDelegate(this, this.updTerminatedOnQueryFailed));
}

function updTerminatedOnQuerySucceeded() {
    console.log('Review outcome updated to Terminated');
    startMSFlow();
}

function updTerminatedOnQueryFailed(sender, args) {
    alert('update Termination failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}



//==================== triggers the Microsoft workflow to send e-mail ==================== 
function startMSFlow() { 

    try { 

        var dataTemplate = '{\r\n "id": ' + URLIDVal + '\r\n}'; 
        var httpPostUrl;
        
        // depending on response paramater in the URL, trigger the "Extend" or "Terminate" workflow
        if (URLResponseVal === 'Continue') { 
            httpPostUrl= "https://prod-184.westeurope.logic.azure.com:443/workflows/43dca8f02c784dc1a00af89f09448427/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=DHHVDOda4hHI71TnQQ5T6U3UMsHQYsA5jmVnmSuMtfc"; 
        }
        if (URLResponseVal === 'Terminate') {
            httpPostUrl = 'https://prod-109.westeurope.logic.azure.com:443/workflows/b65db21984b14962a418bf1493b5c2ac/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EvaV4UZtFFptm4GWinCRcouegkdcumlcg71cEPMvCfw'
        }

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
            console.log("Successfully triggered the Microsoft Flow."); 
        }); 

    } 

    catch (e) { 
        console.log("Error occurred in StartMicrosoftFlowTriggerOperations " + e.message); 
    } 

}

