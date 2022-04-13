// Get values of key "company code" from an array object
const compCodesArr = jsonCompCodes.map((elem) => elem.CompanyCode);

// remove duplicates numbers
const uniqueCompCodes = [...new Set(compCodesArr)];

// sort by ascending order
uniqueCompCodes.sort((a, b) => a - b); 

// Reusable function Toggle hide or show of certain sections depending on which option is selected
    function showHide(selectedOption, choice, section, displayType) {
      if (selectedOption === choice) {
        section.css("display", displayType);
      } else {
        section.hide();
      }
    }
