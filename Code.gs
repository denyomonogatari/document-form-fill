/* What should the add-on do after it is installed */
function onInstall() {
  onOpen();
}

/* What should the add-on do when a document is opened */
function onOpen() {
  DocumentApp.getUi()
  .createAddonMenu() // Add a new option in the Google Docs Add-ons Menu
  .addItem("Start", "showSidebar")
  .addToUi();  // Run the showSidebar function when someone clicks the menu
}

/* Show a 300px sidebar with the HTML from googlemaps.html */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile("missingfields")
    .evaluate()
    .setTitle("Fill missing fields"); // The title shows in the sidebar
  DocumentApp.getUi().showSidebar(html);
  
}

function getFields() {
  var findMe = "[$]{[^${}()]+}([(]([^${}(),]+,)*[^${}(),]+[)]){0,1}"; // tag regex for google doc
  var body = DocumentApp.getActiveDocument().getBody();
  var foundElement = body.findText(findMe); // gets ENTIRE element that has a string that matches findMe
  
  var fields = [];
  while (foundElement != null) {
    var rangeElementType = foundElement.getElement().getType();
    var isPartial = foundElement.isPartial();
    var startOffset = foundElement.getStartOffset();
    var endOffset = foundElement.getEndOffsetInclusive();
    
    if (isPartial && (rangeElementType == DocumentApp.ElementType.TEXT)) {
      // entire string of element that that has a string that matches findMe
      var foundElementString = foundElement.getElement().asText().getText();
      
      // the string we actually were looking for
      var foundString = foundElementString.substring(startOffset, endOffset+1);
      
      var field = foundString.substring(foundString.search('{')+1, foundString.search('}'));
      
      // get options if they exist
      var startOptionIndex = foundString.search('[(]');
      var endOptionIndex = foundString.search('[)]');
      if (startOptionIndex === -1 || endOptionIndex === -1) {
        fields.push({field: field, options: []});
      }
      else {
        startOptionIndex = startOptionIndex + 1;
        var options = foundString.substring(startOptionIndex, endOptionIndex).split(',');
        fields.push({field: field, options: options});
      }
      
      // Change the background color to yellow
      foundElement.getElement().asText().setBackgroundColor(startOffset, endOffset, "#FCFC00");
    }
    
    foundElement = body.findText(findMe, foundElement);
  }
  
  return fields;
}

function fill(toFill) {
  var toReplace = [];
  for (var i = 0; i < toFill.length; i++) {
    var tag = '[$]{1}[{]{1}('+toFill[i].tag+')[}]{1}';
    var inputType = toFill[i].inputType;
    var value = toFill[i].value;
    
    var textToReplace = "" + tag
    if (toFill[i].options.length > 0) {
      textToReplace = textToReplace.concat("[(]{1}");
      textToReplace = textToReplace.concat(toFill[i].options.join());
      textToReplace = textToReplace.concat("[)]{1}");
    }
    
    toReplace.push({textToReplace: textToReplace, value: value});
  }
  
  var j = 0;
  var body = DocumentApp.getActiveDocument().getBody();
  var searchResult = body.findText(toReplace[j].textToReplace, null);
  
  while(j < toReplace.length && searchResult) {
    searchResult.getElement().replaceText(toReplace[j].textToReplace, toReplace[j].value);
    j = j + 1;
    searchResult = body.findText(toReplace[j].textToReplace, null);
  }
}
