/**
 * The allowable shorthand notation.
 * nf = No Face (no armor), location 1
 * pf = Partial Face (partial armor), location 1
 * ng = No Groin (no armor), location 26
 * fo = Front Only 
 */
var allowableShorthand = [
  'nf', 'pf', 'ng', 'fo',
];

/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

function onInstall(e) {
  onOpen(e);
}

/**
 * Runs whenever the spreadsheet is open
 */
function onOpen(e) {
  /*
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
      { name: 'Fill in Armor Table', functionName: 'fillArmorTable'}
    ];
    spreadsheet.addMenu('Character Sheet', menuItems);
    */
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Dragonfire')
    .addItem('Fill in Armor Table', 'fillArmorTable')
    .addToUi();
}

/**
 * Fill the armor table of the character sheet with values
 * extracted from the Armor table
 */
function fillArmorTable() {
  console.log('------------armor Fill table start');
  var cellLocation = {
    armorTop: 111,
    armorLeft: 2,
    armorRows: 24,
    armorCols: 42,
    protectionTop: 111,
    protectionLeft: 54,
    protectionRows: 34,
    protectionCols: 10,
    armorSetCol: 25,
    armorDefenseCol: 13,
    armorProtectionCol: 16,
    armorStealthPenaltyCol: 27,
    armorConcealmentPenaltyCol: 29,
    armorEncumbranceCol: 31,
    armorSkillTop: 24,
    armorSkillLeft: 24
  }
  
  var curSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var characterSheet = curSpreadsheet.getSheetByName('Character Sheet');
  if (characterSheet === null) {
    SpreadsheetApp.getUi().alert("Unable to find Character Sheet\nMake sure the name of the first sheet is exactly 'Character Sheet'");
    return;
  }
  var armorRange = characterSheet.getRange(cellLocation.armorTop, cellLocation.armorLeft, cellLocation.armorRows, cellLocation.armorCols);
  
  var protectionTable = characterSheet.getRange(cellLocation.protectionTop, cellLocation.protectionLeft, cellLocation.protectionRows, cellLocation.protectionCols);
  
  var armorTable = [];
  var armorDoubleLocations = [1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 26];
  for (var i = 1; i <= 30; i++) {
    if (armorDoubleLocations.indexOf(i) != -1) {
      armorTable[i] = [ [ 0, 0 ], [ 0, 0 ], [ 0, 0 ], [ 0, 0 ] ];
      continue;
    }
    armorTable[i] = [ 0, 0, 0, 0 ];
  }
  
  var stealthPenalty = [ 0, 0, 0, 0 ];
  var concealmentPenalty = [ 0, 0, 0, 0 ];
  var encumbrance = [ 0, 0, 0, 0 ];
  
  // Clear Proteciton Table
  protectionTable.setValue('');
  
  for (var armorRow = 1; armorRow <= cellLocation.armorRows; armorRow++) {

    // Check what sets the Armor will go with    
    var armorString = armorRange.getCell(armorRow, cellLocation.armorSetCol).getValue();
    if (armorString == '') { continue; }
    var armorSets = parseRanges(armorString);
    
    // Get the armor value
    var armorValue = Number(armorRange.getCell(armorRow, cellLocation.armorDefenseCol).getValue());
    
    // Get the protection areas
    var protection = armorRange.getCell(armorRow, cellLocation.armorProtectionCol).getValue();
    if (protection == "") { continue; }
    var protectNumbers = parseRanges(protection);
    
    // Loop over Armor Sets
    for (var armorCol = 0; armorCol < armorSets.length; armorCol ++) {
      
      var armorSet = Number(armorSets[armorCol]) - 1;
    
      // Loop over Protection Numbers
      for (var protectNumber in protectNumbers) {
        var protectNum = protectNumbers[protectNumber];
        console.log('protectNum', protectNum);
        
        // Is range
        switch (protectNum) {
          case 'nf':
            armorTable[1][armorSet][0] -= armorValue;
            continue;
            
          case 'pf':
            armorTable[1][armorSet][0] -= Math.round(armorValue / 2);
            continue;
            
          case 'ng':
            armorTable[26][armorSet][0] -= armorValue;
            continue;
            
          case 'fo':
            for (var i = 5; i <= 15; i++) {
              armorTable[i][armorSet][0] -= armorValue;
            }
            continue;
        }
        
        if (armorDoubleLocations.indexOf(protectNum) !== -1) {
          console.log('armorDouble' + protectNum, armorTable[protectNum][armorSet]);
          armorTable[protectNum][armorSet][0] += armorValue;
          armorTable[protectNum][armorSet][1] += armorValue;
        } else {
          armorTable[protectNum][armorSet] += armorValue;
        }
        
      }
      
      stealthPenalty[armorSet] += Number(armorRange.getCell(armorRow, cellLocation.armorStealthPenaltyCol).getValue());
      concealmentPenalty[armorSet] += Number(armorRange.getCell(armorRow, cellLocation.armorConcealmentPenaltyCol).getValue());
      encumbrance[armorSet] += Number(armorRange.getCell(armorRow, cellLocation.armorEncumbranceCol).getValue());
    }
    
  }
  
  /*
   * Fill in the Armor Table
   */
  // Loop over armor sets
  for (var s = 0; s <= 3; s++) {
    
    // Loop over protection areas
    for (var i = 1; i <= 30; i++) {
      
      var protectValue = '';
      
      // If we're a double counted area, handle separately
      if (armorDoubleLocations.indexOf(i) != -1) {
        console.log('armor Double Location' + i, armorTable[i][s]);
        // If the values are the same, then just show a single number
        if (armorTable[i][s][0] == armorTable[i][s][1]) {
          
          // Only show the number if it's above 0
          if (armorTable[i][s][0] > 0) {
            protectValue = armorTable[i][s][0];
          }
          
        } else {
          protectValue = armorTable[i][s][0].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") +
            '/' +
            armorTable[i][s][1].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        }
        
      // else just show the value
      } else {
        
        // Only show the number if it's above 0
        if (armorTable[i][s] > 0) {
          protectValue = armorTable[i][s];
        }
      }
      
      // Fill in new valuein the armor column
      protectionTable.getCell(i, ((s + 1)*3)-2).setValue(protectValue);
      
      // Automatically adjust the font size if it's too long
      if (protectValue.length > 7) {
        protectionTable.getCell(i, ((s + 1)*3)-2).setFontSize(8);
      } else {
        protectionTable.getCell(i, ((s + 1)*3)-2).setFontSize(10);
      }
    }
    
    // Check encumbrance penalty
    if (encumbrance[s] > 0) {
      var encString = encumbrance[s].toFixed(1) + '';
      var armorSkill = Number(characterSheet.getRange(cellLocation.armorSkillTop, cellLocation.armorSkillLeft).getValue());
      if (armorSkill < encumbrance[s]) {
        encString += ' (' + (encumbrance[s].toFixed(1) - armorSkill) + ')';
      } else {
        encString += ' (0)';
      }
      protectionTable.getCell(32, ((s + 1)*3)-2).setValue(encString);
    }
    
    // Check the stealth penalty
    if (stealthPenalty[s] > 0) { protectionTable.getCell(33, ((s + 1)*3)-2).setValue(stealthPenalty[s]); }
    
    // Check the concealment Penalty
    if (concealmentPenalty[s] > 0) { protectionTable.getCell(34, ((s + 1)*3)-2).setValue(concealmentPenalty[s]); }
  }
}

/**
 * Check to see if a variable is numeric
 */
function isNumeric(n)
{
  return !isNaN(parseFloat(n)) && isFinite(n);
}

/**
 * Parse a range of inputs and return an array of locations
 */
function parseRanges(inputRanges)
{
  var returnArr = [];
  var arr = inputRanges.split(',');
  for (var i1 in arr) {
    if (arr[i1].indexOf('-') != -1) {
      var range = arr[i1].split('-');
      if (!isNumeric(range[0]) || !isNumeric(range[1])) {
        continue;
      }
      for (i2 = parseInt(range[0]); i2 <= parseInt(range[1]); i2++) {
        returnArr.push(i2);
      }
    } else {
      var val = arr[i1].trim();
      if (!isNumeric(val)) {
        if (allowableShorthand.indexOf(val) != -1) {
          returnArr.push(val);
        }
      } else {
        returnArr.push(parseInt(val));
      }
    }
  }
  
  return returnArr;
}
