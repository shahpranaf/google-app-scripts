function calculateAvg() {  
  
  const name = 'Caregrowth Portfolio' ;
  const activeSheet = SpreadsheetApp.getActive().getSheetByName(name);
  const sheet = activeSheet.getDataRange().getValues();

  if(!activeSheet) {
    return;
  }  
  
  const stock_code_addr = 1;
  const bp_addr = 3;
  const avg_addr = 5;
  
  const obj = {};
  let mergeVal;
  for(let i=2; i< sheet.length; i++) {
    let base = sheet[i][0];
    const partOfMerge = activeSheet.getRange(i+1, 1).isPartOfMerge();
    if( partOfMerge ) {
      if(!mergeVal) {
        mergeVal = base;
      }
      else {
        base = mergeVal; 
      }
    } else {
     mergeVal = null; 
    }
        
    if(!obj[base]) {    
      obj[base] = [];
    }
    Logger.log(sheet[i][bp_addr], base, activeSheet.getRange(i+1, 1).isPartOfMerge());
    obj[base].push(sheet[i][bp_addr]);
  }
  Logger.log(obj);

  for(let i=2; i< sheet.length; i++) {
   
    if(!sheet[i][stock_code_addr]) {
      continue;
    }
    const base = sheet[i][0];
    const length = obj[base] ? obj[base].length : 0;
    const avg = length ? (obj[base].reduce((a,c) => a+c))/length : obj[base];
   
    activeSheet.getRange(i+1, avg_addr).setValue(avg);    
   
  }  
}
