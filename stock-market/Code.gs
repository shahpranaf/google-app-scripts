function getQuotesValue() {  
  
  const name = 'Caregrowth Portfolio' ;
  const activeSheet = SpreadsheetApp.getActive().getSheetByName(name);
  const sheet = activeSheet.getDataRange().getValues();

  if(!activeSheet) {
    return;
  }  
  const sc_id_addr = 8;
  const stock_code_addr = 1;
  const curr_price_addr = 6;
  const fifty2H_price_addr = 7;
  const fifty2L_price_addr = 8;
  
  for(let i=2; i< sheet.length; i++) {
    let sc_id;
    if(sheet[i][sc_id_addr]) {
      sc_id = sheet[i][sc_id_addr];
    } else {
      if(!sheet[i][stock_code_addr]) {
        continue;
      }
      Logger.log(sheet[i][stock_code_addr]);
      const url = `https://www.moneycontrol.com/mccode/common/autosuggestion_solr.php?classic=true&query=${sheet[i][0]}&type=1&format=json`;
      sc_id = JSON.parse(UrlFetchApp.fetch(url))[0].sc_id;
      
      activeSheet.getRange(i+1, sc_id_addr+1).setValue(sc_id);    
    }
    const response = UrlFetchApp.fetch("https://priceapi-aws.moneycontrol.com/pricefeed/nse/equitycash/"+sc_id);
    var res = JSON.parse(response.getContentText());
    
    if(res.code === "201" ) {        
        const response = UrlFetchApp.fetch("https://priceapi-aws.moneycontrol.com/pricefeed/bse/equitycash/"+sc_id);
        res = JSON.parse(response.getContentText());
    }
    
    if(res.code === "200" ) {
      const data  = res.data;
      const currPrice = data.pricecurrent;
      const fifty2H = data['52H'];
      const fifty2L = data['52L'];
      
      
      Logger.log(i+1,curr_price_addr, currPrice);
      activeSheet.getRange(i+1,curr_price_addr).setValue(currPrice);
      activeSheet.getRange(i+1,fifty2H_price_addr).setValue(fifty2H);
      activeSheet.getRange(i+1,fifty2L_price_addr).setValue(fifty2L);
      
    } else {
      
        Logger.log("Script not found"); 
 
    }
  }
  
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Menu')
      .addItem('Refresh Stock','getQuotesValue')
      .addItem("Calculate AVG", 'calculateAvg')
      .addToUi();
}
