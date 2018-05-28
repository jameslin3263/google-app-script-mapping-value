function connect(){
    var collection = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("collection")
    var lc = collection.getLastColumn()
    var lr = collection.getLastRow()
    
    collection.getRange(2, 1, lr-1, 4).clearContent()
    collection.getRange(2, 6, lr-1, 1).clearContent()
    collection.getRange(2, 8, lr-1, lc).clearContent()
    
    getSheet1Value();
    getSheet2Value();
  }
  
  function getSheet1Value(){
    var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet1")
    var sheet1NumRows = sheet1.getLastRow() - 1
    var getSalesValue = getColumnIndex("總業績","sheet1")
    
    sheet1.sort(getSalesValue, false)
    
    var totalValue = sheet1.getRange(2, getSalesValue, sheet1NumRows, 1).getValues()
    var lastWeekValue = totalValue.filter(function(x) { return x > 0 })
    
    var getProductName = getColumnIndex("產品名稱","sheet1")
    var productName = sheet1.getRange(2, getProductName, lastWeekValue.length, 1).getValues()
    
   
    var getSalesPercentage = getColumnIndex("總銷售比","sheet1")
    var salePercentage = sheet1.getRange(2, getSalesPercentage, lastWeekValue.length, 1).getValues()
    
    var getPageView = getColumnIndex("PV","sheet1")
    var pageView = sheet1.getRange(2, getPageView, lastWeekValue.length, 1).getValues()
    
    var getUniquePageView = getColumnIndex("UPV","sheet1")
    var uniquePageView = sheet1.getRange(2, getUniquePageView, lastWeekValue.length, 1).getValues()
    
    var getProductGroupID = getColumnIndex("PGID","sheet1")
    var productGroupID = sheet1.getRange(2, getProductGroupID, lastWeekValue.length, 1).getValues()
    
    
    var collection =            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("collection")
    var numRows =               sheet1.getLastRow() - 1
    var setProductName =        getColumnIndex("產品名稱","collection")
    var setTotalValues =        getColumnIndex("總業績","collection")
    var setSalesPercentage =    getColumnIndex("總銷售比","collection")
    var setPageView =           getColumnIndex("PV","collection")
    var setUniquePageView =     getColumnIndex("UPV","collection")
    var setProductGroupID =     getColumnIndex("PGID","collection")
    
    collection.getRange(2, setTotalValues, lastWeekValue.length, 1).setValues(lastWeekValue)
    collection.getRange(2, setProductName, lastWeekValue.length,1).setValues(productName)
    collection.getRange(2, setSalesPercentage, lastWeekValue.length, 1).setValues(salePercentage)
    collection.getRange(2, setPageView, lastWeekValue.length, 1).setValues(pageView)
    collection.getRange(2, setUniquePageView, lastWeekValue.length, 1).setValues(uniquePageView)
    collection.getRange(2, setProductGroupID, lastWeekValue.length, 1).setValues(productGroupID)
  
  }
  
  function getSheet2Value() {
    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet2")
    var sheet2NumRows = sheet2.getLastRow() - 1
    
    var getSalesValue = getColumnIndex("總業績","sheet2")
    var totalValue = sheet2.getRange(2, getSalesValue, sheet2NumRows, 1).getValues()
  
    var getProductName = getColumnIndex("產品名稱","sheet2")
    var productName = sheet2.getRange(2, getProductName, sheet2NumRows, 1).getValues()
    
    
    var collection = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("collection")
    var collectionNumRows = collection.getLastRow() - 1
    
    var getCollectionProductName = getColumnIndex("產品名稱","collection")
    var collectionProductName = collection.getRange(2, getCollectionProductName, collectionNumRows, 1).getValues()
   
  
    function putObjInNewArray (productName,totalValue){
      var result = []
      for (var i=0; i<productName.length;i++){
        var object = { 'key':'','name': productName[i][0], 'value': totalValue[i][0]}
        result.push(object)
      }
      return result
    }
    
    var arr = []
    var collection =            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("collection")
    var setProductName =        getColumnIndex("上上週產品名稱","collection")
    var setTotalValues =        getColumnIndex("上上週總業績","collection")
  
    putObjInNewArray(productName,totalValue).forEach(function(x){
      collectionProductName.forEach(function(y){
        if (x.name == y[0]) {
          arr.push(x)  
        }
      })
    })
    
    var result = []
    for (var i=0; i<collectionProductName.length;i++){
        var object = { 'key': i,'name': collectionProductName[i][0] }
        result.push(object)
      }
    
    for(var i=0;i<result.length;i++){
       for(var j=0;j<arr.length;j++){
         if(result[i].name == arr[j].name) {
           arr[j].key = result[i].key
           collection.getRange(arr[j].key+2,setProductName,1,1).setValue(arr[j].name)
           collection.getRange(arr[j].key+2,setTotalValues,1,1).setValue(arr[j].value)
         }
       }
    }
  }
  
  //Helper Function
  
  function getColumnValues(label,sheetName) {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var colIndex = getColumnIndex(label,sheetName);
    var numRows = ss.getLastRow() - 1;
    var colValues = ss.getRange(2, colIndex, numRows, 1).getValues();
    
    return colValues;
  }
  
  
  function getColumnIndex(label,sheetName) {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var lc = ss.getLastColumn();
    var lookupRangeValues = ss.getRange(1, 1, 1, lc).getValues()[0];
    var index = lookupRangeValues.indexOf(label) + 1;
    
    return index;
  }