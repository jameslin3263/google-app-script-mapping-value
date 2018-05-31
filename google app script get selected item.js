function connect(){
    var collection = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("collection")
    var lc = collection.getLastColumn()
    var lr = collection.getLastRow()
    collection.getRange(2, 1, lr, lc).clearContent()
    
    getSheet1Value()
    }
  function getSheet1Value() {
    var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet1")
    var sheet1NumRows = sheet1.getLastRow() - 1
    
    var getProductGroupId = getColumnIndex("產品群組 ID","sheet1")
    
    sheet1.sort(getProductGroupId, true)
    
    var getValue = []
    var rule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("rule")
    var productId = rule.getRange(1, 1).getValues()
    getValue.push(productId)
    var reduceValue = getValue.reduce(function(pre,next){
      return pre+next
    })
    var reduceValue2 = reduceValue.reduce(function(pre,next){
      return pre+next
    }) 
    var value = reduceValue2[0].split(",").map(Number);
  
    var getValidProductGroupId = sheet1.getRange(2, getProductGroupId, sheet1NumRows, 1).getValues()
    
    var getVendor = getColumnIndex("客戶名稱","sheet1")
    var Vendor = sheet1.getRange(2, getVendor, sheet1NumRows, 1).getValues()
    
    var getProductName = getColumnIndex("活動名稱","sheet1")
    var productName = sheet1.getRange(2, getProductName, sheet1NumRows, 1).getValues()
    
    var getOrderNumber = getColumnIndex("訂單編號","sheet1")
    var orderNumber = sheet1.getRange(2, getOrderNumber, sheet1NumRows, 1).getValues()
    
    var getSetTime = getColumnIndex("建立時間","sheet1")
    var Time = sheet1.getRange(2, getSetTime, sheet1NumRows, 1).getValues()
    
    
    var getDepartureDate = getColumnIndex("出發日期","sheet1")
    var departureDate = sheet1.getRange(2, getDepartureDate, sheet1NumRows, 1).getValues()
    
    var getOrderDetails = getColumnIndex("訂單明細","sheet1")
    var orderDetails = sheet1.getRange(2, getOrderDetails, sheet1NumRows, 1).getValues()
    
    var getOriginalOrderPrice = getColumnIndex("訂單原價","sheet1")
    var originalOrderPrice = sheet1.getRange(2, getOriginalOrderPrice, sheet1NumRows, 1).getValues()
   
    var getTransDiscountAmount = getColumnIndex("轉單折抵金額","sheet1")
    var transDiscountAmount = sheet1.getRange(2, getTransDiscountAmount, sheet1NumRows, 1).getValues()
    
    var getVendorDiscountAmount = getColumnIndex("廠商主動折抵金額","sheet1")
    var VendorDiscountAmount = sheet1.getRange(2, getVendorDiscountAmount, sheet1NumRows, 1).getValues()
    
    var getNicedayDiscountAmount = getColumnIndex("Niceday主動折抵金額","sheet1")
    var nicedayDiscountAmount = sheet1.getRange(2, getNicedayDiscountAmount, sheet1NumRows, 1).getValues()
    
    var getTotalValue = getColumnIndex("實付金額","sheet1")
    var totalValue = sheet1.getRange(2, getTotalValue, sheet1NumRows, 1).getValues()				
    
    var arr = []
    for (var i=0;i<totalValue.length;i++){
      var object = {
        "key":i,
        "Vendor":Vendor[i],
        "Time":Time[i],
        "getValidProductGroupId":getValidProductGroupId[i],
        "productName":productName[i],
        "orderNumber":orderNumber[i],
        "departureDate":departureDate[i],
        "orderDetails":orderDetails[i],
        "originalOrderPrice":originalOrderPrice[i],
        "transDiscountAmount":transDiscountAmount[i],
        "VendorDiscountAmount":VendorDiscountAmount[i],
        "nicedayDiscountAmount":nicedayDiscountAmount[i],
        "totalValue":totalValue[i]
      }
      arr.push(object)
    }
    
    var collection =            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("collection")
    var setVendorName =         getColumnIndex("Vendor","collection")
    var setProductName =        getColumnIndex("產品名稱","collection")
    var setPGID =               getColumnIndex("PGID","collection")
    var setOrderNumber =        getColumnIndex("訂單編號","collection")
    var setTime =               getColumnIndex("建立時間","collection")
    var setdepartureDate =      getColumnIndex("出發日期","collection")
    var setOrderDetails =       getColumnIndex("訂單明細","collection")
    var setOriginalOrderPrice = getColumnIndex("訂單原價","collection")
    var setDisValue =           getColumnIndex("轉單折抵金額","collection")
    var setVenDisValue =        getColumnIndex("廠商主動折抵金額","collection")
    var setNDDisValue =         getColumnIndex("Niceday主動折抵金額","collection")
    var setValue =              getColumnIndex("實付金額","collection")
    
  
    var array = []
    for(var i=0;i<totalValue.length;i++){
       for(var j=0;j<value.length;j++){
         if(arr[i].getValidProductGroupId == value[j]) {
           array.push(arr[i])
         }
       }
    }
   
    for(var i=0;i<array.length;i++){
      collection.getRange(i+2,setVendorName,1,1).setValue(array[i].Vendor)
      collection.getRange(i+2,setPGID,1,1).setValue(array[i].getValidProductGroupId)
      collection.getRange(i+2,setProductName,1,1).setValue(array[i].productName)
      collection.getRange(i+2,setOrderNumber,1,1).setValue(array[i].orderNumber)
      collection.getRange(i+2,setTime,1,1).setValue(array[i].Time)
      collection.getRange(i+2,setdepartureDate,1,1).setValue(array[i].departureDate)
      collection.getRange(i+2,setOrderDetails,1,1).setValue(array[i].orderDetails)
      collection.getRange(i+2,setOriginalOrderPrice,1,1).setValue(array[i].originalOrderPrice)
      collection.getRange(i+2,setDisValue,1,1).setValue(array[i].transDiscountAmount)
      collection.getRange(i+2,setVenDisValue,1,1).setValue(array[i].VendorDiscountAmount)
      collection.getRange(i+2,setNDDisValue,1,1).setValue(array[i].nicedayDiscountAmount)
      collection.getRange(i+2,setValue,1,1).setValue(array[i].totalValue)
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