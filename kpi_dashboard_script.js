function createKPIDashboard() {
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     const rawSheet = ss.getSheetByName("Raw_Data");
     let dashboard = ss.getSheetByName("KPI_Dashboard");
   
     if (dashboard) {
       ss.deleteSheet(dashboard);
     }
   
     dashboard = ss.insertSheet("KPI_Dashboard");
   
     const data = rawSheet.getDataRange().getValues();
     data.shift();
   
     let totalRevenue = 0;
     let totalQuantity = 0;
     let monthlySales = {};
   
     data.forEach(row => {
       const dateStr = row[1];
       const quantity = row[3];
       const price = row[4];
       const date = new Date(dateStr);
       const month = date.toLocaleString("default", { month: "short" });
   
       totalRevenue += quantity * price;
       totalQuantity += quantity;
   
       if (!monthlySales[month]) monthlySales[month] = 0;
       monthlySales[month] += quantity * price;
     });
   
     const averageOrderValue = totalRevenue / totalQuantity;
   
     dashboard.getRange("A1").setValue("📊 KPI Dashboard");
     dashboard.getRange("A3").setValue("Total Revenue:");
     dashboard.getRange("B3").setValue(totalRevenue);
     dashboard.getRange("A4").setValue("Total Quantity Sold:");
     dashboard.getRange("B4").setValue(totalQuantity);
     dashboard.getRange("A5").setValue("Average Order Value:");
     dashboard.getRange("B5").setValue(averageOrderValue.toFixed(2));
   
     dashboard.getRange("A7").setValue("📅 Monthly Sales:");
     let row = 8;
     for (let month in monthlySales) {
       dashboard.getRange("A" + row).setValue(month);
       dashboard.getRange("B" + row).setValue(monthlySales[month]);
       row++;
     }
   
     SpreadsheetApp.flush();
   }
   