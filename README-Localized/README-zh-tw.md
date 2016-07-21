# Excel 2016 的大學預算追蹤器工作窗格增益集範例

_適用版本：Excel 2016_

這個工作窗格增益集示範如何使用 Excel 2016 中的 JavaScript API 來建立大學預算追蹤器。共有兩種型態︰程式碼編輯器和 Visual Studio。

![大學預算追蹤器範例](../images/CollegeBudgetTracker_tracker.PNG)

## 進行測試
### 程式碼編輯器版本

部署及測試增益集的最簡單方式，是將資訊清單複製到網路共用。

1.  在網路共用上建立資料夾 (例如，\\\MyShare\CollegeBudgetTracker)。  
2.  將資訊清單 (CollegeBudgetTrackerManifest.xml) 複製到網路共用 (例如 \\\MyShare\MyManifests)。
3.  在 Excel 中，將包含資訊清單的共用位置新增為受信任的應用程式目錄。

    a.啟動 Excel，並開啟空白的試算表。  
    
    b.選擇 **[檔案]** 索引標籤，然後選擇 **[選項]**。
    
    c.選擇 **[信任中心]**，然後選擇 **[信任中心設定]** 按鈕。
    
    d.選擇 **[受信任的增益集目錄]**。
    
    e.在 **[目錄 URL]** 方塊中，輸入您在步驟 3 建立的網路共用路徑，然後選擇 **[新增目錄]**。
    
   f. 選取 [在功能表中顯示]<e /> 核取方塊，然後選擇 [確定]<e />。接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。 
        
4.  測試並執行增益集。 

    a.在 Excel 2016 的 **[插入]** 索引標籤上，選擇 **[我的增益集]**。
    
    b.在 **[Office 增益集]** 對話方塊中，選擇 **[共用資料夾]**。
    
    c.在 [首頁] 索引標籤上，按一下 **[大學預算追蹤器]** 命令。 增益集會在工作窗格中開啟，並在使用中工作表中建立大學預算追蹤器，如此圖表所示。 
      
   ![大學預算追蹤器範例](../images/CollegeBudgetTracker_tracker.PNG) 

    d.使用 **[新增費用]** 和 **[新增收入]** 索引標籤來新增一些支出和收入，並查看資料和圖表如何動態變更。
    
      ![大學預算追蹤器範例](../images/CollegeBudgetTracker_taskpane1.PNG) 

若要在您自己的增益集中使用資訊清單，編輯資訊清單檔案的 <SourceLocation> 元素，讓它指向 Home.html 檔案的共用位置。
    
### Visual Studio 版本
1.  將專案複製到本機資料夾，並在 Visual Studio 中開啟 Excel-Add-in-JS-CollegeBudgetTracker.sln。
2.  按 F5 建置及部署範例增益集。Excel 會啟動，且增益集會在工作表右側空白部分的工作窗格中開啟，如下圖所示。 
        
  ![大學預算追蹤器範例](../images/CollegeBudgetTracker_tracker.PNG) 

3.  使用 [新增費用]<e /> 和 [新增收入]<e /> 索引標籤來新增一些支出和收入，並查看資料和圖表如何動態變更。

  ![大學預算追蹤器範例](../images/CollegeBudgetTracker_taskpane1.PNG) 


### 深入了解

Excel JavaScript API 還有其他許多功能，可供您用於開發增益集。以下列出其中幾個可用的資源。 

1.  [Excel 增益集程式設計概觀](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Excel 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel 增益集程式碼範例](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Excel 增益集 JavaScript API 參考](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [建立第一個 Excel 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
