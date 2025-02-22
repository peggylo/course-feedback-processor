// 常數定義 - 使用不同的命名以避免衝突
const FEEDBACK_SOURCE_SPREADSHEET_ID = '11N3O17AWSOi61ASVbNUzk5bP_3jqwZvYNEFqreMKT_0';
const HACKFOLDR_SPREADSHEET_ID = '14XcP4wHaxj6iL_emG9029mGb1vC2KRL1qSXayVuk3s4';
const TEACHER_NAME = '請填入講師名稱';
const FEEDBACK_SHEET_NAME = '學員回饋';

// 主要執行函數
function createTeacherFeedback() {
  try {
    // 1. 取得講師的試算表 URL
    const teacherSpreadsheetUrl = getTeacherSpreadsheetUrl();
    if (!teacherSpreadsheetUrl) {
      throw new Error(`找不到${TEACHER_NAME}的試算表連結`);
    }
    
    // 2. 從 URL 取得 spreadsheet ID
    const teacherSpreadsheetId = extractSpreadsheetId(teacherSpreadsheetUrl);
    
    // 3. 取得並處理問卷資料
    const feedbackData = processFeedbackData();
    
    // 4. 在講師的試算表中建立新的 sheet
    createFeedbackSheet(teacherSpreadsheetId, feedbackData);
    
    Logger.log('處理完成！');
    // 新增：打印講師試算表網址
    Logger.log(`${TEACHER_NAME}的試算表網址：${teacherSpreadsheetUrl}`);
  } catch (error) {
    Logger.log('錯誤：' + error.message);
    throw error;
  }
}

// 取得講師試算表 URL
function getTeacherSpreadsheetUrl() {
  const hackfoldrSheet = SpreadsheetApp.openById(HACKFOLDR_SPREADSHEET_ID).getSheetByName('list');
  const data = hackfoldrSheet.getDataRange().getValues();
  
  // 找到包含講師權限描述的列
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === `僅${TEACHER_NAME}老師有權限`) {
      return data[i][0]; // 返回 #url 欄位的值
    }
  }
  return null;
}

// 從 URL 提取 spreadsheet ID
function extractSpreadsheetId(url) {
  const matches = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (matches && matches[1]) {
    return matches[1];
  }
  throw new Error('無法從 URL 提取試算表 ID');
}

// 處理問卷資料
function processFeedbackData() {
  const sourceSheet = SpreadsheetApp.openById(FEEDBACK_SOURCE_SPREADSHEET_ID).getSheetByName('課後問卷');
  const workSheet = SpreadsheetApp.openById(FEEDBACK_SOURCE_SPREADSHEET_ID).getSheetByName('工作表');
  
  const sourceData = sourceSheet.getDataRange().getValues();
  const workData = workSheet.getDataRange().getValues();
  
  // 建立學員資料對照表
  const studentInfo = createStudentInfoMap(workData);
  
  // 找到講師評分區塊的欄位位置
  const { startCol, headers: ratingHeaders } = findTeacherColumns(sourceData[0], TEACHER_NAME);
  
  // 處理每個學員的資料
  const feedbackData = [];
  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const studentName = row[0];
    
    // 檢查該列是否有劉正吉的評分資料
    if (row[startCol] === TEACHER_NAME) {
      const rating = findRating(row, startCol);
      if (rating) {
        feedbackData.push({
          name: studentName,
          school: studentInfo[studentName]?.school || '--',
          position: studentInfo[studentName]?.position || '--',
          rating: rating,
          feedback: row[getFeedbackColumn(sourceData[0])] || '--'
        });
      }
    }
  }
  
  // 加入統計資訊的日誌
  Logger.log(`總共找到 ${feedbackData.length} 位參與 ${TEACHER_NAME} 課程並填寫評分的學員`);
  
  return feedbackData;
}

// 建立學員資料對照表
function createStudentInfoMap(workData) {
  const map = {};
  const nameCol = workData[0].indexOf('教師姓名');
  const schoolCol = workData[0].indexOf('學校名稱');
  const positionCol = workData[0].indexOf('身分');
  
  for (let i = 1; i < workData.length; i++) {
    const row = workData[i];
    if (row[nameCol]) {  // 確保有教師姓名才加入
      map[row[nameCol]] = {
        school: row[schoolCol] || '--',
        position: row[positionCol] || '--'
      };
    }
  }
  
  return map;
}

// 找到講師評分區塊的欄位位置
function findTeacherColumns(headers, teacherName) {
  for (let i = 0; i < headers.length; i++) {
    // 找到包含「您心中的【收穫程度】是」的欄位
    if (headers[i].includes('您心中的【收穫程度】是') && 
        headers[i].includes('議程')) {
      // 檢查下一列的資料是否為講師名稱
      const sourceSheet = SpreadsheetApp.openById(FEEDBACK_SOURCE_SPREADSHEET_ID).getSheetByName('課後問卷');
      const data = sourceSheet.getDataRange().getValues();
      
      // 檢查第二列開始的資料
      for (let row = 1; row < data.length; row++) {
        if (data[row][i] === teacherName) {
          return {
            startCol: i,
            headers: headers.slice(i, i + 6)
          };
        }
      }
    }
  }
  throw new Error(`找不到${teacherName}的評分欄位`);
}

// 找出評分
function findRating(row, startCol) {
  const ratings = row.slice(startCol + 1, startCol + 6);
  for (let i = 0; i < ratings.length; i++) {
    if (ratings[i]) {
      return getRatingText(i);
    }
  }
  return null;
}

// 取得評分文字
function getRatingText(index) {
  const ratingTexts = [
    '沒有學到新東西',
    '普通',
    '有學到新東西',
    '學到非常多新東西',
    '收穫度Max，所有課程中這是我心中NO1'
  ];
  return ratingTexts[index];
}

// 找到心得回饋欄位
function getFeedbackColumn(headers) {
  for (let i = 0; i < headers.length; i++) {
    // 修改搜尋條件以符合實際問卷欄位名稱
    if (headers[i].includes(`給【課程名稱-${TEACHER_NAME}】課程內容的建議或心得`)) {
      return i;
    }
  }
  throw new Error('找不到心得回饋欄位');
}

// 建立回饋 sheet
function createFeedbackSheet(spreadsheetId, feedbackData) {
  const teacherSheet = SpreadsheetApp.openById(spreadsheetId);
  
  // 檢查是否已存在同名 sheet
  let sheet = teacherSheet.getSheetByName(FEEDBACK_SHEET_NAME);
  if (sheet) {
    // 如果已存在，先刪除
    teacherSheet.deleteSheet(sheet);
  }
  
  // 建立新的 sheet 在最前面（索引 0）
  sheet = teacherSheet.insertSheet(FEEDBACK_SHEET_NAME, 0);
  
  // 凍結第一列
  sheet.setFrozenRows(1);
  
  // 設定頁籤顏色為紅色
  sheet.setTabColor('#FF0000');
  
  // 設定欄寬
  sheet.setColumnWidth(1, 80);    // 學員姓名
  sheet.setColumnWidth(2, 150);   // 學校名稱 (縮短)
  sheet.setColumnWidth(3, 80);    // 身分 (縮短)
  sheet.setColumnWidth(4, 270);   // 收穫程度 (加長)
  sheet.setColumnWidth(5, 400);   // 建議或心得
  
  // 設定標題
  const headers = ['學員姓名', '學校名稱', '身分', 
                  `【${TEACHER_NAME}老師】課程給你的收穫程度\n(選項:收穫度Max,學到很多,有學到,普通,沒學到新東西)`, 
                  `給【${TEACHER_NAME}老師】分享內容的建議或心得`];
  
  // 寫入標題
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 設定標題格式
  headerRange.setFontColor("#FFFFFF")  // 文字顏色：白色
            .setHorizontalAlignment("center")  // 文字置中
            .setVerticalAlignment("middle")    // 垂直置中
            .setFontWeight("bold")  // 文字粗體
            .setWrap(true);  // 允許文字換行
            
  // 設定第四欄標題的字體格式（使用 Rich Text）
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(headers[3])
    .setTextStyle(0, headers[3].indexOf('\n'), SpreadsheetApp.newTextStyle().setFontSize(11).build())
    .setTextStyle(headers[3].indexOf('\n') + 1, headers[3].length, SpreadsheetApp.newTextStyle().setFontSize(8).build())
    .build();
  
  sheet.getRange(1, 4).setRichTextValue(richText);
  
  // 設定標題背景顏色
  sheet.getRange(1, 1, 1, 3).setBackground("#6C6C6C");  // 前三欄淺灰色
  sheet.getRange(1, 4, 1, 2).setBackground("#272727");  // 後兩欄深灰色
  
  // 寫入資料
  if (feedbackData.length > 0) {
    // 根據收穫程度排序
    const ratingOrder = [
      '收穫度Max，所有課程中這是我心中NO1',
      '學到非常多新東西',
      '有學到新東西',
      '普通',
      '沒有學到新東西'
    ];
    
    const sortedData = feedbackData.sort((a, b) => 
      ratingOrder.indexOf(a.rating) - ratingOrder.indexOf(b.rating)
    );
    
    const values = sortedData.map(d => [d.name, d.school, d.position, d.rating, d.feedback]);
    const dataRange = sheet.getRange(2, 1, values.length, headers.length);
    dataRange.setValues(values);
    
    // 設定前四欄的所有資料置中對齊
    sheet.getRange(1, 1, values.length + 1, 4)
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
    
    // 設定最後一欄自動換行
    sheet.getRange(1, 5, values.length + 1, 1)
         .setWrap(true);
  }
  
  // 刪除所有超過5欄的欄位
  const totalColumns = sheet.getMaxColumns();
  if (totalColumns > 5) {
    sheet.deleteColumns(6, totalColumns - 5);
  }
}

// 新增選單項目
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('問卷回饋工具')
    .addItem('處理講師回饋', 'createTeacherFeedback')
    .addToUi();
}

function processTeacherFeedback(feedbackData, studentInfoMap) {
  const headers = feedbackData[0];
  
  const day1TeacherCol = headers.indexOf('您2/8(週六)上哪位講師的課：');
  const day2TeacherCol = headers.indexOf('您2/9(週日)上哪位講師的課：');
  
  const teacherCols = findTeacherColumns(headers, TEACHER_NAME);
  
  // 打印標題資訊
  Logger.log('=== 標題資訊 ===');
  Logger.log('週六上課欄位: ' + day1TeacherCol);
  Logger.log('週日上課欄位: ' + day2TeacherCol);
  Logger.log('評分欄位起始位置: ' + teacherCols.startCol);
  Logger.log('評分欄位標題: ' + JSON.stringify(teacherCols.headers));

  const feedbackCol = headers.indexOf(`給【語音聲控 Lesson 1-${TEACHER_NAME}】課程內容的建議或心得？`);
  Logger.log('回饋欄位: ' + feedbackCol);

  const processedData = [];
  
  // 打印每位學員的資料
  Logger.log('\n=== 學員資料 ===');
  for (let i = 1; i < feedbackData.length; i++) {
    const row = feedbackData[i];
    const name = row[0];
    
    Logger.log('\n學員: ' + name);
    Logger.log('週六上課: ' + row[day1TeacherCol]);
    Logger.log('週日上課: ' + row[day2TeacherCol]);
    
    const hasAttendedClass = row[day1TeacherCol] === TEACHER_NAME || 
                           row[day2TeacherCol] === TEACHER_NAME;
    
    Logger.log('是否上課: ' + hasAttendedClass);
    
    if (hasAttendedClass) {
      const rating = findRating(row, teacherCols.startCol);
      Logger.log('評分: ' + rating);
      Logger.log('回饋: ' + row[feedbackCol]);
      
      if (rating) {
        const studentInfo = studentInfoMap[name] || { school: '--', position: '--' };
        processedData.push({
          name: name,
          school: studentInfo.school,
          position: studentInfo.position,
          rating: rating,
          feedback: row[feedbackCol] || '--'
        });
      }
    }
  }
  
  // 打印最終處理結果
  Logger.log('\n=== 最終結果 ===');
  Logger.log('處理資料筆數: ' + processedData.length);
  Logger.log(JSON.stringify(processedData, null, 2));
  
  return processedData;
} 
