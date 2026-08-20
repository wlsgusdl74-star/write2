const SS_ID = ""; // 비워두면 현재 연결된 스프레드시트를 사용합니다.



function getActiveSheet_() {

  if (SS_ID) {

    return SpreadsheetApp.openById(SS_ID);

  }

  return SpreadsheetApp.getActiveSpreadsheet();

}



function getOrCreateSheet_(ss, sheetName) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {

    sheet = ss.insertSheet(sheetName);

    if (sheetName === 'Teachers') {

      sheet.appendRow(['className', 'password']);

    } else if (sheetName === 'Articles') {

      sheet.appendRow(['id', 'timestamp', 'targetTeacher', 'name', 'mood', 'shortMsg', 'activityTitle', 'w6', 'draftText']);

    }

  }

  return sheet;

}



function doGet(e) {

  const action = e.parameter.action;

  const ss = getActiveSheet_();

 

  try {

    if (action === 'getClasses') {

      const sheet = getOrCreateSheet_(ss, 'Teachers');

      const rows = sheet.getDataRange().getValues();

      const classes = [];

      for (let i = 1; i < rows.length; i++) {

        if (rows[i][0]) classes.push(rows[i][0]);

      }

      return createJsonResponse({ success: true, data: classes });

     

    } else if (action === 'getData') {

      const sheet = getOrCreateSheet_(ss, 'Articles');

      const rows = sheet.getDataRange().getValues();

      const articles = [];

      for (let i = 1; i < rows.length; i++) {

        if (!rows[i][0]) continue;

        let w6Obj = {};

        try { w6Obj = JSON.parse(rows[i][7] || '{}'); } catch(e) {}

       

        articles.push({

          id: rows[i][0],

          timestamp: rows[i][1],

          targetTeacher: rows[i][2],

          name: rows[i][3],

          mood: rows[i][4],

          shortMsg: rows[i][5],

          activityTitle: rows[i][6],

          w6: w6Obj,

          draftText: rows[i][8]

        });

      }

      return createJsonResponse({ success: true, data: articles });

     

    } else if (action === 'clearData') {

      const teacher = e.parameter.teacher;

      const sheet = getOrCreateSheet_(ss, 'Articles');

      const rows = sheet.getDataRange().getValues();

     

      for (let i = rows.length - 1; i >= 1; i--) {

        if (rows[i][2] === teacher) {

          sheet.deleteRow(i + 1);

        }

      }

      return createJsonResponse({ success: true, message: `[${teacher}] 반의 데이터가 초기화되었습니다.` });

    }

   

    return createJsonResponse({ success: false, message: '잘못된 요청입니다.' });

  } catch (err) {

    return createJsonResponse({ success: false, message: err.toString() });

  }

}



function doPost(e) {

  try {

    const data = JSON.parse(e.postData.contents);

    const action = data.action;

    const ss = getActiveSheet_();

   

    if (action === 'registerTeacher') {

      const className = data.className.trim();

      const password = data.password.trim();

     

      const sheet = getOrCreateSheet_(ss, 'Teachers');

      const rows = sheet.getDataRange().getValues();

     

      for (let i = 1; i < rows.length; i++) {

        if (rows[i][0] === className) {

          return createJsonResponse({ success: false, message: '이미 존재하는 반 이름입니다.' });

        }

      }

     

      sheet.appendRow([className, password]);

      return createJsonResponse({ success: true, message: '반이 성공적으로 생성되었습니다.' });

     

    } else if (action === 'loginTeacher') {

      const className = data.className;

      const password = data.password;

     

      const sheet = getOrCreateSheet_(ss, 'Teachers');

      const rows = sheet.getDataRange().getValues();

     

      for (let i = 1; i < rows.length; i++) {

        if (rows[i][0] === className && String(rows[i][1]) === String(password)) {

          return createJsonResponse({ success: true });

        }

      }

      return createJsonResponse({ success: false, message: '비밀번호가 일치하지 않습니다.' });

     

    } else if (action === 'saveArticle') {

      const sheet = getOrCreateSheet_(ss, 'Articles');

      const id = 'art_' + new Date().getTime() + Math.random().toString(36).substring(2, 7);

      const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

     

      sheet.appendRow([

        id,

        timestamp,

        data.targetTeacher,

        data.name,

        data.mood,

        data.shortMsg,

        data.activityTitle,

        JSON.stringify(data.w6 || {}),

        data.draftText

      ]);

     

      return createJsonResponse({ success: true, message: '🎉 기사가 성공적으로 제출되었습니다!' });

    }

   

    return createJsonResponse({ success: false, message: '지원하지 않는 액션입니다.' });

  } catch (err) {

    return createJsonResponse({ success: false, message: err.toString() });

  }

}



function createJsonResponse(data) {

  return ContentService.createTextOutput(JSON.stringify(data))

    .setMimeType(ContentService.MimeType.JSON);

}

