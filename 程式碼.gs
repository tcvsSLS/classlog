let LogSS_URL = 'https://docs.google.com/spreadsheets/d/18Ainr-msGb2Lf9VFuiSopinguJP8gVpJbmT4-pvkjAU/edit'; //日誌資料
let UserSS_URL = 'https://docs.google.com/spreadsheets/d/1snxynZ_Ae2RyuSaxkeQ3DrevnuZNI4qmeVj1DA1BaaM/edit'; //現在人員資料
let KeyefudaoSS_URL = 'https://docs.google.com/spreadsheets/d/1VE-vrCoYsw1WELkX4LAYqiSs1qQERNPmx05PefLsV-g/edit'; //課業輔導開課資料

function doGet() {
  const html = HtmlService.createHtmlOutputFromFile('logform')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getUserName(){
  let currentUser=Session.getActiveUser().getEmail();  
//  測試用教師電子郵件
//  let currentUser="a2247@tcvs.ilc.edu.tw";
  let ss = SpreadsheetApp.openByUrl(UserSS_URL);
  const user = [];
  user[0]=currentUser;
  if(currentUser.indexOf("@tcvs.ilc.edu.tw") > 0){
    user[1]="教師";
    let sheet = ss.getSheetByName(user[1]);
    let range = sheet.getDataRange();
    let userdata = range.getValues();
    let lastRow = range.getLastRow();
    for(let i=0;i<lastRow;i++){
      if(userdata[i][1]==currentUser){user[0]=userdata[i][0];break;}
    }
  }else{
    if(currentUser.indexOf("@stu.tcvs.ilc.edu.tw") > 0){
      user[1]="學生";
      let sheet = ss.getSheetByName(user[1]);
      let range = sheet.getDataRange();
      let userdata = range.getValues();
      let lastRow = range.getLastRow();
      let schoolno = '';
      for(let i=0;i<lastRow;i++){
        if(userdata[i][2].length==5){schoolno='0'+userdata[i][2];}else{schoolno=userdata[i][2];}
        if((schoolno+"@stu.tcvs.ilc.edu.tw")==currentUser){user[0]=userdata[i][5];break;}
    }
    }else{
      user[1]="非本校人員";
    }
  }
  return user;  
}

function getSubjectType(){
  let sheet = SpreadsheetApp.openByUrl(LogSS_URL).getSheetByName("科目類型");
  let range = sheet.getDataRange();
  let data = range.getValues();
  return data;
}

function getTeacherData(){
  let sheet = SpreadsheetApp.openByUrl(LogSS_URL).getSheetByName("配課");
  let tdata = sheet.getDataRange().getValues();
  return tdata;
}

function getlog(stype,sname){
  let sheet = SpreadsheetApp.openByUrl(LogSS_URL).getSheetByName("日誌");
  let range = sheet.getDataRange();
  let lastRow = range.getLastRow();
  let data = range.getValues();
  const originallog=[];
  for(let j=0;j<lastRow;j++){
    if(data[j][0]==stype){
      if(data[j][1]==sname){
        if(data[j][7]==""){
          originallog.push([data[j][3],data[j][4],data[j][6],'',j+1]);
        }
      }
    }
  }
  return originallog;
}

function changeData(sdata){
  let currentDate=new Date();
  let linenum=sdata[7];
  let sheet = SpreadsheetApp.openByUrl(LogSS_URL).getSheetByName("日誌");
  if(sdata[6]=='刪除'){
    sheet.getRange("H"+linenum).setValue("刪除");
  }
  if(sdata[6]=='新增'){
    sheet.appendRow([sdata[0],sdata[1],sdata[2],sdata[3],sdata[4],currentDate,sdata[5],'']);
  }
  if(sdata[6]=='修改'){
    let copyRow = sheet.getRange("A"+linenum+":H"+linenum).getValues();
    copyRow[0][7]='刪除';
    sheet.appendRow(copyRow[0]);
    sheet.getRange("D"+linenum).setValue(sdata[3]);
    sheet.getRange("E"+linenum).setValue(sdata[4]);
    sheet.getRange("F"+linenum).setValue(currentDate);
    sheet.getRange("G"+linenum).setValue(sdata[5]);
  }
}