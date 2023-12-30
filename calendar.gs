 @OnlyCurrentDoc 

function testMacro() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('calendar');
  sheet.getRange(5,5).setValue('fortest');
};

 メイン処理 
const main=()={
  const yukiCalendar = CalendarApp.getCalendarById(nun.odakazu@gmail.com)
  const firstDay = new Date(202411)
  const lastDay = new Date(20241231)
  const allSchedules = yukiCalendar.getEvents(firstDay,lastDay)

  const calenderSheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName(calendar)
  let currentRow = 4
  const dateRow = 2

  const dateRange= calenderSheet.getRange(`${dateRow}${dateRow}`)


  for (let i =0;i  allSchedules.length;i++){

     予定の必要情報を抽出 
    const scheduleTitle= allSchedules[i].getTitle()
    const startTime = convertDateToString(allSchedules[i].getStartTime())
    const endTime = convertDateToString(allSchedules[i].getEndTime())
    const scheduleColor = allSchedules[i].getColor()
    const scheduleId = allSchedules[i].getId()

    その予定の日付がどの列に対応するか算出 
    const targetColumn = findTargetDateColumn(calenderSheet,startTime)

     id タイトル記入 
    calenderSheet.getRange(currentRow,1).setValue(scheduleId)
    calenderSheet.getRange(currentRow,2).setValue(scheduleTitle)

     xを記入 
    inputScheduleX(calenderSheet,startTime,endTime,currentRow)

    currentRow+=1

  }

}

 カレンダーシートの初期化 
const initCalendarSheet=()={

}

 シート上に予定の対象日付をxで表記 
const inputScheduleX=(sheet,startTime,endTime,currentRow)={
  let startColumn = findTargetDateColumn(sheet,startTime)
  let endColumn = findTargetDateColumn(sheet,endTime)

   一日のみの予定はとばす 
  if (startColumn==endColumn){
    sheet.getRange(currentRow,startColumn).setValue(x)
    return
  }

   x記入 
  for(let col = startColumn;col=endColumn;col++){
    sheet.getRange(currentRow,col).setValue(x)
  }

}

 calendarシート上でyyyymmdd形式の文字列がどの行に該当するか行番号を算出する 
const findTargetDateColumn=(sheet,DateString)={
  let targetColumn
  for(let i =1;i100;i++){
    const targetCell= sheet.getRange(2,i)
    const cellValue = targetCell.getValue()

     日付でないセルならスキップ 
    if(cellValue instanceof Date ==false){
      continue
    }

    const cellDateString=convertDateToString(cellValue)

    if(cellDateString==DateString){
      targetColumn=i
      break
    }
  }

  return targetColumn
}

 日付型の変数を二つ渡すとそれが一致する日付かどうかを判断する。戻り値はboolean 
const judgeMatchDate=(dateA,dateB)={

   yyyymmdd の形式に日付型を変換 
  const dateAString = convertDateToString(dateA)
  const dateBString = convertDateToString(dateB)

   dateAとdateBの比較 
  if (dateAString==dateBString){
    return true
  }else{
    return false
  }

}

 yyyymmdd の形式に日付型を変換 
const convertDateToString=(targetDate)={
  const yyyy= targetDate.getFullYear()
  const mm=  (0 + (targetDate.getMonth()+1)).slice(-2)
  const dd=(0 + targetDate.getDate()).slice(-2)
  const dateString= `${yyyy}${mm}${dd}`
  return dateString
}


















