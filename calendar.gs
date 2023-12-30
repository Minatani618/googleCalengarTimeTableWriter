/** @OnlyCurrentDoc */

function testMacro() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('calendar');
  sheet.getRange(5,5).setValue('fortest');
};

/* スケジュール管理オブジェクトテスト用 */
class schduler {
  constructor(calendarId,firstDay,lastDay){
    this.calendar=CalendarApp.getCalendarById(calendarId) //対象カレンダー
    this.currentRow = 4
  }

  /* 開始日を設定 */
  setFirstDay(firstDay){
    this.firstDay = firstDay 
  }

  /* 終了日を設定 */
  setLastDay(lastDay){
    this.lastDay = lastDay   
  }

  /* 期間内の全予定を取得 */
  fetchSchedules(){
    this.allEvents= this.calendar.getEvents(this.firstDay,this.lastDay)
    }
  
  /* 保有している全予定を返す */
  getEvents(){
    return this.allEvents
  }

  /* 編集をするカレンダーシートを設定 */
  setCalenderSheet(sheetName){
    this.calendarSheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  }

  /* スケジュール対象者の */

  /* 取得している全予定の対象日付をカレンダーシートに記入していく */
  writeAllScheduleX(){
    for (let i =0;i < this.allEvents.length;i++){

      /* 予定の必要情報を抽出 */
      const scheduleTitle= this.allEvents[i].getTitle()
      const startTime = convertDateToString(this.allEvents[i].getStartTime())
      const endTime = convertDateToString(this.allEvents[i].getEndTime())
      const scheduleColor = this.allEvents[i].getColor()
      const scheduleId = this.allEvents[i].getId()

      /* id タイトル記入 */
      this.calendarSheet.getRange(this.currentRow,1).setValue(scheduleId)
      this.calendarSheet.getRange(this.currentRow,2).setValue(scheduleTitle)

      /* xを記入 */
      inputScheduleX(this.calendarSheet,startTime,endTime,this.currentRow)

      this.currentRow+=1

    }
  }

  /* カレンダーシートの初期化 */
  initCalendarSheet(){
    const lastRow=this.calendarSheet.getRange("B:B").getLastRow()
    const lastColumn=this.calendarSheet.getRange("2:2").getLastColumn()
    this.calendarSheet.getRange(4,1,lastRow-3,lastColumn).setValue("")
  }
}

const main=()=>{
  const firstDay = new Date("2024/1/1")
  const lastDay = new Date("2024/12/31")
  const yukiScheduler= new schduler("nun.odakazu@gmail.com")
  yukiScheduler.setFirstDay(firstDay)
  yukiScheduler.setLastDay(lastDay)
  yukiScheduler.fetchSchedules()
  yukiScheduler.setCalenderSheet("calendar")
  yukiScheduler.initCalendarSheet()
  yukiScheduler.writeAllScheduleX()

}

/* シート上に予定の対象日付をxで表記 */
const inputScheduleX=(sheet,startTime,endTime,currentRow)=>{
  let startColumn = findTargetDateColumn(sheet,startTime)
  let endColumn = findTargetDateColumn(sheet,endTime)

  /* 一日のみの予定はとばす */
  if (startColumn==endColumn){
    sheet.getRange(currentRow,startColumn).setValue("x")
    return
  }

  /* x記入 */
  for(let col = startColumn;col<=endColumn;col++){
    sheet.getRange(currentRow,col).setValue("x")
  }

}

/* calendarシート上でyyyy/mm/dd形式の文字列がどの行に該当するか行番号を算出する */
const findTargetDateColumn=(sheet,DateString)=>{
  let targetColumn
  for(let i =1;i<100;i++){
    const targetCell= sheet.getRange(2,i)
    const cellValue = targetCell.getValue()

    /* 日付でないセルならスキップ */
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

/* 日付型の変数を二つ渡すとそれが一致する日付かどうかを判断する。戻り値はboolean */
const judgeMatchDate=(dateA,dateB)=>{

  /* yyyy/mm/dd の形式に日付型を変換 */
  const dateAString = convertDateToString(dateA)
  const dateBString = convertDateToString(dateB)

  /* dateAとdateBの比較 */
  if (dateAString==dateBString){
    return true
  }else{
    return false
  }

}

/* yyyy/mm/dd の形式に日付型を変換 */
const convertDateToString=(targetDate)=>{
  const yyyy= targetDate.getFullYear()
  const mm=  ("0" + (targetDate.getMonth()+1)).slice(-2)
  const dd=("0" + targetDate.getDate()).slice(-2)
  const dateString= `${yyyy}/${mm}/${dd}`
  return dateString
}


















