/** @OnlyCurrentDoc */

function testMacro() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('calendar');
  const finder= sheet.createTextFinder("ゆう")
  const f= finder.findAll()
  console.log(f)
  console.log(f[0].getColumn())
};

const ctest=()=>{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('calendar');
  console.log(sheet.getRange("D2").getValue())
  console.log(convertDateToString(sheet.getRange("D2").getValue()))
}

const inputHolidays=()=>{
  for(let i = 0;i<holidays.length;i++){
    console.log(`${holidays[i].getTitle()}:${convertDateToString(holidays[i].getStartTime())}:${convertDateToString(holidays[i].getEndTime())}` )
  }
}

/* 休日記入クラス */
class holidayWriter{
  constructor(){
  this.holidayCalendar=CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com") 
  }

  //開始日を設定
  setFirstDay(firstDay){
    this.firstDay = firstDay
  }

  //終了日を設定
  setLastDay(lastDay){
    this.lastDay = lastDay   
  }

  getHolidays(){
    this.holidays= this.holidayCalendar.getEvents(this.firstDay,this.lastDay)
  }

  setSheet(sheetName){
    this.targetSheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  }

  //祝日と土日を反映する
  writeHolidays(){

    for (let col = 5;col < 500;col++){
      
      const targetCellValue = this.targetSheet.getRange(2,col).getValue()
      if(targetCellValue==""){
        continue
      }

      //土日に色を付ける
      const targetCellWeekdayNumber =targetCellValue.getDay()
      if(targetCellWeekdayNumber==5 || targetCellWeekdayNumber==6){
        this.targetSheet.getRange(3,col).setBackground("#ea9999")
      }

      //祝日を反映する
      for (let i =0;i <this.holidays.length;i++){
        const targetHoliday= this.holidays[i].getStartTime()
        const targetHolidayEndTimeString= convertDateToString(targetHoliday)

        //何にも書いてないセルなら飛ばす
        if (targetCellValue==""){
          continue
        }

        const targetCellDateString= convertDateToString(targetCellValue)
        console.log(`${targetCellDateString}: ${targetHolidayEndTimeString}`)

        if (targetHolidayEndTimeString==targetCellDateString){
          this.targetSheet.getRange(3,col).setBackground("#ea9999")
          continue
        }
      }
    }

  }
}

const writeHolidays=()=>{
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground('BACKGROUND');
  const writer = new holidayWriter()
  const f = new Date("2024/1/1")
  const l = new Date("2024/12/31")
  writer.setFirstDay(f)
  writer.setLastDay(l)
  writer.getHolidays(f,l)
  writer.setSheet("calendar")
  writer.writeHolidays()
}

/* スケジュール管理オブジェクト */
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

  /* 予定の色番号についてメモ
  何も設定していないデフォルトカラー→空欄
  トマト→11,セージ→2,フラミンゴ→4
   */

  /* 取得している全予定の対象日付をカレンダーシートに記入していく */
  writeAllScheduleX(){
    for (let i =0;i < this.allEvents.length;i++){

      /* 予定の必要情報を抽出 */
      const scheduleTitle= this.allEvents[i].getTitle()
      const startTime = convertDateToString(this.allEvents[i].getStartTime())
      const endTime = convertDateToString(this.allEvents[i].getEndTime())
      const scheduleColor = this.allEvents[i].getColor() //色番号については上記にメモってある
      const scheduleId = this.allEvents[i].getId()
      const isAlldayEvent= this.allEvents[i].isAllDayEvent()

      console.log(`${scheduleTitle} : ${scheduleColor}`)

      /* id タイトル記入 */
      this.calendarSheet.getRange(this.currentRow,1).setValue(scheduleId)
      this.calendarSheet.getRange(this.currentRow,2).setValue(scheduleTitle)

      /* 日付の部分にxを記入 */
      inputScheduleX(this.calendarSheet,startTime,endTime,this.currentRow,isAlldayEvent)

      /* 対象者列にxを記入 */
      this.writeXInTargetUserColumn(this.currentRow,scheduleColor)

      this.currentRow+=1
    }
  }

  /* ゆう・まいの予定対象列を設定 */
  setYumaiColumn(){
    this.yukiColumn = this.calendarSheet.getRange("3:3").createTextFinder("ゆう").findAll()[0].getColumn()
    this.maiColumn = this.calendarSheet.getRange("3:3").createTextFinder("まい").findAll()[0].getColumn()
  }

  /* そのスケジュールの所有者ユーザーの予定対象列を設定 */
  setUserColumn(userName){
    const finder = this.calendarSheet.getRange("3:3").createTextFinder(userName)
    const foundCells= finder.findAll()
    this.userColumn = foundCells[0].getColumn()
  }

  /* 予定の色から対象者列にxを記入 */
  writeXInTargetUserColumn(currentRow,colorNumber){

    /* 色を何も設定していないときは所有者の予定扱いになる */
    if(!colorNumber){
      this.calendarSheet.getRange(currentRow,this.userColumn).setValue("x")
      return
    }

    /* 色が設定されている場合 */
    if(colorNumber==11){ //二人の予定
      this.calendarSheet.getRange(currentRow,this.yukiColumn).setValue("x")
      this.calendarSheet.getRange(currentRow,this.maiColumn).setValue("x")
    }else if(colorNumber==2){ //ゆうきの予定
      this.calendarSheet.getRange(currentRow,this.yukiColumn).setValue("x")
    }else if(colorNumber==4){ //まいの予定
      this.calendarSheet.getRange(currentRow,this.maiColumn).setValue("x")
    }else{ //それ以外の色はとりあえず所有者の予定としておく
      this.calendarSheet.getRange(currentRow,this.userColumn).setValue("x")
    }
  }

  /* カレンダーシートの初期化 */
  initCalendarSheet(){
    const lastRow=this.calendarSheet.getRange("B:B").getLastRow()
    const lastColumn=this.calendarSheet.getRange("2:2").getLastColumn()
    this.calendarSheet.getRange(4,1,lastRow-3,lastColumn).setValue("")
    this.calendarSheet.getRange(4,1,lastRow-3,lastColumn).setBackground("BACKGROUND")
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
  yukiScheduler.setUserColumn("ゆう")
  yukiScheduler.setYumaiColumn()

  yukiScheduler.writeAllScheduleX()
}

/* シート上に予定の対象日付をxで表記 */
const inputScheduleX=(sheet,startTime,endTime,currentRow,isAllDayEvent)=>{
  let startColumn = findTargetDateColumn(sheet,startTime)
  let endColumn = findTargetDateColumn(sheet,endTime)

  //終日予定だとendTimeがその次の日になっている仕様のため-1入れて調整する
  if (isAllDayEvent){
    endColumn-= 1
  }

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
  for(let i =1;i<500;i++){ //とりあえず500
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


















