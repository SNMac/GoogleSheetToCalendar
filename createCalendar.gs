function createCalendar() {
  const T_E_S_T = false
  // 스프레드 시트의 하단에 있는 Sheet tab의 이름
  const SheetTabName = getSheetName();

	// 필요한 데이터의 스프레드 시트 Header명
	const SheetName = "성함";
  const SheetDateOfUse = "이용일";
  const SheetStartTime = "입실"; 
  const SheetEndTime = "퇴실"; 
  const SheetPhoneNumber = "전화번호";
  const SheetMemo = "메모"
  const SheetNumberOfReservation = "예약인원";
  const SheetReferrer = "유입경로";
  const SheetPayment = "결제방식";
  const SheetConfirmed = "예약확정여부";
  const SyncState = "캘린더 등록 상태";       // 업데이트 항목 확인	
  const EventId = "캘린더 이벤트 Id";      // 캘린더 이벤트 Id

	/***************************************************************************************/
	
  // 현재 시트
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(T_E_S_T ? "테스트" : SheetTabName);
  // 설정 시트
  const settingsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설정");
  // 등록할 구글 캘린더의 Id값
  const calendarId = settingsheet.createTextFinder(T_E_S_T ? "테스트 캘린더 Id" : "구글 캘린더 Id").findNext().offset(1, 0).getValue();
  // 동기화할 구글 캘린더
  const eventCal = CalendarApp.getCalendarById(calendarId);
  // 데이터의 시작점(row)
  const startRow = 3;
	// 데이터의 시작점(column)
  const startColumn = 1;
  // 데이터의 끝점(row)
  const endRow = spreadsheet.getLastRow();
  // 데이터의 끝점(column)
  const endColumn = spreadsheet.getLastColumn();

  const count = spreadsheet.getRange(startRow, startColumn, endRow, endColumn).getValues();  // getRange(row, column, numRows, numColumns)

	// 해당 header의 스프레드 시트 ColumnIndex
  const colSheetName = SheetName ? spreadsheet.createTextFinder(SheetName).findNext().getColumnIndex() - startColumn : "";
  const colSheetDateOfUse = spreadsheet.createTextFinder(SheetDateOfUse).findNext().getColumnIndex() - startColumn;
  const colSheetStartTime = SheetStartTime ? spreadsheet.createTextFinder(SheetStartTime).findNext().getColumnIndex() - startColumn : "";
  const colSheetEndTime = SheetEndTime ? spreadsheet.createTextFinder(SheetEndTime).findNext().getColumnIndex() - startColumn : "";
  const colSheetPhoneNumber = SheetPhoneNumber ? spreadsheet.createTextFinder(SheetPhoneNumber).findNext().getColumnIndex() - startColumn : "";
  const colSheetMemo = SheetMemo ? spreadsheet.createTextFinder(SheetMemo).findNext().getColumnIndex() - startColumn : "";
  const colSheetNumberOfReservation = SheetNumberOfReservation ? spreadsheet.createTextFinder(SheetNumberOfReservation).findNext().getColumnIndex() - startColumn : "";
  const colSheetReferrer = SheetReferrer ? spreadsheet.createTextFinder(SheetReferrer).findNext().getColumnIndex() - startColumn : "";
  const colSheetPayment = SheetPayment ? spreadsheet.createTextFinder(SheetPayment).findNext().getColumnIndex() - startColumn : "";
  const colSheetConfirmed = SheetConfirmed ? spreadsheet.createTextFinder(SheetConfirmed).findNext().getColumnIndex() - startColumn : "";
  const colSyncState = SyncState ? spreadsheet.createTextFinder(SyncState).findNext().getColumnIndex() - startColumn : "";
  const colEventId = EventId ? spreadsheet.createTextFinder(EventId).findNext().getColumnIndex() - startColumn : "";
    
  const KR_TIME_DIFF = 9 * 60 * 60 * 1000;

	// 시트 Row 수 만큼 반복
  for (x = 0; x < count.length; x++) {
    // 한꺼번에 많은 캘린더를 등록하면 오류가 발생함
    if (x === 15) Utilities.sleep(2 * 1000);

    const shift = count[x]; // 시트 row (하단 설명 참고)

    // 이용일 설정
    const sheetDateOfUse = new Date(shift[colSheetDateOfUse]);
    const dateOfUseYear = sheetDateOfUse.getFullYear();
    const dateOfUseMonth = sheetDateOfUse.getMonth();
    const dateOfUseDate = sheetDateOfUse.getDate();
				
		// 한국 시간은 UTC보다 9시간 빠름
    const sheetStartTime = new Date(shift[colSheetStartTime]);
    const sheetStartTimeUTC = sheetStartTime.getTime() + sheetStartTime.getTimezoneOffset() * 60 * 1000;
    const sheetStartTimeKST = new Date(sheetStartTimeUTC + KR_TIME_DIFF);
    const sheetStartHour = sheetStartTimeKST.getHours();
    const sheetStartMin = sheetStartTimeKST.getMinutes();
    const calStartTime = new Date(dateOfUseYear, dateOfUseMonth, dateOfUseDate, sheetStartHour, sheetStartMin, 0);
    const sheetEndTime = new Date(shift[colSheetEndTime]);
    const sheetEndTimeUTC = sheetEndTime.getTime() + sheetEndTime.getTimezoneOffset() * 60 * 1000;
    const sheetEndTimeKST = new Date(sheetEndTimeUTC + KR_TIME_DIFF);
    const sheetEndHour = sheetEndTimeKST.getHours();
    const sheetEndMin = sheetEndTimeKST.getMinutes();
    const calEndTime = new Date(dateOfUseYear, dateOfUseMonth, sheetEndHour == 0 ? dateOfUseDate + 1 : dateOfUseDate, sheetEndHour, sheetEndMin, 0);
				
		// var data = shift[가져오고 싶은 header ColumnIndex];

    const sheetName = shift[colSheetName] ? shift[colSheetName] : "";
    const sheetPhoneNumber = shift[colSheetPhoneNumber] ? shift[colSheetPhoneNumber] : "";
    const sheetMemo = shift[colSheetMemo] ? shift[colSheetMemo] : "";
    const sheetNumberOfReservation = shift[colSheetNumberOfReservation] ? shift[colSheetNumberOfReservation] : "";
    const sheetReferrer = shift[colSheetReferrer] ? shift[colSheetReferrer] : "";
    const sheetPayment = shift[colSheetPayment] ? shift[colSheetPayment] : "";
    const sheetConfirmed = shift[colSheetConfirmed] ? shift[colSheetConfirmed] : "";

    // 캘린더 해당 이벤트 Id (캘린더 내용 수정 시 필요)
    const calId = shift[colEventId] ? shift[colEventId] : "";
		// 캘린더 이벤트 등록 상태
    const calSyncState = shift[colSyncState] ? shift[colSyncState] : "";

    const titleSum = (sheetConfirmed == "입금대기" ? "(대기)" : "") + sheetName + (sheetNumberOfReservation == 0 ? "" : " " + sheetNumberOfReservation + "인");
    const descriptionSum = "전화번호: " + sheetPhoneNumber + "\n메모: " + sheetMemo + "\n유입경로: " + sheetReferrer + "\n결제방식: " + sheetPayment;

		// 캘린더에 등록할 이벤트 내용
    const event = {
      description: descriptionSum
    };

    if (sheetConfirmed == "예약확정" || sheetConfirmed == "입금대기") {  // 캘린더에 이벤트 등록
      // 캘린더 등록 상태가 "Y"가 아니라면, 등록
      if (calSyncState != "Y" && !isNaN(sheetDateOfUse) && sheetName !== "") {

        if(calId !== null && calId !== "") {
          // 기존 캘린터 이벤트 삭제
          eventCal.getEventById(calId).deleteEvent();
        }

        let newEvent = eventCal.createEvent(titleSum, calStartTime, calEndTime, event);

        spreadsheet.getRange(Number(startRow + x), colSyncState + startColumn).setValue("Y");  // 등록 상태 "Y"로 업데이트
        spreadsheet.getRange(Number(startRow + x), colEventId + startColumn).setValue(newEvent.getId());  // 새로 등록한 이벤트 Id값 입력

      // 캘린더 등록 상태가 "Y"라면, 변경사항 확인 후 시트상의 데이터로 반영
      } else {
        if(calId !== null && calId !== "") {
          // 변경 사항이 있는 경우
          if (eventCal.getEventById(calId).getTitle() != titleSum
          || eventCal.getEventById(calId).getStartTime() != calStartTime
          || eventCal.getEventById(calId).getEndTime() != calEndTime
          || eventCal.getEventById(calId).getDescription() != descriptionSum) {
            // 기존 캘린터 이벤트 삭제
            eventCal.getEventById(calId).deleteEvent();
            // 이벤트 재등록
            let newEvent = eventCal.createEvent(titleSum, calStartTime, calEndTime, event);

            spreadsheet.getRange(Number(startRow + x), colSyncState + startColumn).setValue("Y");  // 등록 상태 "Y"로 업데이트
            spreadsheet.getRange(Number(startRow + x), colEventId + startColumn).setValue(newEvent.getId());  // 새로 등록한 이벤트 Id값 입력
          }
        }
      }
    } else if (sheetConfirmed == "예약취소") {  // 캘린더에서 이벤트 삭제
    // 캘린더 등록 상태가 "Y"라면, 삭제
      if (calSyncState == "Y") {
        const events = eventCal.getEvents(calStartTime, calEndTime, { search: titleSum });
        for (y = 0; y < events.length; y++) {
          events[y].deleteEvent();
        }
        spreadsheet.getRange(Number(startRow + x), colSyncState + startColumn).setValue("");  // 등록 상태 삭제
        spreadsheet.getRange(Number(startRow + x), colEventId + startColumn).setValue("");  // 이벤트 Id값 삭제
      }
    }
  }
}

// 스프레드 시트에 버튼 생성
function onOpenSchedule() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("캘린더 동기화").addItem("업데이트", "createCalendar").addToUi();
}
