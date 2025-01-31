function addEventsToCalendar() {
    const sheetName = '일정'; // 실제 시트 이름으로 변경
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const calendarId = 'primary'; // 기본 캘린더 사용
    const calendar = CalendarApp.getCalendarById(calendarId);
    const userEmail = ''; // 알림을 받을 이메일 주소로 변경
  
    // Google Calendar 색상 ID를 색상 이름으로 매핑
    const colorIdMap = {
      '1': '기본',        // 라벤더
      '2': '연두',        // 세이지
      '3': '보라',        // 자두
      '4': '주황',        // 플라밍고
      '5': '노랑',        // 바나나
      '6': '진한 주황',   // 탠저린
      '7': '하늘',        // 공작
      '8': '회색',        // 그래파이트g
      '9': '남색',        // 바질
      '10': '초록',       // 토마토
      '11': '빨강'        // 인디고
      // 필요에 따라 추가 색상 ID 및 이름을 여기에 추가하세요
    };
  
  
    try {
      if (!calendar) {
        throw new Error('캘린더를 찾을 수 없습니다. 캘린더 ID를 확인해주세요.');
      }
  
      if (!sheet) {
        throw new Error(`시트 "${sheetName}"을(를) 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
      }
  
      const dataRange = sheet.getDataRange();
      const data = dataRange.getValues();
  
      let addedEvents = [];
      let skippedEvents = [];
  
      for (let i = 1; i < data.length; i++) { // 첫 번째 행은 헤더이므로 1부터 시작
        const row = data[i];
        const subject = row[0];
        const startDate = row[1];
        const startTime = row[2];
        const endDate = row[3];
        const endTime = row[4];
        const description = row[5];
        const location = row[6];
        const colorId = row[7]; // Color ID는 8번째 열 (H 열)
  
        if (!subject || !startDate) {
          continue; // 제목 또는 시작 날짜가 없는 경우 건너뜁니다
        }
  
        let isAllDay = false;
        let startDateTime, endDateTime;
  
        // 날짜 및 시간 설정
        if (startTime && endTime) {
          // 시작 시간과 종료 시간이 모두 있는 경우
          try {
            startDateTime = combineDateTime(startDate, startTime);
            endDateTime = combineDateTime(endDate, endTime);
          } catch (e) {
            skippedEvents.push({
              subject: subject,
              reason: `시간 형식 오류: ${e.message}`
            });
            Logger.log(`이벤트 건너뜀: ${subject} - 시간 형식 오류.`);
            continue;
          }
        } else {
          // 시간 정보가 없으면 종일 일정으로 설정
          isAllDay = true;
          startDateTime = new Date(startDate);
          // 종일 일정의 경우, 종료 날짜는 시작 날짜의 다음 날로 설정
          endDateTime = new Date(startDate);
          endDateTime.setDate(endDateTime.getDate() + 1);
        }
  
        // 날짜 검증: 시작 날짜가 종료 날짜보다 이전인지 확인
        if (startDateTime >= endDateTime) {
          skippedEvents.push({
            subject: subject,
            reason: '시작 날짜/시간이 종료 날짜/시간과 같거나 늦음.'
          });
          Logger.log(`이벤트 건너뜀: ${subject} (${startDateTime} - ${endDateTime}) - 시작이 종료보다 늦거나 같음.`);
          continue;
        }
  
        // 이벤트 중복 확인 (제목과 시작 시간이 동일한 이벤트를 찾음)
        const existingEvents = calendar.getEvents(startDateTime, endDateTime, {search: subject});
        if (existingEvents.length > 0) {
          Logger.log(`이미 존재하는 이벤트: ${subject} (${startDateTime})`);
          continue;
        }
  
        // 이벤트 생성 옵션 설정
        let eventOptions = {
          description: description || '없음',
          location: location || '없음',
        };
  
        if (colorId) {
          eventOptions.color = colorId.toString(); // 색상 ID는 문자열이어야 합니다
        }
  
        let event;
        if (isAllDay) {
          event = calendar.createAllDayEvent(subject, startDateTime, endDateTime, eventOptions);
        } else {
          event = calendar.createEvent(subject, startDateTime, endDateTime, eventOptions);
        }
  
        // 색상 적용 (createEvent와 createAllDayEvent에서 color 옵션을 지원함)
        if (colorId) {
          event.setColor(colorId.toString());
        }
  
        // 날짜 형식 변환 (한국어 형식)
        const formattedStart = Utilities.formatDate(startDateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const formattedEnd = Utilities.formatDate(endDateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const formattedStartWithTime = Utilities.formatDate(startDateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        const formattedEndWithTime = Utilities.formatDate(endDateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  
        addedEvents.push({
          subject: subject,
          start: isAllDay ? formattedStart : formattedStartWithTime,
          end: isAllDay ? formattedEnd : formattedEndWithTime,
          description: description || '없음',
          colorName: colorIdMap[colorId.toString()] || '기본 색상'
        });
  
        Logger.log(`이벤트 추가됨: ${subject} (${startDateTime} - ${endDateTime}) 색상: ${colorIdMap[colorId.toString()] || '기본 색상'}`);
      }
  
      // 이메일 전송 조건 변경: 추가된 일정이 있거나, 오류가 있는 경우에만 이메일 전송
      if (addedEvents.length > 0 || skippedEvents.length > 0) {
        let emailBody = '';
  
        if (addedEvents.length > 0) {
          emailBody += '다음 일정이 Google Calendar에 성공적으로 추가되었습니다:\n\n';
          addedEvents.forEach(event => {
            emailBody += `제목: ${event.subject}\n시작: ${event.start}\n종료: ${event.end}\n설명: ${event.description}\n색상: ${event.colorName}\n\n`;
          });
        }
  
        if (skippedEvents.length > 0) {
          if (addedEvents.length > 0) {
            emailBody += '-----------------------------\n';
          }
          emailBody += '추가되지 않은 일정:\n\n';
          skippedEvents.forEach(event => {
            emailBody += `제목: ${event.subject}\n사유: ${event.reason}\n\n`;
          });
        }
  
        // 이메일 제목 결정
        let emailSubject = 'Google Calendar 일정 추가';
        if (addedEvents.length > 0 && skippedEvents.length > 0) {
          emailSubject += ' (일부 성공)';
        } else if (addedEvents.length > 0) {
          emailSubject += ' 성공';
        } else if (skippedEvents.length > 0) {
          emailSubject += ' 오류 발생';
        }
  
        MailApp.sendEmail(userEmail, emailSubject, emailBody);
      }
  
      SpreadsheetApp.getUi().alert('일정이 성공적으로 추가되었습니다.');
  
    } catch (error) {
      Logger.log(`오류 발생: ${error}`);
      MailApp.sendEmail(userEmail, 'Google Calendar Sync 오류', `스크립트 실행 중 오류가 발생했습니다:\n\n${error}`);
      SpreadsheetApp.getUi().alert('일정 추가 중 오류가 발생했습니다. 관리자에게 문의하세요.');
    }
  }
  
  // 날짜와 시간을 결합하는 헬퍼 함수
  function combineDateTime(date, time) {
    if (!(time instanceof Date)) {
      throw new Error('시간 데이터가 Date 형식이 아닙니다.');
    }
  
    // 시간 객체에서 '오전' 또는 '오후'는 이미 반영된 상태입니다.
    let hours = time.getHours();
    let minutes = time.getMinutes();
  
    // 24시간 형식을 사용하여 시간 설정
    let dateTime = new Date(date);
    dateTime.setHours(hours);
    dateTime.setMinutes(minutes);
    dateTime.setSeconds(0);
    dateTime.setMilliseconds(0);
  
    return dateTime;
  }
  