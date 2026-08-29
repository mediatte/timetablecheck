/**
 * ==========================================================================
 *  2026 이동수업 위치 확인 프로그램 - Code.gs
 * ==========================================================================
 *  "2026 이동수업 위치 확인 프로그램.xlsx" 을 이 스크립트가 바인딩된
 *  구글 시트로 변환/업로드한 뒤 사용합니다. (파일 > Google Sheets로 저장)
 *
 *  구조 요약
 *   - "시간표" 시트   : 학년/반(=이동반 번호)별 요일·교시 칸에
 *                       "A_과목명\n교사명" (이동수업, 앞의 대문자_ 는 슬롯문자)
 *                       또는 "과목명\n교사명" / "자율" / "창체" / "하교" (원반 고정수업)
 *   - "N-M" 시트들    : (예: 2-1, 3-12) 학년 N, 반 M 학생 명단.
 *                       순번/신학번/이름 + (과목, 이동반) 쌍 컬럼들.
 *                       2학년: 타임형A~E (5개)
 *                       3학년: "공학/인공/일/중", "미감비/음감비" 고정 2개
 *                              + 타임형A~F (6개)
 *   - 신학번 5자리 = 학년(1) + 반(2, zero-pad) + 번호(2, zero-pad)
 *
 *  핵심 아이디어 (resolvePeriod_ 기준)
 *   모든 판단은 "학생 자신의 원래 반(homeroom) 칸, 해당 요일/교시 하나"만 본다.
 *   1) 그 칸에 슬롯 문자(A_, B_...)가 붙어 있으면 → 이 시간은 이동수업이다.
 *      학생 명단 시트에서 그 문자에 해당하는 슬롯(타임형X, 또는 3학년의
 *      공학군/미감비군처럼 문자가 헤더에 없는 컬럼)을 찾아 학생 본인의
 *      과목·이동반을 가져와 "이동반 번호 = 실제 위치(교실)"로 표시한다.
 *   2) 문자가 없으면 → 원래 반의 고정 수업 시간이다. "시간표" 시트 그 칸에
 *      적힌 내용(1번째 줄=과목명, 2번째 줄=교사명)을 그대로 과목/교사/위치
 *      (=자기 반 교실)로 사용한다. 자율/창체/하교도 이 경로로 처리된다.
 *      => 이동수업이든 아니든 모든 교시를 항상 채워서 보여준다.
 *
 *   부가: 3학년 "공학/인공/일/중", "미감비/음감비"처럼 헤더에 슬롯 문자가
 *   없는 컬럼은, "시간표"에서 발견됐지만(letterTimeSets) 아직 학생 명단
 *   컬럼과 연결되지 않은 "남는 문자"(보통 G, H)를 헤더 텍스트를 '/' 로 쪼갠
 *   키워드와 실제 과목 텍스트를 대조해 자동으로 매칭해둔다(linkSpecialLetters_).
 *   => 과목/학기가 바뀌어도 헤더 표기 규칙만 유지되면 코드 수정 불필요.
 *
 *  한계 / 유의사항 (테스트 후 필요시 조정)
 *   - "공학/인공/일/중", "미감비/음감비" 같은 고정군 매칭은 자동 추론이므로
 *     학기 초 데이터로 반드시 실제와 대조 확인할 것.
 *   - 이동반 값이 텍스트(예: "교과교실2(3층)")인 경우 해당 반의 시간표 행이
 *     없으므로 담당교사 정보는 표시되지 않는다(위치명만 표시).
 *   - 1학년은 개인별 명단 시트가 없다. 학번이 "1+반(2)+번호(2)" 패턴이면
 *     "시간표" 시트의 해당 학년/반 행만으로 "가상 학생"을 만들어 안내한다
 *     (findGrade1Virtual_). 개인 선택과목이 없으므로 매 시간 원반 교실로 안내됨.
 *   - 조회 결과는 "지금 이 순간의 위치" + "이 학생의 일주일 전체 이동시간표"
 *     (교시 x 요일 그리드, 오늘/현재 교시 강조)를 함께 보여준다.
 * ==========================================================================
 */

var CONFIG = {
  TIMETABLE_SHEET: '시간표',
  ROSTER_SHEET_PATTERN: /^(\d+)-(\d+)$/, // 예: "2-1", "3-12"
  TIMEZONE: 'Asia/Seoul',
  MAX_SEARCH_RESULTS: 30,
  VERSION: '1.2.0',
  VERSION_DATE: '2026-08-29',
  CHANGELOG: [
    { version: '1.0.0', date: '2026-08-29', note: '최초 작성: 학번/이름 일부 검색, 현재 위치 및 오늘 시간표 표시' },
    { version: '1.1.0', date: '2026-08-29', note: '1학년(개인 명단 없음) 학급 시간표 기반 안내 추가, "오늘 시간표"를 "현재 위치 + 일주일 전체 이동시간표(교시x요일 그리드)"로 변경' },
    { version: '1.2.0', date: '2026-08-29', note: 'resolvePeriod_ 단순화: 이동수업 여부를 학생 본인 반 칸의 슬롯 문자 유무로만 판단(다른 반 참고 안 함). 문자가 없는 모든 교시는 시간표 그 칸(과목명/교사명)을 그대로 사용해 고정 시간표도 빠짐없이 표시' }
  ]
};

// ==========================================================================
// 웹앱 진입점
// ==========================================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('이동수업 위치 확인')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getVersionInfo() {
  return {
    version: CONFIG.VERSION,
    date: CONFIG.VERSION_DATE,
    changelog: CONFIG.CHANGELOG
  };
}

// ==========================================================================
// 클라이언트에서 호출하는 API
// ==========================================================================

/**
 * 학번(숫자) 또는 이름 일부로 학생을 검색한다.
 * 결과가 1명이면 바로 상세 정보까지 포함해서 반환한다.
 */
function searchStudent(query, testTime) {
  try {
    var data = buildData_();
    var matches = findStudents_(data, query);

    if (matches.length === 0) {
      var virtual = findGrade1Virtual_(data, query);
      if (virtual) return { status: 'one', detail: buildStudentResult_(data, virtual, testTime) };
      return { status: 'none', message: '일치하는 학생을 찾을 수 없습니다.' };
    }
    if (matches.length === 1) {
      return { status: 'one', detail: buildStudentResult_(data, matches[0], testTime) };
    }
    return {
      status: 'multiple',
      candidates: matches.map(briefStudent_)
    };
  } catch (err) {
    return { status: 'error', message: '오류: ' + (err && err.message ? err.message : err) };
  }
}

/** 검색 결과가 여러 명일 때, 학번으로 상세 정보를 조회한다. */
function getStudentDetail(id, testTime) {
  try {
    var data = buildData_();
    var student = data.roster.byId[String(id)];
    if (!student) student = findGrade1Virtual_(data, id);
    if (!student) return { status: 'none', message: '학번을 찾을 수 없습니다: ' + id };
    return { status: 'one', detail: buildStudentResult_(data, student, testTime) };
  } catch (err) {
    return { status: 'error', message: '오류: ' + (err && err.message ? err.message : err) };
  }
}

/**
 * 1학년은 개인별 이동수업 명단 시트가 없다(이동수업 대상이 아니므로).
 * 학번이 "1+ 반(2자리) + 번호(2자리)" 패턴이면, 학급 전체 시간표("시간표" 시트의
 * 해당 학년/반 행)만으로 안내 가능한 "가상 학생" 객체를 만들어준다.
 * (개인 선택과목 슬롯이 없으므로 student.slots는 항상 빈 배열)
 */
function findGrade1Virtual_(data, query) {
  var q = String(query || '').trim();
  var m = q.match(/^1(\d{2})(\d{2})$/);
  if (!m) return null;
  var homeroom = parseInt(m[1], 10);
  var no = parseInt(m[2], 10);
  if (!data.timetable.grid[1] || !data.timetable.grid[1][homeroom]) return null; // 존재하지 않는 반
  return {
    id: q,
    no: no,
    name: '1학년 ' + homeroom + '반 ' + no + '번',
    grade: 1,
    homeroom: homeroom,
    sheet: null,
    slots: [],
    isVirtual: true
  };
}

// ==========================================================================
// 데이터 로드 (매 요청마다 새로 읽음 - 데이터가 자주 바뀌지 않는 규모라
//  캐싱 없이도 충분히 빠름. 필요시 CacheService 청크 저장으로 확장 가능)
// ==========================================================================

function buildData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('이 스크립트가 스프레드시트에 연결되어 있지 않습니다.');
  var timetable = parseTimetable_(ss);
  var roster = parseRoster_(ss);
  linkSpecialLetters_(timetable, roster);
  return { timetable: timetable, roster: roster };
}

// ---------------------------- 시간표 파싱 --------------------------------

function parseTimetable_(ss) {
  var sheet = ss.getSheetByName(CONFIG.TIMETABLE_SHEET);
  if (!sheet) throw new Error('"' + CONFIG.TIMETABLE_SHEET + '" 시트를 찾을 수 없습니다.');
  var values = sheet.getDataRange().getValues();

  // 헤더 행 찾기: '학년' 과 '반' 이 같은 행에 있는 행
  var headerRowIdx = -1, gradeCol = -1, classCol = -1;
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    var g = values[r].indexOf('학년');
    var c = values[r].indexOf('반');
    if (g !== -1 && c !== -1) { headerRowIdx = r; gradeCol = g; classCol = c; break; }
  }
  if (headerRowIdx === -1) throw new Error('"시간표" 시트에서 헤더(학년/반) 행을 찾지 못했습니다.');

  var dayNameRow = values[headerRowIdx];
  var timeRangeRow = values[headerRowIdx + 1] || [];
  var periodNumRow = values[headerRowIdx + 2] || [];
  var dataStartRow = headerRowIdx + 3;

  // 요일/교시 -> 컬럼 매핑 (요일명은 병합셀이라 오른쪽으로 이어지며 빈칸)
  var colMap = {}; // colIdx -> {day, period}
  var lastCol = dayNameRow.length;
  var curDay = null;
  for (var col = classCol + 1; col < lastCol; col++) {
    var dayVal = dayNameRow[col];
    if (dayVal !== '' && dayVal !== null && dayVal !== undefined) curDay = String(dayVal).trim();
    var pVal = periodNumRow[col];
    var period = null;
    if (typeof pVal === 'number') period = pVal;
    else if (typeof pVal === 'string' && /^\d+$/.test(pVal.trim())) period = parseInt(pVal.trim(), 10);
    if (curDay && period) colMap[col] = { day: curDay, period: period };
  }

  var dayOrder = ['월', '화', '수', '목', '금'];

  // 교시별 시간 (처음 발견된 값 사용)
  var periodTimes = {}; // period -> {startMin, endMin, label}
  Object.keys(colMap).forEach(function (colStr) {
    var col = Number(colStr);
    var period = colMap[col].period;
    if (periodTimes[period]) return;
    var raw = timeRangeRow[col];
    if (typeof raw === 'string' && raw.indexOf('~') !== -1) {
      var parts = raw.split('~');
      var s = toMinutes_(parts[0].trim());
      var e = toMinutes_(parts[1].trim());
      if (s !== null && e !== null) periodTimes[period] = { startMin: s, endMin: e, label: raw.trim() };
    }
  });

  // 데이터 행 순회하며 grid, letterTimeSets 구성
  var grid = {};          // grid[grade][classNum][day][period] = {subject, teacher, letter, raw}
  var letterTimeSets = {}; // letterTimeSets[grade][letter] = ["day,period", ...]
  var curGrade = null;

  for (var row = dataStartRow; row < values.length; row++) {
    var rowVals = values[row];
    var gVal = rowVals[gradeCol];
    if (gVal !== '' && gVal !== null && gVal !== undefined) curGrade = Number(gVal);
    var cVal = rowVals[classCol];
    if (cVal === '' || cVal === null || cVal === undefined) continue;
    var classNum = extractLeadingNumber_(cVal);
    if (classNum === null || curGrade === null) continue;

    grid[curGrade] = grid[curGrade] || {};
    grid[curGrade][classNum] = grid[curGrade][classNum] || {};
    letterTimeSets[curGrade] = letterTimeSets[curGrade] || {};

    Object.keys(colMap).forEach(function (colStr) {
      var col = Number(colStr);
      var cellRaw = rowVals[col];
      if (cellRaw === '' || cellRaw === null || cellRaw === undefined) return;
      var day = colMap[col].day, period = colMap[col].period;
      var parsed = parseCellText_(String(cellRaw));

      grid[curGrade][classNum][day] = grid[curGrade][classNum][day] || {};
      grid[curGrade][classNum][day][period] = parsed;

      if (parsed.letter) {
        letterTimeSets[curGrade][parsed.letter] = letterTimeSets[curGrade][parsed.letter] || [];
        var key = day + ',' + period;
        if (letterTimeSets[curGrade][parsed.letter].indexOf(key) === -1) {
          letterTimeSets[curGrade][parsed.letter].push(key);
        }
      }
    });
  }

  return {
    dayOrder: dayOrder,
    periodTimes: periodTimes,
    grid: grid,
    letterTimeSets: letterTimeSets,
    specialLetterMap: {} // linkSpecialLetters_ 에서 채움: {grade: {letter: label}}
  };
}

function parseCellText_(raw) {
  var lines = raw.split('\n').map(function (s) { return s.trim(); }).filter(function (s) { return s !== ''; });
  var subjectLine = lines[0] || '';
  var teacher = lines[1] || '';
  var m = subjectLine.match(/^([A-Z])_(.+)$/);
  if (m) {
    return { subject: m[2].trim(), teacher: teacher, letter: m[1], raw: raw };
  }
  return { subject: subjectLine, teacher: teacher, letter: null, raw: raw };
}

function toMinutes_(hhmm) {
  var m = hhmm.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

function extractLeadingNumber_(val) {
  if (typeof val === 'number') return val;
  var m = String(val).trim().match(/^(\d+)/);
  return m ? parseInt(m[1], 10) : null;
}

// ---------------------------- 명단 파싱 -----------------------------------

function parseRoster_(ss) {
  var allStudents = [];
  var byId = {};
  var gradeColumns = {}; // grade -> [{type:'letter',letter}|{type:'special',label,keywords}]

  ss.getSheets().forEach(function (sheet) {
    var m = sheet.getName().match(CONFIG.ROSTER_SHEET_PATTERN);
    if (!m) return;
    var grade = parseInt(m[1], 10);
    var homeroom = parseInt(m[2], 10);

    var values = sheet.getDataRange().getValues();
    var headerRowIdx = -1, idCol = -1, nameCol = -1, noCol = -1;
    for (var r = 0; r < Math.min(values.length, 8); r++) {
      var row = values[r].map(function (v) { return String(v || '').trim(); });
      var ic = row.indexOf('신학번');
      if (ic !== -1) {
        headerRowIdx = r; idCol = ic;
        noCol = row.indexOf('순번');
        nameCol = row.indexOf('이름');
        break;
      }
    }
    if (headerRowIdx === -1) return; // 형식이 다른 시트는 건너뜀

    var header = values[headerRowIdx];
    // (과목컬럼, 이동반컬럼) 쌍 찾기: nameCol 다음부터 순회
    var slotDefs = []; // {col, moveCol, type:'letter'|'special', letter?, label?}
    for (var c = nameCol + 1; c < header.length; c++) {
      var h = normalizeHeader_(header[c]);
      if (!h || h === '이동반') continue;
      var lm = h.match(/^타임형([A-Z])$/);
      var moveCol = c + 1;
      if (String(normalizeHeader_(header[moveCol])) !== '이동반') continue; // 짝이 안 맞으면 스킵
      if (lm) {
        slotDefs.push({ col: c, moveCol: moveCol, type: 'letter', letter: lm[1] });
      } else {
        slotDefs.push({ col: c, moveCol: moveCol, type: 'special', label: h });
      }
    }

    if (!gradeColumns[grade]) gradeColumns[grade] = slotDefs;

    for (var rr = headerRowIdx + 1; rr < values.length; rr++) {
      var dr = values[rr];
      var idVal = dr[idCol];
      if (idVal === '' || idVal === null || idVal === undefined) continue; // 명단 끝
      var student = {
        id: String(idVal).trim(),
        no: dr[noCol],
        name: String(dr[nameCol] || '').trim(),
        grade: grade,
        homeroom: homeroom,
        sheet: sheet.getName(),
        slots: []
      };
      slotDefs.forEach(function (def) {
        var subjectRaw = dr[def.col];
        var moveRaw = dr[def.moveCol];
        if (subjectRaw === '' || subjectRaw === null || subjectRaw === undefined) return;
        var moveClass = moveRaw;
        if (typeof moveClass === 'number' && !Number.isInteger(moveClass)) moveClass = Math.round(moveClass);
        var slot = {
          type: def.type,
          subject: String(subjectRaw).trim(),
          moveClass: (typeof moveClass === 'string') ? moveClass.trim() : moveClass
        };
        if (def.type === 'letter') slot.letter = def.letter;
        else slot.label = def.label;
        student.slots.push(slot);
      });

      allStudents.push(student);
      byId[student.id] = student;
    }
  });

  return { allStudents: allStudents, byId: byId, gradeColumns: gradeColumns };
}

function normalizeHeader_(h) {
  return String(h || '').replace(/\s+/g, '');
}

/**
 * "시간표" 상에서 발견됐지만(letterTimeSets) 아직 학생 명단의 '타임형X' 로
 * 연결되지 않은 문자(주로 3학년의 G,H)를, 해당 문자가 붙은 과목 텍스트와
 * 학생 명단의 '특수 컬럼' 헤더(예: "공학/인공/일/중")를 '/' 로 쪼갠
 * 키워드로 대조하여 자동 매칭한다.
 */
function linkSpecialLetters_(timetable, roster) {
  Object.keys(timetable.letterTimeSets).forEach(function (gradeStr) {
    var grade = Number(gradeStr);
    var claimedLetters = {};
    var specialHeaders = []; // {label, keywords}
    (roster.gradeColumns[grade] || []).forEach(function (def) {
      if (def.type === 'letter') claimedLetters[def.letter] = true;
      else if (def.type === 'special' && !specialHeaders.some(function (s) { return s.label === def.label; })) {
        var keywords = def.label.split('/').map(function (s) { return s.replace(/\s+/g, ''); }).filter(Boolean);
        specialHeaders.push({ label: def.label, keywords: keywords });
      }
    });

    var map = {};
    Object.keys(timetable.letterTimeSets[grade]).forEach(function (letter) {
      if (claimedLetters[letter]) return; // 이미 타임형X 로 알려진 문자

      // 이 문자가 붙은 과목 텍스트 샘플 수집
      var samples = {};
      var gridForGrade = timetable.grid[grade] || {};
      Object.keys(gridForGrade).forEach(function (classNum) {
        var byDay = gridForGrade[classNum];
        Object.keys(byDay).forEach(function (day) {
          Object.keys(byDay[day]).forEach(function (period) {
            var cell = byDay[day][period];
            if (cell.letter === letter) samples[cell.subject] = true;
          });
        });
      });
      var sampleTexts = Object.keys(samples);

      var matched = specialHeaders.find(function (sh) {
        return sampleTexts.some(function (txt) {
          return sh.keywords.some(function (kw) { return kw && txt.indexOf(kw) === 0; });
        });
      });
      if (matched) map[letter] = matched.label;
    });

    timetable.specialLetterMap[grade] = map;
  });
}

// ==========================================================================
// 검색 / 조회 로직
// ==========================================================================

function findStudents_(data, query) {
  var q = String(query || '').trim();
  if (!q) return [];
  var list = data.roster.allStudents;

  if (/^\d+$/.test(q)) {
    if (data.roster.byId[q]) return [data.roster.byId[q]];
    return list.filter(function (s) { return s.id.indexOf(q) !== -1; }).slice(0, CONFIG.MAX_SEARCH_RESULTS);
  }
  return list.filter(function (s) { return s.name.indexOf(q) !== -1; }).slice(0, CONFIG.MAX_SEARCH_RESULTS);
}

function briefStudent_(s) {
  return { id: s.id, name: s.name, grade: s.grade, homeroom: s.homeroom, no: s.no };
}

function buildStudentResult_(data, student, testTime) {
  var now = testTime ? new Date(testTime) : new Date();
  var tz = CONFIG.TIMEZONE;
  var dowIso = Number(Utilities.formatDate(now, tz, 'u')); // 1=월 ... 7=일
  var hh = Number(Utilities.formatDate(now, tz, 'HH'));
  var mm = Number(Utilities.formatDate(now, tz, 'mm'));
  var minutes = hh * 60 + mm;
  var dayNames = { 1: '월', 2: '화', 3: '수', 4: '목', 5: '금', 6: '토', 7: '일' };
  var dayLabel = dayNames[dowIso];
  var dayOrder = data.timetable.dayOrder; // ['월','화','수','목','금']
  var todayIsWeekday = dowIso >= 1 && dowIso <= 5;
  var today = todayIsWeekday ? dayOrder[dowIso - 1] : null;

  var result = {
    student: briefStudent_(student),
    generatedAt: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
    usedTestTime: !!testTime,
    dayLabel: dayLabel,
    dayOrder: dayOrder,
    today: today,
    current: null,
    weekSchedule: [],
    specialInfo: [],
    note: student.isVirtual ? '1학년은 개인별 이동수업 명단이 없어, 학급(반) 시간표를 기준으로 안내합니다.' : '',
    version: CONFIG.VERSION
  };

  // 고정군(공학/인공/일/중, 미감비/음감비 등) 참고 정보 - 항상 표시 (해당 학년만)
  student.slots.filter(function (s) { return s.type === 'special'; }).forEach(function (s) {
    result.specialInfo.push({
      label: s.label,
      subject: s.subject,
      location: formatLocation_(student.grade, s.moveClass)
    });
  });

  // ---- 현재 상태 ----
  var periodState = todayIsWeekday ? findPeriodState_(data.timetable.periodTimes, minutes) : { state: 'weekend' };

  if (!todayIsWeekday) {
    result.current = { status: '주말', message: '오늘은 ' + dayLabel + '요일입니다. 수업이 없습니다.' };
  } else if (periodState.state === 'in') {
    result.current = resolvePeriod_(data, student, today, periodState.period);
    result.current.periodLabel = periodState.period + '교시';
    result.current.timeRange = data.timetable.periodTimes[periodState.period].label;
  } else if (periodState.state === 'before') {
    result.current = { status: '등교 전', message: '아직 등교 전 시간입니다.' };
  } else if (periodState.state === 'after') {
    result.current = { status: '방과 후', message: '오늘 수업이 모두 끝났습니다.' };
  } else { // break
    result.current = {
      status: periodState.isLunch ? '점심 시간' : '쉬는 시간',
      message: (periodState.isLunch ? '점심 시간입니다.' : '쉬는 시간입니다.') +
        (periodState.nextPeriod ? ' (다음 ' + periodState.nextPeriod + '교시 대기 중)' : '')
    };
  }

  // ---- 일주일 전체 이동시간표 (교시 x 요일) ----
  var periods = Object.keys(data.timetable.periodTimes).map(Number).sort(function (a, b) { return a - b; });
  periods.forEach(function (p) {
    var row = { period: p, timeRange: data.timetable.periodTimes[p].label, days: {} };
    dayOrder.forEach(function (day) {
      var info = resolvePeriod_(data, student, day, p);
      row.days[day] = {
        status: info.status,
        subject: info.subject || '',
        teacher: info.teacher || '',
        location: info.location || '',
        isCurrent: todayIsWeekday && periodState.state === 'in' && today === day && periodState.period === p
      };
    });
    result.weekSchedule.push(row);
  });

  return result;
}

function findPeriodState_(periodTimes, minutes) {
  var periods = Object.keys(periodTimes).map(Number).sort(function (a, b) { return a - b; });
  if (periods.length === 0) return { state: 'after' };

  if (minutes < periodTimes[periods[0]].startMin) return { state: 'before' };
  var last = periods[periods.length - 1];
  if (minutes > periodTimes[last].endMin) return { state: 'after' };

  for (var i = 0; i < periods.length; i++) {
    var p = periods[i];
    var t = periodTimes[p];
    if (minutes >= t.startMin && minutes <= t.endMin) return { state: 'in', period: p };
  }
  // 교시 사이 쉬는 시간/점심 시간
  for (var j = 0; j < periods.length - 1; j++) {
    var cur = periodTimes[periods[j]], next = periodTimes[periods[j + 1]];
    if (minutes > cur.endMin && minutes < next.startMin) {
      var gap = next.startMin - cur.endMin;
      return { state: 'break', isLunch: gap >= 45, nextPeriod: periods[j + 1] };
    }
  }
  return { state: 'after' };
}

/**
 * 특정 요일/교시에 학생이 어디서 무엇을 하는지 계산한다.
 *
 * 판단 기준은 오직 "학생 자신의 원래 반(homeroom) 칸, 이 요일/교시 하나"뿐이다.
 *  - 그 칸에 슬롯 문자(A_, B_ ...)가 붙어 있으면 → 이동수업 시간 → 학생 명단의
 *    이동반 표(타임형X / 공학군·미감비군 등)를 참고해 실제 위치를 찾는다.
 *  - 문자가 없으면 → 원래 반의 고정 수업 시간 → "전체 시간표" 그 칸에 적힌
 *    내용(1번째 줄=과목명, 2번째 줄=교사명)을 그대로 위치/과목으로 사용한다.
 *    (이동수업 여부와 무관하게, 표시되는 모든 교시를 항상 채운다)
 */
function resolvePeriod_(data, student, day, period) {
  var grid = data.timetable.grid;
  var grade = student.grade;
  var ownCell = grid[grade] && grid[grade][student.homeroom] &&
    grid[grade][student.homeroom][day] && grid[grade][student.homeroom][day][period];

  // "이동수업이다" 라는 판단은 오직 학생 본인 반의 이 칸에 슬롯 문자가
  // 붙어 있는지로만 결정한다 (다른 반 칸은 보지 않음).
  var activeLetter = ownCell && ownCell.letter;

  if (activeLetter) {
    // 1) 학생의 '타임형X' 슬롯에서 직접 매칭
    var slot = student.slots.filter(function (s) { return s.type === 'letter' && s.letter === activeLetter; })[0];
    // 2) 없으면 특수군(G/H 등)으로 연결된 라벨을 통해 매칭
    if (!slot) {
      var label = (data.timetable.specialLetterMap[grade] || {})[activeLetter];
      if (label) slot = student.slots.filter(function (s) { return s.type === 'special' && s.label === label; })[0];
    }
    if (slot) {
      var subjectDisplay = slot.subject.replace(new RegExp('-' + activeLetter + '$'), '');
      var teacher = '';
      if (typeof slot.moveClass === 'number') {
        var destCell = grid[grade] && grid[grade][slot.moveClass] &&
          grid[grade][slot.moveClass][day] && grid[grade][slot.moveClass][day][period];
        if (destCell) teacher = destCell.teacher;
      }
      return {
        status: '이동수업',
        subject: subjectDisplay,
        teacher: teacher,
        location: formatLocation_(grade, slot.moveClass)
      };
    }
    // 학생 데이터에서 이 문자를 찾지 못한 경우 - 원반 칸 내용을 그대로 참고 정보로 표시
    if (ownCell) {
      return {
        status: '이동수업(추정)',
        subject: ownCell.subject,
        teacher: ownCell.teacher,
        location: '확인 필요 (원본 시간표: ' + formatOwnLocation_(grade, student.homeroom) + ')'
      };
    }
  }

  // 슬롯 문자가 없는 일반(원반) 시간
  if (!ownCell) {
    return { status: '정보 없음', subject: '', teacher: '', location: '' };
  }
  if (ownCell.subject === '하교') {
    return { status: '하교', subject: '하교', teacher: '', location: '' };
  }
  if (ownCell.subject === '창체') {
    return { status: '창의적 체험활동', subject: '창체', teacher: ownCell.teacher, location: formatOwnLocation_(grade, student.homeroom) };
  }
  if (ownCell.subject === '자율') {
    return { status: '자율학습', subject: '자율', teacher: ownCell.teacher, location: formatOwnLocation_(grade, student.homeroom) };
  }
  return {
    status: '수업 중(원반)',
    subject: ownCell.subject,
    teacher: ownCell.teacher,
    location: formatOwnLocation_(grade, student.homeroom)
  };
}

/** 이동수업 목적지(이동반 번호 또는 특별실 이름) 표시용 */
function formatLocation_(grade, classNumOrText) {
  if (typeof classNumOrText === 'number') {
    return grade + '학년 ' + classNumOrText + '반 교실 (이동반 ' + classNumOrText + ')';
  }
  return String(classNumOrText || '').trim();
}

/** 학생 자신의 원래 반 교실 표시용 (이동반 문구 없이) */
function formatOwnLocation_(grade, homeroom) {
  return grade + '학년 ' + homeroom + '반 교실';
}
