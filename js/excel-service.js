/**
 * Excel Service using ExcelJS for rich styling, pastel colors, and multi-sheet support.
 * - 학생 편성 엑셀 업로드 시 [학번, 성명]만 추출
 * - 3단 상태(🟢 확정 / 🔺 조정 / ✖️ 폐강 - 회색) 서식 지원
 * - 파스텔톤 멀티시트 학급 시간표 및 원천 데이터 내보내기
 */

const ExcelService = {
  dayNameToId(name) {
    if (!name) return null;
    const clean = String(name).trim();
    const map = {
      '월': 'mon', '월요일': 'mon', 'mon': 'mon',
      '화': 'tue', '화요일': 'tue', 'tue': 'tue',
      '수': 'wed', '수요일': 'wed', 'wed': 'wed',
      '목': 'thu', '목요일': 'thu', 'thu': 'thu',
      '금': 'fri', '금요일': 'fri', 'fri': 'fri',
      '토': 'sat', '토요일': 'sat', 'sat': 'sat',
      '일': 'sun', '일요일': 'sun', 'sun': 'sun'
    };
    return map[clean] || null;
  },

  dayIdToName(id) {
    const map = { mon: '월', tue: '화', wed: '수', thu: '목', fri: '금', sat: '토', sun: '일' };
    return map[id] || '';
  },

  // 5자리 학번 생성: 1학년 1반 1번 -> 10101
  getStudentCode(student) {
    const g = student.grade ? String(student.grade).trim() : '1';
    const c = student.classNum !== undefined && student.classNum !== null ? String(student.classNum).trim().padStart(2, '0') : '01';
    if (student.studentNum !== undefined && student.studentNum !== null && String(student.studentNum).trim() !== '') {
      const n = String(student.studentNum).trim().padStart(2, '0');
      return `${g}${c}${n}`;
    }
    return `${g}${c}`;
  },

  // 학번이름: 10101홍길동
  getStudentCodeName(student) {
    const code = this.getStudentCode(student);
    return `${code}${student.name || ''}`;
  },

  // 상태별 기호 (확정: 🟢, 조정: 노란 세모 ▲, 폐강: ✖️)
  getStatusSymbol(status) {
    if (status === 'adjust') return '▲';
    if (status === 'canceled') return '✖️';
    return '🟢';
  },

  // 상태별 HTML 표기 (노란색 세모 스타일 적용)
  getStatusHtml(status) {
    if (status === 'adjust') {
      return '<span class="text-amber-500 font-black drop-shadow-xs" style="color: #F59E0B;" title="조정(노랑 세모)">▲</span>';
    }
    if (status === 'canceled') {
      return '<span class="text-slate-400 font-bold" title="폐강(회색 엑스)">✖️</span>';
    }
    return '<span class="text-emerald-500" title="확정(초록)">🟢</span>';
  },

  // 상태 텍스트
  getStatusText(status) {
    if (status === 'adjust') return '조정(▲)';
    if (status === 'canceled') return '폐강(✖️)';
    return '확정(🟢)';
  },

  // 상태 문자열 표준화 (개설/확정, 조정, 폐강)
  normalizeStatus(val) {
    if (!val) return { courseStatus: 'open', scheduleStatus: 'confirmed', label: '확정' };
    const str = String(val).trim().toLowerCase();
    if (str.includes('조정') || str.includes('▲') || str.includes('adjust')) {
      return { courseStatus: 'adjust', scheduleStatus: 'adjust', label: '조정' };
    }
    if (str.includes('폐강') || str.includes('✖') || str.includes('x') || str.includes('closed') || str.includes('cancel')) {
      return { courseStatus: 'closed', scheduleStatus: 'canceled', label: '폐강' };
    }
    return { courseStatus: 'open', scheduleStatus: 'confirmed', label: '확정' };
  },

  // 시간표 카드 표기: 과목명(학번성명)
  getTimetableCardLabel(student) {
    const sub = student.subject || '과목';
    const codeName = this.getStudentCodeName(student);
    const sym = this.getStatusSymbol(student.status);
    const memoPart = student.memo ? `[${student.memo}] ` : '';
    return `${sym} ${memoPart}${sub}(${codeName})`;
  },

  // 1. 일과표 양식 다운로드
  async downloadPeriodTemplate() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '강원온라인학교 시간표 시스템';
    const worksheet = workbook.addWorksheet('일과표 설정');

    worksheet.columns = [
      { header: '교시구분', key: 'type', width: 14 },
      { header: '교시번호', key: 'period', width: 12 },
      { header: '교시명', key: 'name', width: 16 },
      { header: '시작시간(HH:mm)', key: 'startTime', width: 18 },
      { header: '종료시간(HH:mm)', key: 'endTime', width: 18 },
      { header: '비고', key: 'memo', width: 25 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.font = { bold: true, color: { argb: 'FF1E293B' }, size: 11 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const sampleRows = [
      { type: '수업', period: 1, name: '1교시', startTime: '09:00', endTime: '09:50', memo: '50분 수업 / 10분 휴식' },
      { type: '수업', period: 1, name: '1교시(선택)', startTime: '09:30', endTime: '10:20', memo: '동일 교시 시간차 예시' },
      { type: '수업', period: 2, name: '2교시', startTime: '10:00', endTime: '10:50', memo: '' },
      { type: '수업', period: 3, name: '3교시', startTime: '11:00', endTime: '11:50', memo: '' },
      { type: '수업', period: 4, name: '4교시', startTime: '12:00', endTime: '12:50', memo: '' },
      { type: '휴식/점심', period: '', name: '점심시간', startTime: '12:50', endTime: '13:40', memo: '점심시간 50분' },
      { type: '수업', period: 5, name: '5교시', startTime: '13:40', endTime: '14:30', memo: '' },
      { type: '수업', period: 6, name: '6교시', startTime: '14:40', endTime: '15:30', memo: '' },
      { type: '수업', period: 7, name: '7교시', startTime: '15:40', endTime: '16:30', memo: '' },
      { type: '수업', period: 8, name: '8교시', startTime: '16:40', endTime: '17:30', memo: '기본 8교시' },
      { type: '수업', period: 9, name: '9교시', startTime: '17:40', endTime: '18:30', memo: '야간 교시 (최대 21:50)' },
    ];

    sampleRows.forEach((row) => {
      const addedRow = worksheet.addRow(row);
      addedRow.height = 24;
      addedRow.eachCell((cell) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        this.applyBorder(cell);
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, '일과표_설정_양식.xlsx');
  },

  // 2. 일과표 엑셀 업로드 파싱
  async parsePeriodExcel(file) {
    const workbook = new ExcelJS.Workbook();
    const buffer = await file.arrayBuffer();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];

    const periods = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const type = String(row.getCell(1).value || '').trim();
      const periodVal = row.getCell(2).value;
      const name = String(row.getCell(3).value || '').trim();
      let startTime = String(row.getCell(4).value || '').trim();
      let endTime = String(row.getCell(5).value || '').trim();

      if (!name && !startTime) return;

      const isBreak = type.includes('점심') || type.includes('휴식') || name.includes('점심') || name.includes('휴식');
      const numPeriod = isBreak ? null : (parseInt(periodVal, 10) || (periods.filter(p => !p.isBreak).length + 1));

      startTime = this.formatTimeString(startTime);
      endTime = this.formatTimeString(endTime);

      periods.push({
        id: `p_uploaded_${Date.now()}_${rowNumber}`,
        period: numPeriod,
        name: name || (isBreak ? '점심시간' : `${numPeriod}교시`),
        startTime: startTime || '09:00',
        endTime: endTime || '09:50',
        isBreak: isBreak
      });
    });

    if (periods.length === 0) {
      throw new Error('유효한 일과표 데이터를 찾을 수 없습니다.');
    }
    return periods;
  },

  // 3. 수강생 명단 및 시간표 통합 표준 양식 다운로드 (일원화 양식)
  // 3. 수강생 명단 및 시간표 통합 표준 양식 다운로드 (1·2학기 전체 포함 빈 양식)
  async downloadStudentTemplate(academicYear = '2026') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '강원온라인학교 시간표 시스템';
    const worksheet = workbook.addWorksheet(`${academicYear}학년도_1·2학기_통합데이터`, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 10 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '개설상태', key: 'status', width: 12 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '시작시간', key: 'startTime', width: 12 },
      { header: '종료시간', key: 'endTime', width: 12 },
      { header: '시간중복', key: 'conflict', width: 22 },
      { header: '비고', key: 'memo', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    // 1학기 빈 서식 10줄 (순번 1~10)
    for (let i = 1; i <= 10; i++) {
      const addedRow = worksheet.addRow({
        index: i,
        year: academicYear,
        term: '1학기',
        grade: '',
        classNum: '',
        studentNum: '',
        studentCode: '',
        name: '',
        subject: '',
        status: '',
        day: '',
        periodName: '',
        startTime: '',
        endTime: '',
        conflict: '',
        memo: ''
      });
      addedRow.height = 24;
      addedRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);
        if (colNum === 3) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
          cell.font = { bold: true, color: { argb: 'FF0369A1' } };
        }
      });
    }

    // 2학기 빈 서식 10줄 (순번 11~20)
    for (let i = 11; i <= 20; i++) {
      const addedRow = worksheet.addRow({
        index: i,
        year: academicYear,
        term: '2학기',
        grade: '',
        classNum: '',
        studentNum: '',
        studentCode: '',
        name: '',
        subject: '',
        status: '',
        day: '',
        periodName: '',
        startTime: '',
        endTime: '',
        conflict: '',
        memo: ''
      });
      addedRow.height = 24;
      addedRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);
        if (colNum === 3) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDE9FE' } };
          cell.font = { bold: true, color: { argb: 'FF6D28D9' } };
        }
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${academicYear}학년도_1·2학기_시간표_통합데이터_표준양식.xlsx`);
  },

  // 4. 수강생 명단 및 시간표 통합 엑셀 업로드 파싱 (통합 양식 & 나이스 양식 완벽 호환)
  async parseStudentExcel(file) {
    const workbook = new ExcelJS.Workbook();
    const buffer = await file.arrayBuffer();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];

    let headerRowIndex = -1;
    let colIndex = -1;
    let colYear = -1;
    let colSemester = -1;
    let colGrade = -1;
    let colClass = -1;
    let colNumber = -1;
    let colStudentCode = -1;
    let colName = -1;
    const colSubjects = [];
    let colStatus = -1;
    let colDay = -1;
    let colPeriod = -1;
    let colStartTime = -1;
    let colEndTime = -1;
    let colMemo = -1;

    // 헤더 행 탐색 (최대 15행 스캔)
    const maxScanRows = Math.min(15, worksheet.rowCount);
    for (let r = 1; r <= maxScanRows; r++) {
      const row = worksheet.getRow(r);
      const cellEntries = [];
      row.eachCell((cell, colNum) => {
        const text = String(cell.value || '').trim();
        if (text) cellEntries.push({ colNum, text });
      });

      const hasGrade = cellEntries.some(c => c.text === '학년' || c.text.toLowerCase() === 'grade');
      const hasCode = cellEntries.some(c => c.text === '학번' || c.text === '학생번호' || c.text.toLowerCase() === 'studentcode' || c.text.toLowerCase() === 'id');
      const hasName = cellEntries.some(c => c.text === '성명' || c.text === '이름' || c.text === '학생명' || c.text.toLowerCase() === 'name');

      if ((hasGrade || hasCode) && hasName) {
        headerRowIndex = r;
        cellEntries.forEach(c => {
          const t = c.text;
          const lower = t.toLowerCase();
          if (t === '순번' || lower === 'no' || lower === 'idx' || lower === 'index') {
            colIndex = c.colNum;
          } else if (t === '학년도' || t === '연도' || lower === 'year') {
            colYear = c.colNum;
          } else if (t === '학기' || lower === 'term' || lower === 'semester') {
            colSemester = c.colNum;
          } else if (t === '학년' || lower === 'grade') {
            colGrade = c.colNum;
          } else if (t === '반' || lower === 'class') {
            colClass = c.colNum;
          } else if (t === '번호' || t === '번' || lower === 'studentnum' || (lower === 'no' && colIndex !== c.colNum)) {
            colNumber = c.colNum;
          } else if (t === '학번' || t === '학생번호' || lower === 'studentcode' || lower === 'id') {
            colStudentCode = c.colNum;
          } else if (t === '성명' || t === '이름' || t === '학생명' || lower === 'name') {
            colName = c.colNum;
          } else if (t.includes('과목') || t.includes('교과') || t.includes('신청') || t.includes('수강') || t.includes('희망')) {
            colSubjects.push(c.colNum);
          } else if (t === '개설상태' || t === '상태' || t === '구분' || lower === 'status') {
            colStatus = c.colNum;
          } else if (t === '배정요일' || t === '요일' || lower === 'day') {
            colDay = c.colNum;
          } else if (t === '배정교시' || t === '교시' || t === '교시명' || lower === 'period') {
            colPeriod = c.colNum;
          } else if (t === '시작시간' || t === '시작' || lower === 'starttime' || lower === 'start') {
            colStartTime = c.colNum;
          } else if (t === '종료시간' || t === '종료' || lower === 'endtime' || lower === 'end') {
            colEndTime = c.colNum;
          } else if (t.includes('중복') || lower.includes('conflict')) {
            // 시간중복 열은 안내용이므로 과목명이나 기타 항목으로 오인되지 않도록 안전하게 분기
          } else if (t === '비고' || t === '강의실' || t === '개설강의실' || lower === 'memo' || lower === 'room') {
            colMemo = c.colNum;
          }
        });
        break;
      }
    }

    // 헤더를 못 찾은 경우 기본 NEIS 수강편성 양식 위치 적용
    if (headerRowIndex === -1) {
      headerRowIndex = 2;
      colSubjects.push(5); // 개설과목(학점)
      colMemo = 6; // 개설강의실
      colGrade = 3;
      colClass = 8;
      colNumber = 9;
      colName = 10;
    }

    // 학생별 통합 Map (학번_성명 기준)
    const studentMap = new Map();

    for (let r = headerRowIndex + 1; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const rawName = colName > 0 ? String(row.getCell(colName).value || '').trim() : '';
      if (!rawName || rawName === '성명' || rawName === '이름' || rawName === '학생명') continue;

      let grade = 1;
      let classNum = 1;
      let studentNum = null;
      let studentCode = '';

      const rawCode = colStudentCode > 0 ? String(row.getCell(colStudentCode).value || '').trim() : '';
      const rawGrade = colGrade > 0 ? row.getCell(colGrade).value : null;
      const rawClass = colClass > 0 ? row.getCell(colClass).value : null;
      const rawNumber = colNumber > 0 ? row.getCell(colNumber).value : null;

      if (rawGrade !== null && rawGrade !== undefined && String(rawGrade).trim() !== '') {
        grade = parseInt(String(rawGrade).replace(/[^0-9]/g, ''), 10) || 1;
      }
      if (rawClass !== null && rawClass !== undefined && String(rawClass).trim() !== '') {
        classNum = parseInt(String(rawClass).replace(/[^0-9]/g, ''), 10) || 1;
      }
      if (rawNumber !== null && rawNumber !== undefined && String(rawNumber).trim() !== '') {
        studentNum = parseInt(String(rawNumber).replace(/[^0-9]/g, ''), 10) || null;
      }

      // 학번 열이 존재하고 4~5자리 숫자인 경우 학년/반/번호 보정
      const cleanRawCode = rawCode.replace(/[^0-9]/g, '');
      if (cleanRawCode.length >= 4 && cleanRawCode.length <= 5) {
        if (!rawGrade && cleanRawCode.length === 5) grade = parseInt(cleanRawCode[0], 10) || grade;
        if (!rawClass && cleanRawCode.length === 5) classNum = parseInt(cleanRawCode.slice(1, 3), 10) || classNum;
        if (!studentNum && cleanRawCode.length === 5) studentNum = parseInt(cleanRawCode.slice(3), 10) || studentNum;
        studentCode = cleanRawCode;
      } else {
        studentCode = this.getStudentCode({ grade, classNum, studentNum });
      }

      // 상태 추출
      const rawStatusVal = colStatus > 0 ? row.getCell(colStatus).value : null;
      const normStatus = this.normalizeStatus(rawStatusVal);

      // 요일 추출 ('미입력', '-', '미배정' 등은 null 처리)
      let parsedDay = null;
      if (colDay > 0) {
        const rawDayVal = String(row.getCell(colDay).value || '').trim();
        if (rawDayVal && rawDayVal !== '미입력' && rawDayVal !== '미배정' && rawDayVal !== '-' && rawDayVal !== '없음') {
          parsedDay = this.dayNameToId(rawDayVal);
        }
      }

      // 교시 추출 ('미입력', '-', '미배정' 등은 null 처리)
      let parsedPeriod = null;
      let parsedPeriodName = null;
      if (colPeriod > 0) {
        const rawPeriodVal = String(row.getCell(colPeriod).value || '').trim();
        if (rawPeriodVal && rawPeriodVal !== '미입력' && rawPeriodVal !== '미배정' && rawPeriodVal !== '-' && rawPeriodVal !== '없음') {
          const num = parseInt(rawPeriodVal.replace(/[^0-9]/g, ''), 10);
          if (!isNaN(num) && num > 0) parsedPeriod = num;
          parsedPeriodName = rawPeriodVal;
        }
      }

      // 시작/종료 시간 추출
      let parsedStartTime = null;
      let parsedEndTime = null;
      if (colStartTime > 0) {
        const rawST = String(row.getCell(colStartTime).value || '').trim();
        if (rawST && rawST !== '미입력' && rawST !== '미배정' && rawST !== '-') {
          parsedStartTime = this.formatTimeString(rawST);
        }
      }
      if (colEndTime > 0) {
        const rawET = String(row.getCell(colEndTime).value || '').trim();
        if (rawET && rawET !== '미입력' && rawET !== '미배정' && rawET !== '-') {
          parsedEndTime = this.formatTimeString(rawET);
        }
      }

      // 학년도 / 학기 추출
      let parsedYear = null;
      let parsedTerm = null;
      if (colYear > 0) {
        const rawY = String(row.getCell(colYear).value || '').replace(/[^0-9]/g, '');
        if (rawY.length === 4) parsedYear = rawY;
      }
      if (colSemester > 0) {
        const rawT = String(row.getCell(colSemester).value || '').replace(/[^0-9]/g, '');
        if (rawT === '1' || rawT === '2') parsedTerm = rawT;
      }

      // 비고 추출
      let parsedMemo = colMemo > 0 ? String(row.getCell(colMemo).value || '').trim() : '';
      if (parsedMemo === '-' || parsedMemo === '미입력') parsedMemo = '';

      const uniqueKey = `${studentCode}_${rawName}`;

      // 과목 컬럼들에서 과목명 추출
      const rowSubjects = [];
      colSubjects.forEach(colIdx => {
        const val = row.getCell(colIdx).value;
        if (!val) return;
        const text = String(val).trim();
        if (!text || text === '-' || text === '미배정' || text === '없음' || text === '미입력') return;

        // 쉼표, 슬래시, 플러스, 개행 등으로 복수 과목 구분
        const tokens = text.split(/[,/+\n\r]/);
        tokens.forEach(token => {
          let cleanSub = token.trim();
          if (!cleanSub) return;
          // 뒤쪽의 학점 표기 (예: (3), (3학점), [3학점]) 제거
          cleanSub = cleanSub.replace(/[\(\[]\s*\d+\s*(?:학점)?\s*[\)\]]$/, '').trim();
          if (cleanSub && cleanSub !== '-' && cleanSub !== '미입력' && !rowSubjects.includes(cleanSub)) {
            rowSubjects.push(cleanSub);
          }
        });
      });

      if (rowSubjects.length > 0) {
        rowSubjects.forEach(cleanSub => {
          const scheduleEntry = {
            subject: cleanSub,
            status: normStatus.scheduleStatus,
            courseStatus: normStatus.courseStatus,
            statusLabel: normStatus.label,
            day: parsedDay,
            period: parsedPeriod,
            periodName: parsedPeriodName,
            startTime: parsedStartTime,
            endTime: parsedEndTime,
            year: parsedYear,
            semester: parsedTerm,
            memo: parsedMemo
          };

          if (studentMap.has(uniqueKey)) {
            const existing = studentMap.get(uniqueKey);
            if (!existing.subjects.includes(cleanSub)) {
              existing.subjects.push(cleanSub);
            }
            if (!existing.schedules) existing.schedules = [];
            existing.schedules.push(scheduleEntry);
          } else {
            studentMap.set(uniqueKey, {
              id: `s_${Date.now()}_${r}_${Math.random().toString(36).substr(2, 4)}`,
              grade: grade,
              classNum: classNum,
              studentNum: studentNum,
              studentCode: studentCode,
              name: rawName,
              subject: cleanSub,
              subjects: [cleanSub],
              schedules: [scheduleEntry],
              day: parsedDay,
              period: parsedPeriod,
              periodId: null,
              status: normStatus.scheduleStatus,
              year: parsedYear,
              semester: parsedTerm,
              memo: parsedMemo
            });
          }
        });
      } else {
        if (!studentMap.has(uniqueKey)) {
          studentMap.set(uniqueKey, {
            id: `s_${Date.now()}_${r}_${Math.random().toString(36).substr(2, 4)}`,
            grade: grade,
            classNum: classNum,
            studentNum: studentNum,
            studentCode: studentCode,
            name: rawName,
            subject: '',
            subjects: [],
            schedules: [],
            day: null,
            period: null,
            periodId: null,
            status: 'confirmed',
            year: parsedYear,
            semester: parsedTerm,
            memo: parsedMemo
          });
        }
      }
    }

    const students = Array.from(studentMap.values());
    if (students.length === 0) {
      throw new Error('유효한 학생 데이터를 찾을 수 없습니다. [학년, 반, 번호, 성명] 또는 [학번, 성명, 과목] 열을 확인해주세요.');
    }
    return students;
  },

  // 5. 결괏값 형태 1: 학급 시간표 형태 (파스텔톤 디자인 멀티시트 엑셀)
  async exportTimetableExcel(periods, days, students, subjectColorMap, academicYear = '2026', semester = '1', schoolName = '강원온라인학교', courseEnrollments = null, masterStudents = null, courseStatus = null, yearData = null, noticeText = '') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;

    // 5-0. [1,2학기 통합시간표] 시트 생성 (엑셀의 가장 첫 번째 탭 위치)
    if (yearData && typeof yearData === 'object') {
      const unifiedSheet = workbook.addWorksheet('1,2학기 통합시간표', { views: [{ showGridLines: true, zoomScale: 40 }] });
      this.buildUnifiedSemesterTimetableSheet(unifiedSheet, periods, academicYear, schoolName, yearData, noticeText, subjectColorMap);
    }

    // 5-1. 학년-반별 개별 시트 자동 생성
    // ★ '1,2학기 통합시간표' 양식을 모두 따르되, 해당 학년/반 학생이 수강하는 과목만 표시 (함께 듣는 다른 반 학생도 함께 표기)
    const classSet = new Set();
    const collectClasses = (semData) => {
      if (!semData) return;
      if (Array.isArray(semData.masterStudents)) {
        semData.masterStudents.forEach(m => {
          if (m && m.grade && m.classNum) classSet.add(`${m.grade}-${m.classNum}`);
        });
      }
      if (Array.isArray(semData.students)) {
        semData.students.forEach(s => {
          if (s && s.grade && s.classNum) classSet.add(`${s.grade}-${s.classNum}`);
        });
      }
    };

    if (yearData && typeof yearData === 'object') {
      collectClasses(yearData['1']);
      collectClasses(yearData['2']);
    }
    if (students && Array.isArray(students)) {
      students.forEach(s => {
        if (s && s.grade && s.classNum) classSet.add(`${s.grade}-${s.classNum}`);
      });
    }

    const sortedClasses = Array.from(classSet).sort((a, b) => {
      const [gA, cA] = a.split('-').map(Number);
      const [gB, cB] = b.split('-').map(Number);
      return gA === gB ? cA - cB : gA - gB;
    });

    for (const cls of sortedClasses) {
      const [grade, classNum] = cls.split('-').map(Number);
      const sheetName = `${grade}학년 ${classNum}반`;
      const classSheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true, zoomScale: 40 }] });
      this.buildUnifiedSemesterTimetableSheet(
        classSheet,
        periods,
        academicYear,
        schoolName,
        yearData,
        noticeText,
        subjectColorMap,
        { grade, classNum }
      );
    }

    // 5-2. [전체 수강신청 내역] 시트 구성 (1,2학기 구분 없이 전체 대상)
    const rosterSheet = workbook.addWorksheet('전체 수강신청 내역', { views: [{ showGridLines: true }] });
    rosterSheet.columns = [
      { header: '학년도', key: 'academicYear', width: 14 },
      { header: '학기', key: 'semester', width: 11 },
      { header: '과목명', key: 'subject', width: 28 },
      { header: '학번', key: 'studentCode', width: 14 },
      { header: '학생 이름', key: 'studentName', width: 16 },
      { header: '개설/조정/폐강', key: 'statusText', width: 16 }
    ];

    // 자동 필터 적용 (A1부터 F1까지 드롭다운 필터 생성)
    rosterSheet.autoFilter = 'A1:F1';

    // 헤더 행 스타일
    const rHeaderRow = rosterSheet.getRow(1);
    rHeaderRow.height = 30;
    rHeaderRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E3A8A' } // 짙은 네이비
      };
      cell.font = { name: '맑은 고딕', size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const getStatusLabelForSem = (statusMap, sub) => {
      if (!statusMap) return '개설';
      const st = statusMap[sub] || 'open';
      if (st === 'closed' || st === '폐강') return '폐강';
      if (st === 'adjust' || st === '조정') return '조정';
      if (st === 'excluded' || st === '제외') return '폐강'; // '제외'는 폐강으로 표기
      return '개설';
    };

    const rosterRows = [];

    const extractRosterForSemester = (semData, semNum) => {
      if (!semData || typeof semData !== 'object') return;
      const semLabel = `${semNum}학기`;
      const statusMap = semData.courseStatus || {};
      const enrollments = semData.courseEnrollments || {};
      const masterStudents = semData.masterStudents || [];
      const masterMap = new Map();
      masterStudents.forEach(m => {
        if (!m) return;
        const code = m.studentCode || (m.grade && m.classNum && m.studentNum ? `${m.grade}${String(m.classNum).padStart(2,'0')}${String(m.studentNum).padStart(2,'0')}` : '');
        if (m.id) masterMap.set(m.id, { code, name: m.name });
        if (m.name) masterMap.set(m.name, { code, name: m.name });
      });

      let addedFromEnroll = false;
      if (enrollments && typeof enrollments === 'object') {
        Object.entries(enrollments).forEach(([subject, studentIds]) => {
          if (!Array.isArray(studentIds) || studentIds.length === 0) return;
          addedFromEnroll = true;
          const statusText = getStatusLabelForSem(statusMap, subject);
          studentIds.forEach(id => {
            const m = masterMap.get(id);
            const code = m ? m.code : '';
            const name = m ? m.name : String(id);
            rosterRows.push({
              academicYear: `${academicYear}학년도`,
              semester: semLabel,
              subject: subject,
              studentCode: code,
              studentName: name,
              statusText: statusText
            });
          });
        });
      }

      if (!addedFromEnroll && Array.isArray(semData.students)) {
        const seen = new Set();
        semData.students.forEach(s => {
          if (!s || !s.subject || !s.name) return;
          const code = s.studentCode || (s.grade && s.classNum && s.studentNum ? `${s.grade}${String(s.classNum).padStart(2,'0')}${String(s.studentNum).padStart(2,'0')}` : '');
          const key = `${semLabel}_${code}_${s.subject}`;
          if (!seen.has(key)) {
            seen.add(key);
            const statusText = getStatusLabelForSem(statusMap, s.subject);
            rosterRows.push({
              academicYear: `${academicYear}학년도`,
              semester: semLabel,
              subject: s.subject,
              studentCode: code,
              studentName: s.name,
              statusText: statusText
            });
          }
        });
      }
    };

    if (yearData && typeof yearData === 'object') {
      extractRosterForSemester(yearData['1'], 1);
      extractRosterForSemester(yearData['2'], 2);
    } else {
      // Fallback
      if (courseEnrollments && typeof courseEnrollments === 'object') {
        const studentMasterMap = new Map();
        if (masterStudents && Array.isArray(masterStudents)) {
          masterStudents.forEach(m => studentMasterMap.set(m.id, m));
        }
        Object.entries(courseEnrollments).forEach(([subject, studentIds]) => {
          if (!Array.isArray(studentIds)) return;
          const statusText = getStatusLabelForSem(courseStatus, subject);
          studentIds.forEach(id => {
            const m = studentMasterMap.get(id);
            const code = m ? (m.studentCode || this.getStudentCode(m)) : '';
            const name = m ? m.name : String(id);
            rosterRows.push({
              academicYear: `${academicYear}학년도`,
              semester: `${semester}학기`,
              subject: subject,
              studentCode: code,
              studentName: name,
              statusText: statusText
            });
          });
        });
      }
    }

    // ★ 정렬 우선순위: 1순위 상태(개설 -> 조정 -> 폐강), 2순위 학기(1학기 -> 2학기), 3순위 과목명(가나다), 4순위 학번(오름차순)
    const getStatusPriority = (st) => {
      if (st === '개설') return 1;
      if (st === '조정') return 2;
      if (st === '폐강') return 3;
      return 4;
    };

    rosterRows.sort((a, b) => {
      // 1순위: 상태 (개설 -> 조정 -> 폐강)
      const stComp = getStatusPriority(a.statusText) - getStatusPriority(b.statusText);
      if (stComp !== 0) return stComp;

      // 2순위: 학기 (1학기 -> 2학기)
      const semComp = (a.semester || '').localeCompare(b.semester || '', 'ko', { numeric: true });
      if (semComp !== 0) return semComp;

      // 3순위: 과목명 (가나다순)
      const subComp = (a.subject || '').localeCompare(b.subject || '', 'ko');
      if (subComp !== 0) return subComp;

      // 4순위: 학번 (오름차순)
      const codeComp = (a.studentCode || '').localeCompare(b.studentCode || '', 'ko', { numeric: true });
      if (codeComp !== 0) return codeComp;

      // 5순위: 이름 (가나다순)
      return (a.studentName || '').localeCompare(b.studentName || '', 'ko');
    });

    // 데이터 행 추가
    rosterRows.forEach((rData, idx) => {
      const row = rosterSheet.addRow({
        academicYear: rData.academicYear,
        semester: rData.semester,
        subject: rData.subject,
        studentCode: rData.studentCode,
        studentName: rData.studentName,
        statusText: rData.statusText
      });
      row.height = 24;
      const isEven = idx % 2 === 1;

      const isCanceled = rData.statusText === '폐강';
      const isOpen = rData.statusText === '개설';

      row.eachCell((cell, colNum) => {
        let cellColor = isCanceled ? 'FF94A3B8' : 'FF1E293B';
        let cellStrike = isCanceled;
        let cellBold = false;
        let fillBgColor = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

        if (isCanceled) {
          // 폐강인 경우: 데이터 행 전체 회색 채우기 + 취소선
          fillBgColor = 'FFE2E8F0';
        } else if (isOpen) {
          // 개설인 경우: 데이터 행 전체 노란색 채우기
          fillBgColor = 'FFFEF9C3'; // 부드럽고 가독성 높은 파스텔 옐로우
        }

        // 6열 상태 텍스트 색상
        if (colNum === 6) {
          cellBold = true;
          if (rData.statusText === '개설') cellColor = 'FF059669'; // 진한 초록
          else if (rData.statusText === '조정') cellColor = 'FFD97706'; // 진한 주황
          else if (rData.statusText === '폐강') cellColor = 'FFDC2626'; // 진한 빨강
        }

        cell.font = {
          name: '맑은 고딕',
          size: 10,
          bold: cellBold,
          strike: cellStrike,
          color: { argb: cellColor }
        };
        cell.alignment = { vertical: 'middle', horizontal: colNum === 3 ? 'left' : 'center' };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: fillBgColor }
        };
        this.applyBorder(cell);
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    // ★ 다운로드 시 파일명: 2027학년도_강원온라인학교_학급별_시간표_20261112_1745.xlsx
    const fileName = `${academicYear}학년도_${schoolName}_학급별_시간표_${this.getTimestamp()}.xlsx`;
    this.saveFile(buffer, fileName);
  },

  // 개별 시간표 시트 디자인 생성 헬퍼
                  // =========================================================================
  // 5-4. [신규] 1,2학기 통합시간표 시트 디자인 생성 헬퍼
  // 월화수목금 각 4개 열 (1학기 2개열: 과목명/(학번) 학생명, 2학기 2개열: 과목명/(학번) 학생명)
  // - [화면 배율] 파일을 열었을 때 전체가 한눈에 보이도록 40% 축소 상태(zoomScale: 40)로 저장
  // - [안내사항] '안내사항: '(빨간색) 및 사용자 입력 내용 모두 글자 크기 32pt 적용
  // - [행 높이] 1명 신청 행(단일 행)은 과목명-줄바꿈-개설상태가 모두 잘 보이도록 최소 44pt 적용
  // - [탭 위치] 엑셀 파일 내 가장 첫 번째 탭에 위치
  // - [타이틀] '2027학년도 강원온라인학교 1,2학기 통합시간표', 글자크기 26pt
  // - [교시 표기] 시간 '~' 다음 줄바꿈: 5교시\n(13:40~\n14:30)
  // - [과목명/상태] 줄바꿈 및 가운데 정렬: 과목명\n(개설상태) - 상태 글자크기 10pt
  // - [셀 서식] 월~금 열에 있는 모든 데이터: '셀에 맞춤(shrinkToFit: true)' 설정
  // - [폐강 과목] 과목명 취소선(strike: true) 적용, 과목 및 학생 셀 배경 연회색(#E2E8F0) 처리
  // - [인쇄 설정] 가로 B4 1장에 꽉 차게 인쇄 (fitToWidth: 1, fitToHeight: 1, simplex)
  // =========================================================================
  buildUnifiedSemesterTimetableSheet(sheet, periods, academicYear, schoolName, yearData, noticeText = '', subjectColorMap = {}, classFilter = null) {
    const sem1Data = (yearData && yearData['1']) || {};
    const sem2Data = (yearData && yearData['2']) || {};

    // -------------------------------------------------------------------------
    // ★ 화면 보기 설정: 파일을 열었을 때 전체가 한눈에 보이도록 70% 축소 배율 적용
    // -------------------------------------------------------------------------
    sheet.views = [
      {
        state: 'normal',
        showGridLines: true,
        zoomScale: 40 // 전체 21개 열과 시간표 전체가 한눈에 들어오는 40% 축소 배율
      }
    ];

    // -------------------------------------------------------------------------
    // ★ 페이지 인쇄 설정: B4 가로 1장에 꽉 차게 인쇄 배율 설정 (단면 인쇄)
    // -------------------------------------------------------------------------
    sheet.pageSetup = {
      orientation: 'landscape',  // 용지 방향: 가로
      paperSize: 12,            // 크기: B4 (JIS)
      duplex: 'simplex',        // 단면 인쇄
      margins: {                // 여백: 좁게 (Narrow)
        left: 0.25,
        right: 0.25,
        top: 0.4,
        bottom: 0.4,
        header: 0.2,
        footer: 0.2
      },
      fitToPage: true,          // 1페이지 꽉 차게 맞춤
      fitToWidth: 1,            // 가로 1페이지 폭에 맞춤
      fitToHeight: 1            // 세로 1페이지 높이에 맞춤 (전체 1장에 꽉 차게 인쇄!)
    };

    const targetDays = [
      { id: 'mon', name: '월요일' },
      { id: 'tue', name: '화요일' },
      { id: 'wed', name: '수요일' },
      { id: 'thu', name: '목요일' },
      { id: 'fri', name: '금요일' }
    ];

    // 연한 파스텔톤 팔레트 (진하지 않고 부드러운 톤)
    const softPastelPalette = [
      { bg: 'FFE0F2FE', text: 'FF0369A1' }, // 소프트 스카이블루
      { bg: 'FFD1FAE5', text: 'FF047857' }, // 소프트 민트그린
      { bg: 'FFEDE9FE', text: 'FF6D28D9' }, // 소프트 라벤더
      { bg: 'FFFFE4E6', text: 'FFBE123C' }, // 소프트 로즈핑크
      { bg: 'FFFEF3C7', text: 'FFB45309' }, // 소프트 앰버옐로우
      { bg: 'FFFFEDD5', text: 'FFC2410C' }, // 소프트 피치오렌지
      { bg: 'FFCCFBF1', text: 'FF0F766E' }, // 소프트 틸
      { bg: 'FFFCE7F3', text: 'FFBE185D' }, // 소프트 핑크
      { bg: 'FFE0E7FF', text: 'FF4338CA' }, // 소프트 인디고블루
      { bg: 'FFECFCCB', text: 'FF4D7C0F' }, // 소프트 라임그린
      { bg: 'FFCFFAFE', text: 'FF0E7490' }, // 소프트 시안
      { bg: 'FFF3E8FF', text: 'FF7E22CE' }, // 소프트 퍼플
      { bg: 'FFFEE2E2', text: 'FFB91C1C' }, // 소프트 코랄
      { bg: 'FFFEF9C3', text: 'FFA16207' }, // 소프트 버터옐로우
      { bg: 'FFE2E8F0', text: 'FF334155' }  // 소프트 슬레이트그레이
    ];

    // 과목명에 따른 파스텔 색상 반환 헬퍼
    const getSubjectColor = (subName) => {
      if (!subName) return null;
      if (subjectColorMap && subjectColorMap[subName]) {
        const info = subjectColorMap[subName];
        return {
          bg: info.excelBg ? 'FF' + info.excelBg : softPastelPalette[0].bg,
          text: info.excelText ? 'FF' + info.excelText : 'FF1E293B'
        };
      }
      let hash = 0;
      for (let i = 0; i < subName.length; i++) {
        hash = (hash << 5) - hash + subName.charCodeAt(i);
        hash |= 0;
      }
      const idx = Math.abs(hash) % softPastelPalette.length;
      return softPastelPalette[idx];
    };

    // 과목 상태 텍스트 판정 헬퍼 ('제외'는 '폐강'으로 표기)
    const getStatusText = (statusMap, subName) => {
      if (!statusMap) return '개설';
      const st = statusMap[subName] || 'open';
      if (st === 'closed' || st === '폐강') return '폐강';
      if (st === 'adjust' || st === '조정') return '조정';
      if (st === 'excluded' || st === '제외') return '폐강';
      return '개설';
    };

    // 열 너비 설정 (A열: 교시, 각 요일당 4개 열씩 총 21개 열)
    const columnsConfig = [
      { width: 18 } // A: 교시 (16pt 글씨 최적화)
    ];
    targetDays.forEach(() => {
      columnsConfig.push({ width: 23 }); // 1학기 과목명
      columnsConfig.push({ width: 20 }); // 1학기 (학번) 학생명
      columnsConfig.push({ width: 23 }); // 2학기 과목명
      columnsConfig.push({ width: 20 }); // 2학기 (학번) 학생명
    });
    sheet.columns = columnsConfig;

    // 1행: 타이틀 -> [2027학년도 강원온라인학교 1,2학기 통합시간표] 또는 [2027학년도 강원온라인학교 [2학년 1반] 1,2학기 시간표]
    sheet.mergeCells(1, 1, 1, 21);
    const titleCell = sheet.getCell(1, 1);
    if (classFilter) {
      titleCell.value = `${academicYear}학년도 ${schoolName} [${classFilter.grade}학년 ${classFilter.classNum}반] 1,2학기 시간표`;
    } else {
      titleCell.value = `${academicYear}학년도 ${schoolName} 1,2학기 통합시간표`;
    }
    titleCell.font = { name: '맑은 고딕', size: 26, bold: true, color: { argb: 'FF1E293B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(1).height = 54;

    let rowDayIdx = 2;
    // 2행: 안내사항 -> '안내사항: '(빨간색) 및 사용자가 입력한 내용 모두 동일하게 32pt 적용!
    const trimmedNotice = (noticeText || '').trim();
    if (trimmedNotice) {
      sheet.mergeCells(2, 1, 2, 21);
      const noticeCell = sheet.getCell(2, 1);
      const inlineNotice = trimmedNotice.replace(/\r?\n/g, '   ');

      noticeCell.value = {
        richText: [
          {
            text: '안내사항: ',
            font: { name: '맑은 고딕', size: 32, bold: true, color: { argb: 'FFDC2626' } } // 빨간색 32pt
          },
          {
            text: inlineNotice,
            font: { name: '맑은 고딕', size: 32, bold: true, color: { argb: 'FF1E293B' } } // 입력 내용도 동일하게 32pt!
          }
        ]
      };
      noticeCell.alignment = { vertical: 'middle', horizontal: 'left', shrinkToFit: true };
      sheet.getRow(2).height = 68; // 32pt 폰트에 맞춘 넉넉하고 시원한 행 높이

      for (let c = 1; c <= 21; c++) {
        const cell = sheet.getCell(2, c);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFBEB' } // 부드럽고 따뜻한 크림/아이보리 옐로우
        };
        cell.border = {
          top: { style: 'medium', color: { argb: 'FFD97706' } },
          bottom: { style: 'medium', color: { argb: 'FFD97706' } },
          left: { style: 'medium', color: { argb: 'FFD97706' } },
          right: { style: 'medium', color: { argb: 'FFD97706' } }
        };
      }
      rowDayIdx = 3;
    }

    const rowSemIdx = rowDayIdx + 1;
    const rowSubjIdx = rowSemIdx + 1;
    let currentRow = rowSubjIdx + 1;

    // 요일 대분류 헤더 (A열 교시, 각 요일 4개열씩 병합) - 글자 크기 16
    const rowDay = sheet.getRow(rowDayIdx);
    rowDay.height = 36;
    targetDays.forEach((day, idx) => {
      const startCol = 2 + idx * 4;
      const endCol = startCol + 3;
      sheet.mergeCells(rowDayIdx, startCol, rowDayIdx, endCol);
      const dayCell = sheet.getCell(rowDayIdx, startCol);
      dayCell.value = day.name;
      dayCell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      dayCell.alignment = { vertical: 'middle', horizontal: 'center', shrinkToFit: true };
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(rowDayIdx, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        this.applyBorder(cell);
      }
    });

    // 학기 중분류 헤더 (1학기 2개열, 2학기 2개열 병합) - 글자 크기 16
    const rowSem = sheet.getRow(rowSemIdx);
    rowSem.height = 34;
    targetDays.forEach((day, idx) => {
      const colSem1 = 2 + idx * 4;
      const colSem2 = colSem1 + 2;

      // 1학기 헤더 (2개 열 병합)
      sheet.mergeCells(rowSemIdx, colSem1, rowSemIdx, colSem1 + 1);
      const sem1Cell = sheet.getCell(rowSemIdx, colSem1);
      sem1Cell.value = '1학기';
      sem1Cell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF1D4ED8' } };
      sem1Cell.alignment = { vertical: 'middle', horizontal: 'center', shrinkToFit: true };
      for (let c = colSem1; c <= colSem1 + 1; c++) {
        const cell = sheet.getCell(rowSemIdx, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
        this.applyBorder(cell);
      }

      // 2학기 헤더 (2개 열 병합)
      sheet.mergeCells(rowSemIdx, colSem2, rowSemIdx, colSem2 + 1);
      const sem2Cell = sheet.getCell(rowSemIdx, colSem2);
      sem2Cell.value = '2학기';
      sem2Cell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF4338CA' } };
      sem2Cell.alignment = { vertical: 'middle', horizontal: 'center', shrinkToFit: true };
      for (let c = colSem2; c <= colSem2 + 1; c++) {
        const cell = sheet.getCell(rowSemIdx, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF2FF' } };
        this.applyBorder(cell);
      }
    });

    // 소분류 헤더 (과목명, (학번) 학생명) - 글자 크기 16
    const rowSubj = sheet.getRow(rowSubjIdx);
    rowSubj.height = 34;
    targetDays.forEach((day, idx) => {
      const colBase = 2 + idx * 4;
      const headers = ['과목명', '(학번) 학생명', '과목명', '(학번) 학생명'];
      headers.forEach((h, hIdx) => {
        const cell = sheet.getCell(rowSubjIdx, colBase + hIdx);
        cell.value = h;
        cell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF334155' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', shrinkToFit: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        this.applyBorder(cell);
      });
    });

    // A열 교시 헤더 병합 (rowDayIdx ~ rowSubjIdx)
    sheet.mergeCells(rowDayIdx, 1, rowSubjIdx, 1);
    const colPeriodCell = sheet.getCell(rowDayIdx, 1);
    colPeriodCell.value = '교시';
    colPeriodCell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    colPeriodCell.alignment = { vertical: 'middle', horizontal: 'center' };
    for (let r = rowDayIdx; r <= rowSubjIdx; r++) {
      const cell = sheet.getCell(r, 1);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      this.applyBorder(cell);
    }

    const statusMap1 = sem1Data.courseStatus || {};
    const statusMap2 = sem2Data.courseStatus || {};

    // 특정 학기 데이터(semData)에서 특정 요일/교시 슬롯의 과목 및 학생 목록 추출 헬퍼
    const getSlotItems = (semData, dayId, period) => {
      if (!semData || typeof semData !== 'object') return [];

      const periodIdStr = String(period.id);
      const periodNum = period.period;

      // 1) semData.students 배열(시간표에 매핑된 최종 평탄화 데이터) 최우선 활용
      if (Array.isArray(semData.students) && semData.students.length > 0) {
        const matching = semData.students.filter(s => {
          if (!s || !s.subject) return false;
          if (s.day !== dayId) return false;
          const matchPeriodId = s.periodId && String(s.periodId) === periodIdStr;
          const matchPeriodNum = s.period !== undefined && s.period !== null && s.period === periodNum;
          return matchPeriodId || matchPeriodNum;
        });

        if (matching.length > 0) {
          const subjectMap = new Map();
          matching.forEach(s => {
            if (!subjectMap.has(s.subject)) {
              subjectMap.set(s.subject, { memo: s.memo || '', list: [] });
            }
            if (s.name) {
              const code = s.studentCode || (s.grade && s.classNum && s.studentNum ? `${s.grade}${String(s.classNum).padStart(2,'0')}${String(s.studentNum).padStart(2,'0')}` : '');
              const grade = s.grade || (code ? code.charAt(0) : '');
              const classNum = s.classNum !== undefined && s.classNum !== null ? s.classNum : (code && code.length >= 3 ? parseInt(code.substring(1, 3), 10) : '');
              subjectMap.get(s.subject).list.push({
                stId: code,
                stName: s.name,
                grade: grade,
                classNum: classNum
              });
            }
          });

          const items = [];
          const sortedSubjects = Array.from(subjectMap.keys()).sort((a, b) => a.localeCompare(b, 'ko'));
          sortedSubjects.forEach(sub => {
            const entry = subjectMap.get(sub);
            const memo = entry.memo || '';
            const stList = entry.list;

            // ★ classFilter가 있는 경우: 해당 학급 학생이 1명이라도 듣는 과목만 선별
            if (classFilter) {
              const hasTargetClass = stList.some(st => {
                const matchG = String(st.grade) === String(classFilter.grade);
                const matchC = Number(st.classNum) === Number(classFilter.classNum);
                return matchG && matchC;
              });
              if (!hasTargetClass) return; // 이 학급 학생이 안 듣는 과목은 미표기
            }

            stList.sort((a, b) => {
              if (a.stId && b.stId) return a.stId.localeCompare(b.stId, 'ko', { numeric: true });
              return a.stName.localeCompare(b.stName, 'ko');
            });

            if (stList.length === 0) {
              items.push({ subject: sub, stId: '', stName: '-', memo, grade: '', classNum: '' });
            } else {
              stList.forEach(st => {
                items.push({ subject: sub, stId: st.stId, stName: st.stName, memo, grade: st.grade, classNum: st.classNum });
              });
            }
          });
          return items;
        }
      }

      // 2) semData.courseSchedules 및 courseEnrollments 에서 직접 추출
      const schedules = semData.courseSchedules || {};
      const enrollments = semData.courseEnrollments || {};
      const masterStudents = semData.masterStudents || [];

      const masterIdMap = new Map();
      const masterNameMap = new Map();
      masterStudents.forEach(m => {
        if (!m) return;
        const code = m.studentCode || (m.grade && m.classNum && m.studentNum ? `${m.grade}${String(m.classNum).padStart(2,'0')}${String(m.studentNum).padStart(2,'0')}` : '');
        const grade = m.grade || (code ? code.charAt(0) : '');
        const classNum = m.classNum !== undefined && m.classNum !== null ? m.classNum : (code && code.length >= 3 ? parseInt(code.substring(1, 3), 10) : '');
        const info = { id: m.id, code, name: m.name, grade, classNum };
        if (m.id) masterIdMap.set(m.id, info);
        if (m.name) masterNameMap.set(m.name, info);
      });

      const matchedSubjects = [];
      Object.entries(schedules).forEach(([subject, scheduleSlots]) => {
        if (!scheduleSlots) return;
        const slotsArray = Array.isArray(scheduleSlots) ? scheduleSlots : [scheduleSlots];
        const matchedSlot = slotsArray.find(slot => {
          if (!slot || slot.day !== dayId) return false;
          const matchId = slot.periodId && String(slot.periodId) === periodIdStr;
          const matchNum = slot.period !== undefined && slot.period !== null && slot.period === periodNum;
          return matchId || matchNum;
        });

        if (matchedSlot) {
          matchedSubjects.push({ subject, memo: matchedSlot.memo || '' });
        }
      });

      matchedSubjects.sort((a, b) => a.subject.localeCompare(b.subject, 'ko'));

      const items = [];
      matchedSubjects.forEach(matchObj => {
        const subject = matchObj.subject;
        const memo = matchObj.memo || '';
        const enrolled = enrollments[subject] || [];
        const studentItems = [];

        enrolled.forEach(stRef => {
          const m = masterIdMap.get(stRef) || masterNameMap.get(stRef);
          if (m) {
            studentItems.push({ stId: m.code || '', stName: m.name, grade: m.grade, classNum: m.classNum });
          } else {
            studentItems.push({ stId: '', stName: String(stRef), grade: '', classNum: '' });
          }
        });

        // ★ classFilter가 있는 경우: 해당 학급 학생이 1명이라도 듣는 과목만 선별
        if (classFilter) {
          const hasTargetClass = studentItems.some(st => {
            const matchG = String(st.grade) === String(classFilter.grade);
            const matchC = Number(st.classNum) === Number(classFilter.classNum);
            return matchG && matchC;
          });
          if (!hasTargetClass) return; // 이 학급 학생이 안 듣는 과목은 미표기
        }

        studentItems.sort((a, b) => {
          if (a.stId && b.stId) return a.stId.localeCompare(b.stId, 'ko', { numeric: true });
          return a.stName.localeCompare(b.stName, 'ko');
        });

        if (studentItems.length === 0) {
          items.push({ subject, stId: '', stName: '-', memo, grade: '', classNum: '' });
        } else {
          studentItems.forEach(st => {
            items.push({ subject, stId: st.stId, stName: st.stName, memo, grade: st.grade, classNum: st.classNum });
          });
        }
      });

      return items;
    };

    // 교시별 행 생성
    const periodStartRows = new Set();
    const periodEndRows = new Set();

    (periods || []).forEach(period => {
      const daySlotsData = targetDays.map(day => {
        const items1 = getSlotItems(sem1Data, day.id, period);
        const items2 = getSlotItems(sem2Data, day.id, period);
        return { day, items1, items2 };
      });

      let maxBlockRows = 1;
      daySlotsData.forEach(d => {
        const count = Math.max(d.items1.length, d.items2.length);
        if (count > maxBlockRows) maxBlockRows = count;
      });

      const blockStartRow = currentRow;
      const blockEndRow = blockStartRow + maxBlockRows - 1;

      // ★ 방안 2 적용: 1명 신청 과목이 다른 과목과 같은 교시에 묶여 1행만 차지하더라도 텍스트(과목명+상태)가 잘리지 않도록
      // 2명 이상 블록의 기본 행당 높이를 기존 29pt -> 38pt로 대폭 상향! (1명 과목도 최소 38pt 확보)
      const dynamicRowHeight = maxBlockRows === 1 ? 44 : 38;
      for (let r = blockStartRow; r <= blockEndRow; r++) {
        sheet.getRow(r).height = dynamicRowHeight;
      }

      // 교시 (A열) 세로 병합 및 시간 표기 ('~' 다음 줄바꿈)
      if (blockEndRow > blockStartRow) {
        sheet.mergeCells(blockStartRow, 1, blockEndRow, 1);
      }
      const periodCell = sheet.getCell(blockStartRow, 1);
      periodCell.value = `${period.name}\n(${period.startTime}~\n${period.endTime})`;
      periodCell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF1E293B' } };
      periodCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, shrinkToFit: true };

      for (let r = blockStartRow; r <= blockEndRow; r++) {
        const cell = sheet.getCell(r, 1);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        this.applyBorder(cell);
      }

      // 각 요일별 열 채우기 (월~금 열의 모든 데이터: '셀에 맞춤(shrinkToFit: true)')
      daySlotsData.forEach((d, dayIdx) => {
        const colSub1 = 2 + dayIdx * 4;
        const colSt1 = colSub1 + 1;
        const colSub2 = colSub1 + 2;
        const colSt2 = colSub2 + 1;

        // ==========================================
        // 1학기: 1) 데이터 기본 입력 및 (학번) 학생명 서식 적용
        // ==========================================
        for (let i = 0; i < maxBlockRows; i++) {
          const r = blockStartRow + i;
          const item = d.items1[i];
          const cSub = sheet.getCell(r, colSub1);
          const cSt = sheet.getCell(r, colSt1);

          cSub.value = item ? item.subject : '';

          const isCanceled1 = item && item.subject && getStatusText(statusMap1, item.subject) === '폐강';

          if (item) {
            if (item.stId) {
              cSt.value = {
                richText: [
                  { 
                    text: `(${item.stId}) `, 
                    font: { 
                      name: '맑은 고딕', 
                      size: 9, 
                      strike: isCanceled1,
                      color: { argb: isCanceled1 ? 'FF94A3B8' : 'FF475569' } 
                    } 
                  },
                  { 
                    text: item.stName, 
                    font: { 
                      name: '맑은 고딕', 
                      size: 13, 
                      strike: isCanceled1,
                      color: { argb: isCanceled1 ? 'FF94A3B8' : 'FF1E293B' } 
                    } 
                  }
                ]
              };
            } else {
              cSt.value = item.stName;
              cSt.font = { 
                name: '맑은 고딕', 
                size: 13, 
                strike: isCanceled1,
                color: { argb: isCanceled1 ? 'FF94A3B8' : 'FF1E293B' } 
              };
            }
          } else {
            cSt.value = '';
            cSt.font = { name: '맑은 고딕', size: 13 };
          }

          // (학번) 학생명: 왼쪽 정렬 + 셀에 맞춤(shrinkToFit: true)
          cSt.alignment = { vertical: 'middle', horizontal: 'left', shrinkToFit: true };

          // 폐강 과목인 경우 학생 셀 배경 회색 처리 / 해당 학급 학생은 노란색 강조 채우기
          if (isCanceled1) {
            cSt.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE2E8F0' }
            };
          } else if (classFilter && item) {
            const isTargetStudent = (
              (String(item.grade) === String(classFilter.grade) && Number(item.classNum) === Number(classFilter.classNum)) ||
              (item.stId && item.stId.startsWith(`${classFilter.grade}${String(classFilter.classNum).padStart(2, '0')}`))
            );
            if (isTargetStudent) {
              cSt.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFEF08A' } // 밝고 화사한 파스텔 옐로우 (해당 학급 학생 강조)
              };
            }
          }

          this.applyBorder(cSub);
          this.applyBorder(cSt);
        }

        // 1학기: 2) 동일 과목명 세로 병합을 먼저 수행
        const subSpans1 = [];
        if (d.items1.length > 0) {
          let spanStart = 0;
          while (spanStart < d.items1.length) {
            let spanEnd = spanStart;
            while (spanEnd + 1 < d.items1.length && d.items1[spanEnd + 1].subject === d.items1[spanStart].subject) {
              spanEnd++;
            }
            subSpans1.push({ start: spanStart, end: spanEnd, subject: d.items1[spanStart].subject, memo: d.items1[spanStart].memo || '' });
            if (spanEnd > spanStart) {
              sheet.mergeCells(blockStartRow + spanStart, colSub1, blockStartRow + spanEnd, colSub1);
            }
            spanStart = spanEnd + 1;
          }
        }

        // 1학기: 3) 병합 완료 후 -> 과목명 13pt, 줄바꿈, (개설상태) 10pt 가운데 정렬 + 셀에 맞춤
        subSpans1.forEach(span => {
          const topRow = blockStartRow + span.start;
          const endRow = blockStartRow + span.end;
          const cell = sheet.getCell(topRow, colSub1);
          const color = getSubjectColor(span.subject);
          const isCanceled = span.subject && getStatusText(statusMap1, span.subject) === '폐강';

          if (span.subject) {
            const stText = getStatusText(statusMap1, span.subject);
            let stColor = 'FF059669'; // 개설: 진한 초록
            if (stText === '조정') stColor = 'FFD97706'; // 조정: 진한 주황
            else if (stText === '폐강') stColor = 'FFDC2626'; // 폐강: 진한 빨강

            cell.value = {
              richText: [
                {
                  text: span.subject,
                  font: {
                    name: '맑은 고딕',
                    size: 13,
                    bold: true,
                    strike: isCanceled,
                    color: { argb: isCanceled ? 'FF64748B' : (color ? color.text : 'FF1E293B') }
                  }
                },
                {
                  text: `\n(1학기/${stText}${span.memo ? '/' + span.memo : ''})`,
                  font: {
                    name: '맑은 고딕',
                    size: 10,
                    bold: true,
                    color: { argb: stColor }
                  }
                }
              ]
            };
          } else {
            cell.value = '';
          }

          // 가운데 정렬, 줄바꿈 허용, 셀에 맞춤 설정
          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, shrinkToFit: true };

          // 배경색 채우기 (폐강 과목은 회색 #E2E8F0, 개설/조정은 고유 파스텔톤)
          const fillBgColor = isCanceled ? 'FFE2E8F0' : (color ? color.bg : null);
          if (fillBgColor) {
            for (let r = topRow; r <= endRow; r++) {
              const c = sheet.getCell(r, colSub1);
              c.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: fillBgColor }
              };
              c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, shrinkToFit: true };
              this.applyBorder(c);
            }
          } else {
            this.applyBorder(cell);
          }

          if (isCanceled) {
            for (let r = topRow; r <= endRow; r++) {
              const cSt = sheet.getCell(r, colSt1);
              cSt.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE2E8F0' }
              };
              this.applyBorder(cSt);
            }
          }
        });

        // ==========================================
        // 2학기: 1) 데이터 기본 입력 및 (학번) 학생명 서식 적용
        // ==========================================
        for (let i = 0; i < maxBlockRows; i++) {
          const r = blockStartRow + i;
          const item = d.items2[i];
          const cSub = sheet.getCell(r, colSub2);
          const cSt = sheet.getCell(r, colSt2);

          cSub.value = item ? item.subject : '';

          const isCanceled2 = item && item.subject && getStatusText(statusMap2, item.subject) === '폐강';

          if (item) {
            if (item.stId) {
              cSt.value = {
                richText: [
                  { 
                    text: `(${item.stId}) `, 
                    font: { 
                      name: '맑은 고딕', 
                      size: 9, 
                      strike: isCanceled2,
                      color: { argb: isCanceled2 ? 'FF94A3B8' : 'FF475569' } 
                    } 
                  },
                  { 
                    text: item.stName, 
                    font: { 
                      name: '맑은 고딕', 
                      size: 13, 
                      strike: isCanceled2,
                      color: { argb: isCanceled2 ? 'FF94A3B8' : 'FF1E293B' } 
                    } 
                  }
                ]
              };
            } else {
              cSt.value = item.stName;
              cSt.font = { 
                name: '맑은 고딕', 
                size: 13, 
                strike: isCanceled2,
                color: { argb: isCanceled2 ? 'FF94A3B8' : 'FF1E293B' } 
              };
            }
          } else {
            cSt.value = '';
            cSt.font = { name: '맑은 고딕', size: 13 };
          }

          // (학번) 학생명: 왼쪽 정렬 + 셀에 맞춤(shrinkToFit: true)
          cSt.alignment = { vertical: 'middle', horizontal: 'left', shrinkToFit: true };

          // 폐강 과목인 경우 학생 셀 배경 회색 처리 / 해당 학급 학생은 노란색 강조 채우기
          if (isCanceled2) {
            cSt.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE2E8F0' }
            };
          } else if (classFilter && item) {
            const isTargetStudent = (
              (String(item.grade) === String(classFilter.grade) && Number(item.classNum) === Number(classFilter.classNum)) ||
              (item.stId && item.stId.startsWith(`${classFilter.grade}${String(classFilter.classNum).padStart(2, '0')}`))
            );
            if (isTargetStudent) {
              cSt.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFEF08A' } // 밝고 화사한 파스텔 옐로우 (해당 학급 학생 강조)
              };
            }
          }

          this.applyBorder(cSub);
          this.applyBorder(cSt);
        }

        // 2학기: 2) 동일 과목명 세로 병합을 먼저 수행
        const subSpans2 = [];
        if (d.items2.length > 0) {
          let spanStart = 0;
          while (spanStart < d.items2.length) {
            let spanEnd = spanStart;
            while (spanEnd + 1 < d.items2.length && d.items2[spanEnd + 1].subject === d.items2[spanStart].subject) {
              spanEnd++;
            }
            subSpans2.push({ start: spanStart, end: spanEnd, subject: d.items2[spanStart].subject, memo: d.items2[spanStart].memo || '' });
            if (spanEnd > spanStart) {
              sheet.mergeCells(blockStartRow + spanStart, colSub2, blockStartRow + spanEnd, colSub2);
            }
            spanStart = spanEnd + 1;
          }
        }

        // 2학기: 3) 병합 완료 후 -> 과목명 13pt, 줄바꿈, (개설상태) 10pt 가운데 정렬 + 셀에 맞춤
        subSpans2.forEach(span => {
          const topRow = blockStartRow + span.start;
          const endRow = blockStartRow + span.end;
          const cell = sheet.getCell(topRow, colSub2);
          const color = getSubjectColor(span.subject);
          const isCanceled = span.subject && getStatusText(statusMap2, span.subject) === '폐강';

          if (span.subject) {
            const stText = getStatusText(statusMap2, span.subject);
            let stColor = 'FF059669'; // 개설: 진한 초록
            if (stText === '조정') stColor = 'FFD97706'; // 조정: 진한 주황
            else if (stText === '폐강') stColor = 'FFDC2626'; // 폐강: 진한 빨강

            cell.value = {
              richText: [
                {
                  text: span.subject,
                  font: {
                    name: '맑은 고딕',
                    size: 13,
                    bold: true,
                    strike: isCanceled,
                    color: { argb: isCanceled ? 'FF64748B' : (color ? color.text : 'FF1E293B') }
                  }
                },
                {
                  text: `\n(2학기/${stText}${span.memo ? '/' + span.memo : ''})`,
                  font: {
                    name: '맑은 고딕',
                    size: 10,
                    bold: true,
                    color: { argb: stColor }
                  }
                }
              ]
            };
          } else {
            cell.value = '';
          }

          // 가운데 정렬, 줄바꿈 허용, 셀에 맞춤 설정
          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, shrinkToFit: true };

          // 배경색 채우기 (폐강 과목은 회색 #E2E8F0, 개설/조정은 고유 파스텔톤)
          const fillBgColor = isCanceled ? 'FFE2E8F0' : (color ? color.bg : null);
          if (fillBgColor) {
            for (let r = topRow; r <= endRow; r++) {
              const c = sheet.getCell(r, colSub2);
              c.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: fillBgColor }
              };
              c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, shrinkToFit: true };
              this.applyBorder(c);
            }
          } else {
            this.applyBorder(cell);
          }

          if (isCanceled) {
            for (let r = topRow; r <= endRow; r++) {
              const cSt = sheet.getCell(r, colSt2);
              cSt.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE2E8F0' }
              };
              this.applyBorder(cSt);
            }
          }
        });
      });

      periodStartRows.add(blockStartRow);
      periodEndRows.add(blockEndRow);

      currentRow = blockEndRow + 1;
    });

    // ==========================================
    // ★ 요일 구분 테두리 스타일 일괄 적용
    // - 요일과 요일 사이: 굵은 테두리 (medium)
    // - 학기 및 일반 셀: 이전 기본 thin 테두리 유지
    // - 표 상단/하단 경계: 굵은 테두리 (medium)
    // ==========================================
    const endDataRow = currentRow - 1;
    const borderDark = { argb: 'FF334155' };
    const borderLight = { argb: 'FFCBD5E1' };

    for (let r = 3; r <= endDataRow; r++) {
      for (let c = 1; c <= 21; c++) {
        const cell = sheet.getCell(r, c);
        const curBorder = cell.border || {};

        let leftStyle = curBorder.left?.style || 'thin';
        let leftColor = curBorder.left?.color || borderLight;
        let rightStyle = curBorder.right?.style || 'thin';
        let rightColor = curBorder.right?.color || borderLight;
        let topStyle = curBorder.top?.style || 'thin';
        let topColor = curBorder.top?.color || borderLight;
        let bottomStyle = curBorder.bottom?.style || 'thin';
        let bottomColor = curBorder.bottom?.color || borderLight;

        // 1) 좌/우 경계 설정
        const borderDotted = { argb: 'FF64748B' };

        if (c === 1) {
          // 교시 열: 왼쪽 굵은 선, 오른쪽 굵은 선 (교시와 월요일 사이)
          leftStyle = 'medium';
          leftColor = borderDark;
          rightStyle = 'medium';
          rightColor = borderDark;
        } else {
          // c: 2 ~ 21 (월~금, 각 4열씩: 과목1, 학생1, 과목2, 학생2)
          const subCol = (c - 2) % 4; // 0: 1학기 과목, 1: 1학기 학생, 2: 2학기 과목, 3: 2학기 학생

          if (subCol === 0) {
            // 요일 시작 (1학기 과목) -> 왼쪽 굵은 선 (요일 구분)
            leftStyle = 'medium';
            leftColor = borderDark;
          } else if (subCol === 1) {
            // ★ 학기 간 구분: 1학기 끝 (1학기 학생) -> 오른쪽 점선 테두리
            rightStyle = 'dotted';
            rightColor = borderDotted;
          } else if (subCol === 2) {
            // ★ 학기 간 구분: 2학기 시작 (2학기 과목) -> 왼쪽 점선 테두리
            leftStyle = 'dotted';
            leftColor = borderDotted;
          } else if (subCol === 3) {
            // 요일 끝 (2학기 학생) -> 오른쪽 굵은 선 (요일 구분)
            rightStyle = 'medium';
            rightColor = borderDark;
          }
        }

        // 2) 상/하 경계 설정
        // ★ 교시 간 구분 테두리: 선 스타일 첫 번째의 일반선(thin) 테두리 (굵은 테두리 아님)
        const borderBlack = { argb: 'FF000000' };

        if (r === 3) {
          // 요일 헤더 맨 위
          topStyle = 'medium';
          topColor = borderDark;
        }
        if (r === 5) {
          // 헤더와 데이터 경계선 (1교시 상단: 일반선 검은색)
          bottomStyle = 'thin';
          bottomColor = borderBlack;
        }
        if (r === 6) {
          // 데이터 첫 행 상단 (일반선 검은색)
          topStyle = 'thin';
          topColor = borderBlack;
        }
        if (periodEndRows.has(r)) {
          // ★ 교시 간 하단 구분선: 일반선(thin) 검은색 테두리 (굵은 테두리 아님)
          bottomStyle = 'thin';
          bottomColor = borderBlack;
        }
        if (periodStartRows.has(r) && r > 6) {
          // ★ 교시 간 상단 구분선: 일반선(thin) 검은색 테두리 (굵은 테두리 아님)
          topStyle = 'thin';
          topColor = borderBlack;
        }
        if (r === endDataRow) {
          // 표 맨 아래 외곽선
          bottomStyle = 'medium';
          bottomColor = borderDark;
        }

        cell.border = {
          top: { style: topStyle, color: topColor },
          bottom: { style: bottomStyle, color: bottomColor },
          left: { style: leftStyle, color: leftColor },
          right: { style: rightStyle, color: rightColor }
        };
      }
    }
  },

  buildTimetableSheet(sheet, periods, days, studentList, subjectColorMap, titleText) {
    sheet.columns = [
      { width: 22 },
      ...days.map(() => ({ width: 28 }))
    ];

    // 타이틀 행
    sheet.mergeCells(1, 1, 1, days.length + 1);
    const titleCell = sheet.getCell(1, 1);
    titleCell.value = titleText;
    titleCell.font = { name: '맑은 고딕', size: 15, bold: true, color: { argb: 'FF1E293B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(1).height = 40;

    // 헤더 행 (요일)
    const headerRow = sheet.getRow(2);
    headerRow.height = 30;
    const firstHeader = headerRow.getCell(1);
    firstHeader.value = '교시 / 시간';
    firstHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    firstHeader.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF334155' } };
    firstHeader.alignment = { vertical: 'middle', horizontal: 'center' };
    this.applyBorder(firstHeader);

    days.forEach((day, idx) => {
      const cell = headerRow.getCell(idx + 2);
      cell.value = day.isWeekend ? `${day.name} (주말)` : `${day.name}요일`;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: day.isWeekend ? 'FFFEE2E2' : 'FFDBEAFE' }
      };
      cell.font = {
        name: '맑은 고딕',
        bold: true,
        color: { argb: day.isWeekend ? 'FF991B1B' : 'FF1E40AF' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    let currentExcelRow = 3;

    periods.forEach((period) => {
      const row = sheet.getRow(currentExcelRow);
      row.height = period.isBreak ? 28 : 55;

      // 1~4교시 및 점심시간은 기본 숨기기(hidden: true) 처리
      const pNum = Number(period.period);
      if ((pNum >= 1 && pNum <= 4) || period.id === 'p_lunch' || (period.name && (period.name.includes('1교시') || period.name.includes('2교시') || period.name.includes('3교시') || period.name.includes('4교시') || period.name.includes('점심')))) {
        row.hidden = true;
      }

      const periodCell = row.getCell(1);
      periodCell.value = period.isBreak
        ? `[${period.name}]\n${period.startTime}~${period.endTime}`
        : `${period.name}\n(${period.startTime}~${period.endTime})`;
      periodCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: period.isBreak ? 'FFF1F5F9' : 'FFF8FAFC' }
      };
      periodCell.font = {
        name: '맑은 고딕',
        bold: true,
        size: period.isBreak ? 9 : 10,
        color: { argb: period.isBreak ? 'FF64748B' : 'FF334155' }
      };
      periodCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      this.applyBorder(periodCell);

      if (period.isBreak) {
        sheet.mergeCells(currentExcelRow, 2, currentExcelRow, days.length + 1);
        const breakCell = row.getCell(2);
        breakCell.value = `☕ ${period.name} (${period.startTime} ~ ${period.endTime})`;
        breakCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        breakCell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF94A3B8' }, italic: true };
        breakCell.alignment = { vertical: 'middle', horizontal: 'center' };
        for (let c = 2; c <= days.length + 1; c++) {
          this.applyBorder(row.getCell(c));
        }
      } else {
        days.forEach((day, dayIdx) => {
          const cell = row.getCell(dayIdx + 2);
          
          const assigned = studentList.filter(s => {
            if (s.day !== day.id) return false;
            if (s.periodId) return s.periodId === period.id;
            return s.period === period.period;
          });

          if (assigned.length > 0) {
            // 셀 배경은 흰색으로 통일하여 특정 과목 하나의 색으로 전체가 덮이지 않도록 함
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFFFF' }
            };

            // 셀 내 각 과목별 텍스트 음영(폰트 컬러) 처리 (ExcelJS RichText 활용)
            const richTextItems = [];

            assigned.forEach((s, aIdx) => {
              if (aIdx > 0) {
                richTextItems.push({ text: '\n' });
              }

              const label = this.getTimetableCardLabel(s);
              const isCanceled = s.status === 'canceled' || s.status === '폐강';

              if (isCanceled) {
                // 폐강 과목: 은은한 회색 + 취소선
                richTextItems.push({
                  text: label,
                  font: {
                    name: '맑은 고딕',
                    size: assigned.length > 2 ? 8.5 : 9.5,
                    strike: true,
                    color: { argb: 'FF94A3B8' }
                  }
                });
              } else {
                // 개설/조정 과목: 과목 고유의 선명한 글자색 + 굵게 (흰 배경에서 완벽한 가독성)
                const subName = s.subject || '';
                const colorInfo = subjectColorMap[subName] || { excelBg: 'E0F2FE', excelText: '0369A1' };
                const textColor = colorInfo.excelText || '0369A1';

                richTextItems.push({
                  text: label,
                  font: {
                    name: '맑은 고딕',
                    bold: true,
                    size: assigned.length > 2 ? 8.5 : 9.5,
                    color: { argb: 'FF' + textColor }
                  }
                });
              }
            });

            cell.value = { richText: richTextItems };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

            const neededHeight = Math.max(55, assigned.length * 22);
            if (row.height < neededHeight) {
              row.height = neededHeight;
            }
          } else {
            cell.value = '';
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
          }
          this.applyBorder(cell);
        });
      }

      currentExcelRow++;
    });
  },

  // -------------------------------------------------------------
  // 시간 중복 판별 헬퍼 (동일 요일/교시에 2개 이상의 유효 과목 수강 판정)
  // 반환: Map<number, string[]> (행 인덱스 -> 충돌 상대 과목명 배열)
  // -------------------------------------------------------------
  buildConflictMap(students) {
    const conflictMap = new Map();
    const slotMap = new Map(); // key -> Array<{ subject, idx }>

    (students || []).forEach((s, idx) => {
      const isAssigned = s.day && (s.periodId || s.period);
      if (!isAssigned) return;
      const normStatus = this.normalizeStatus(s.status);
      if (normStatus.label === '폐강') return; // 폐강 과목은 제외

      const studentKey = s.studentCode || this.getStudentCode(s) || s.studentMasterId || `${s.grade}_${s.classNum}_${s.studentNum}_${s.name}`;
      const periodKey = s.periodId || s.period;
      const slotKey = `${studentKey}__${s.day}__${periodKey}`;

      if (!slotMap.has(slotKey)) {
        slotMap.set(slotKey, []);
      }
      slotMap.get(slotKey).push({ subject: s.subject, idx });
    });

    slotMap.forEach((entries) => {
      if (entries.length > 1) {
        entries.forEach(curr => {
          const others = entries.filter(e => e.subject !== curr.subject).map(e => e.subject);
          const uniqueOthers = Array.from(new Set(others.filter(Boolean)));
          if (uniqueOthers.length > 0) {
            conflictMap.set(curr.idx, uniqueOthers);
          }
        });
      }
    });

    return conflictMap;
  },

  // 6. 결괏값 형태 2: 엑셀데이터 형태 (원천 데이터 테이블 / 통합 데이터 내보내기)
  async exportDataExcel(students, periods, days, academicYear = '2026', semester = '1', schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    const yearLabel = `${academicYear}학년도 ${semester}학기`;
    const filePrefix = `${academicYear}학년도_${semester}학기_${schoolName}`;
    workbook.creator = `${schoolName} 시간표 시스템`;
    const worksheet = workbook.addWorksheet(`${academicYear}학년도_${semester}학기_통합데이터`, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 8 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '개설상태', key: 'status', width: 12 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '시작시간', key: 'startTime', width: 12 },
      { header: '종료시간', key: 'endTime', width: 12 },
      { header: '시간중복', key: 'conflict', width: 22 },
      { header: '비고', key: 'memo', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const dayNameMap = { mon: '월요일', tue: '화요일', wed: '수요일', thu: '목요일', fri: '금요일', sat: '토요일', sun: '일요일' };
    const conflictMap = this.buildConflictMap(students);

    students.forEach((s, idx) => {
      const isAssigned = s.day && (s.periodId || s.period);
      const periodObj = periods.find(p => p.id === s.periodId || (!s.periodId && p.period === s.period));
      const periodName = isAssigned ? (periodObj ? periodObj.name : (s.period ? `${s.period}교시` : '미입력')) : '미입력';
      const startTime = isAssigned ? (periodObj ? periodObj.startTime : '미입력') : '미입력';
      const endTime = isAssigned ? (periodObj ? periodObj.endTime : '미입력') : '미입력';
      const dayText = isAssigned ? (dayNameMap[s.day] || s.day) : '미입력';

      const studentCode = this.getStudentCode(s);
      const normStatus = this.normalizeStatus(s.status);
      const statusLabel = normStatus.label;

      const conflicting = conflictMap.get(idx) || [];
      const conflictText = conflicting.length > 0 ? `⚠️ 중복: ${conflicting.join(', ')}` : '';

      const row = worksheet.addRow({
        index: idx + 1,
        year: academicYear,
        term: `${semester}학기`,
        grade: s.grade ? `${s.grade}` : '1',
        classNum: s.classNum ? `${s.classNum}` : '1',
        studentNum: s.studentNum ? `${s.studentNum}` : '',
        studentCode: studentCode,
        name: s.name,
        subject: s.subject || '미입력',
        status: statusLabel,
        day: dayText,
        periodName: periodName,
        startTime: startTime,
        endTime: endTime,
        conflict: conflictText,
        memo: s.memo || ''
      });

      row.height = 24;
      row.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);

        // 개설상태 열 (col 10)
        if (colNum === 10) {
          if (statusLabel === '확정') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            cell.font = { bold: true, color: { argb: 'FF065F46' } };
          } else if (statusLabel === '조정') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            cell.font = { bold: true, color: { argb: 'FFB45309' } };
          } else if (statusLabel === '폐강') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            cell.font = { strike: true, color: { argb: 'FF64748B' } };
          }
        }

        // 미입력 셀 하이라이트 (배정요일, 배정교시, 시작시간, 종료시간)
        if ([11, 12, 13, 14].includes(colNum) && cell.value === '미입력') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
          cell.font = { color: { argb: 'FF94A3B8' } };
        }

        // 시간중복 열 (col 15)
        if (colNum === 15 && conflictText) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E6' } };
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFE11D48' }, size: 9 };
        }
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_시간표_통합데이터_${this.getTimestamp()}.xlsx`);
  },

  // 7. 학년도 전체 (1학기 + 2학기) 통합 엑셀 다운로드 (작성 데이터 포함)
  async exportYearFullDataExcel(academicYear = '2026', semesterList = [], schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    const filePrefix = `${academicYear}학년도_1·2학기_${schoolName}`;
    workbook.creator = `${schoolName} 시간표 시스템`;
    const worksheet = workbook.addWorksheet(`${academicYear}학년도_전체(1·2학기)_통합데이터`, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 10 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '개설상태', key: 'status', width: 12 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '시작시간', key: 'startTime', width: 12 },
      { header: '종료시간', key: 'endTime', width: 12 },
      { header: '시간중복', key: 'conflict', width: 22 },
      { header: '비고', key: 'memo', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const dayNameMap = { mon: '월요일', tue: '화요일', wed: '수요일', thu: '목요일', fri: '금요일', sat: '토요일', sun: '일요일' };

    let globalRowIdx = 1;

    semesterList.forEach(({ semester, students, periods }) => {
      const conflictMap = this.buildConflictMap(students);

      (students || []).forEach((s, sIdx) => {
        const isAssigned = s.day && (s.periodId || s.period);
        const periodObj = (periods || []).find(p => p.id === s.periodId || (!s.periodId && p.period === s.period));
        const periodName = isAssigned ? (periodObj ? periodObj.name : (s.period ? `${s.period}교시` : '미입력')) : '미입력';
        const startTime = isAssigned ? (periodObj ? periodObj.startTime : '미입력') : '미입력';
        const endTime = isAssigned ? (periodObj ? periodObj.endTime : '미입력') : '미입력';
        const dayText = isAssigned ? (dayNameMap[s.day] || s.day) : '미입력';

        const studentCode = this.getStudentCode(s);
        const normStatus = this.normalizeStatus(s.status);
        const statusLabel = normStatus.label;

        const conflicting = conflictMap.get(sIdx) || [];
        const conflictText = conflicting.length > 0 ? `⚠️ 중복: ${conflicting.join(', ')}` : '';

        const row = worksheet.addRow({
          index: globalRowIdx++,
          year: academicYear,
          term: `${semester}학기`,
          grade: s.grade ? `${s.grade}` : '1',
          classNum: s.classNum ? `${s.classNum}` : '1',
          studentNum: s.studentNum ? `${s.studentNum}` : '',
          studentCode: studentCode,
          name: s.name,
          subject: s.subject || '미입력',
          status: statusLabel,
          day: dayText,
          periodName: periodName,
          startTime: startTime,
          endTime: endTime,
          conflict: conflictText,
          memo: s.memo || ''
        });

        row.height = 24;
        row.eachCell((cell, colNum) => {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.font = { name: '맑은 고딕', size: 10 };
          this.applyBorder(cell);

          // 개설상태 열 (col 10)
          if (colNum === 10) {
            if (statusLabel === '확정') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
              cell.font = { bold: true, color: { argb: 'FF065F46' } };
            } else if (statusLabel === '조정') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
              cell.font = { bold: true, color: { argb: 'FFB45309' } };
            } else if (statusLabel === '폐강') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
              cell.font = { strike: true, color: { argb: 'FF64748B' } };
            }
          }

          // 학기 열 구분 강조 (1학기: 연한 하늘, 2학기: 연한 보라)
          if (colNum === 3) {
            if (String(semester) === '1') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
              cell.font = { bold: true, color: { argb: 'FF0369A1' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDE9FE' } };
              cell.font = { bold: true, color: { argb: 'FF6D28D9' } };
            }
          }

          // 미입력 셀 하이라이트
          if ([11, 12, 13, 14].includes(colNum) && cell.value === '미입력') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
            cell.font = { color: { argb: 'FF94A3B8' } };
          }

          // 시간중복 열 (col 15)
          if (colNum === 15 && conflictText) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E6' } };
            cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFE11D48' }, size: 9 };
          }
        });
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_시간표_통합데이터_${this.getTimestamp()}.xlsx`);
  },

  // -------------------------------------------------------------
  // 8. 1단계 전용: 학생 및 과목 편제 명단 엑셀 다운로드 (1학기 + 2학기 연속 출력 & 옅은 파스텔 음영 구분)
  // -------------------------------------------------------------
  async exportStudentRosterExcel(optionsOrStudents, maybeEnrollments, academicYear = '2026', semester = '1', schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;

    let year = academicYear;
    let gradeLabel = '';
    let semestersData = [];

    // 옵션 객체 형태인지 레거시 단일 학기 호출인지 판별
    if (optionsOrStudents && typeof optionsOrStudents === 'object' && !Array.isArray(optionsOrStudents) && (optionsOrStudents.sem1Students || optionsOrStudents.sem2Students || optionsOrStudents.semesters)) {
      const opts = optionsOrStudents;
      year = opts.academicYear || year;
      gradeLabel = opts.grade ? `${opts.grade}학년` : '';
      schoolName = opts.schoolName || schoolName;

      if (Array.isArray(opts.semesters)) {
        semestersData = opts.semesters;
      } else {
        // 기본 1학기 + 2학기 순서 구조 (1학기 먼저, 2학기 이어서)
        semestersData = [
          { semester: '1', students: opts.sem1Students || [], enrollments: opts.sem1Enrollments || {} },
          { semester: '2', students: opts.sem2Students || [], enrollments: opts.sem2Enrollments || {} }
        ];
      }
    } else {
      // 레거시 단일 학기 호출 대응
      semestersData = [
        { semester: String(semester), students: optionsOrStudents || [], enrollments: maybeEnrollments || {} }
      ];
    }

    const titleGrade = gradeLabel ? `_${gradeLabel}` : '';
    const filePrefix = `${year}학년도${titleGrade}_1·2학기통합_${schoolName}`;
    const sheetName = `${year}학년도${titleGrade}_학생명단`.slice(0, 31);
    const worksheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 10 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '편제상태', key: 'enrollStatus', width: 12 },
      { header: '수강과목수', key: 'courseCount', width: 12 },
      { header: '수강과목목록', key: 'courses', width: 40 }
    ];

    // 헤더 행 스타일링
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // 세련된 딥 슬레이트
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    let globalRowIdx = 1;

    // 학기별 옅은 파스텔톤 테마 (1학기: 연한 하늘색, 2학기: 연한 보라색)
    const semesterTheme = {
      '1': {
        // 1학기: 연한 하늘색 계열 (Sky)
        rowBg: 'FFF0F9FF',        // 행 전체 은은한 연하늘색
        termBg: 'FFE0F2FE',       // 학기 셀 연하늘 뱃지
        termText: 'FF0369A1'      // 학기 셀 진파랑 볼드
      },
      '2': {
        // 2학기: 연한 보라색 계열 (Purple)
        rowBg: 'FFFAF5FF',        // 행 전체 은은한 연보라색
        termBg: 'FFEDE9FE',       // 학기 셀 연보라 뱃지
        termText: 'FF6D28D9'      // 학기 셀 진보라 볼드
      },
      default: {
        rowBg: 'FFFFFFFF',
        termBg: 'FFF1F5F9',
        termText: 'FF334155'
      }
    };

    semestersData.forEach(({ semester, students, enrollments }) => {
      const termLabel = `${semester}학기`;
      const theme = semesterTheme[String(semester)] || semesterTheme.default;

      (students || []).forEach(st => {
        const code = st.studentCode || this.getStudentCode(st);
        const courses = [];
        Object.entries(enrollments || {}).forEach(([sub, ids]) => {
          if (ids && ids.includes(st.id)) courses.push(sub);
        });

        const isEnrolled = courses.length > 0;
        const statusLabel = isEnrolled ? '편제완료' : '미편성';
        const courseCountText = `${courses.length}과목`;
        const coursesText = isEnrolled ? courses.join(', ') : '미편성';

        const row = worksheet.addRow({
          index: globalRowIdx,
          year: year,
          term: termLabel,
          grade: st.grade ? `${st.grade}` : '1',
          classNum: st.classNum ? `${st.classNum}` : '1',
          studentNum: st.studentNum ? `${st.studentNum}` : '',
          studentCode: code,
          name: st.name || '',
          enrollStatus: statusLabel,
          courseCount: courseCountText,
          courses: coursesText
        });

        row.height = 24;
        row.eachCell((cell, colNum) => {
          cell.alignment = { vertical: 'middle', horizontal: colNum === 11 ? 'left' : 'center' };
          cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF1E293B' } };
          this.applyBorder(cell);

          // 행 전체 옅은 파스텔 배경 적용 (1학기: 연하늘, 2학기: 연보라)
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.rowBg } };

          // 학기 열 (col 3): 학기 전용 강조 뱃지 색상
          if (colNum === 3) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
            cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
          }

          // 편제상태 열 (col 9): 편제완료(연초록) / 미편성(연주황)
          if (colNum === 9) {
            if (isEnrolled) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF047857' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFB45309' } };
            }
          }

          // 수강과목수 열 (col 10)
          if (colNum === 10) {
            if (!isEnrolled) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFB45309' } };
            } else {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText } };
            }
          }

          // 수강과목목록 열 (col 11)
          if (colNum === 11 && !isEnrolled) {
            cell.font = { name: '맑은 고딕', color: { argb: 'FF94A3B8' }, italic: true };
          }
        });

        globalRowIdx++;
      });
    });

    // 자동 필터 적용
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 11 }
    };

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_학생_과목편제명단_${this.getTimestamp()}.xlsx`);
  },

  // -------------------------------------------------------------
  // 9. 학기-과목별 수강자 명단 및 인원 엑셀 다운로드 (1학기 -> 2학기 연속 출력, 학년/반/번호 정렬)
  // -------------------------------------------------------------
  async exportSubjectRosterExcel({ academicYear = '2026', schoolName = '강원온라인학교', sem1Subjects = [], sem2Subjects = [] }) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;
    const filePrefix = `${academicYear}학년도_${schoolName}_학기_과목별_수강자명단`;

    // -----------------------------------------------------------
    // 시트 1: 과목별 상세 수강자 명단
    // -----------------------------------------------------------
    const detailSheet = workbook.addWorksheet(`${academicYear}학년도_과목별수강명단`, { views: [{ showGridLines: true }] });
    detailSheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학기', key: 'term', width: 10 },
      { header: '과목명', key: 'subject', width: 26 },
      { header: '과목 수강인원', key: 'subjectTotal', width: 14 },
      { header: '과목내 번호', key: 'subjectSeq', width: 12 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '상태', key: 'status', width: 12 }
    ];

    const headerRow1 = detailSheet.getRow(1);
    headerRow1.height = 30;
    headerRow1.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // 세련된 딥 슬레이트
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    let globalRowIdx = 1;

    // 학기별 파스텔 테마 (1학기: 연한 하늘색, 2학기: 연한 보라색)
    const semesterThemes = {
      '1': {
        rowBg: 'FFF0F9FF',        // 행 전체 은은한 연하늘색
        termBg: 'FFE0F2FE',       // 학기 셀 연하늘 뱃지
        termText: 'FF0369A1',     // 학기 셀 진파랑 볼드
        subText: 'FF0C4A6E'       // 과목명 진파랑
      },
      '2': {
        rowBg: 'FFFAF5FF',        // 행 전체 은은한 연보라색
        termBg: 'FFEDE9FE',       // 학기 셀 연보라 뱃지
        termText: 'FF6D28D9',     // 학기 셀 진보라 볼드
        subText: 'FF581C87'       // 과목명 진보라
      }
    };

    const processSemesterData = (semester, subjectList) => {
      const termLabel = `${semester}학기`;
      const theme = semesterThemes[String(semester)] || semesterThemes['1'];

      (subjectList || []).forEach(subItem => {
        const subName = subItem.subject || '';
        const count = subItem.count || (subItem.students ? subItem.students.length : 0);
        const countText = `${count}명`;
        const students = subItem.students || [];
        const statusText = subItem.status === 'canceled' ? '폐강' : (subItem.status === 'adjusted' ? '조정' : '개설');

        if (students.length === 0) {
          // 수강생이 없는 과목도 명단에 1행으로 명시
          const row = detailSheet.addRow({
            index: globalRowIdx,
            term: termLabel,
            subject: subName,
            subjectTotal: '0명',
            subjectSeq: '-',
            grade: '-',
            classNum: '-',
            studentNum: '-',
            studentCode: '-',
            name: '(수강생 없음)',
            status: statusText
          });
          row.height = 24;
          row.eachCell((cell, colNum) => {
            cell.alignment = { vertical: 'middle', horizontal: colNum === 3 ? 'left' : 'center' };
            cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF94A3B8' } };
            this.applyBorder(cell);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.rowBg } };
            if (colNum === 2) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
            }
            if (colNum === 3) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.subText }, size: 10 };
            }
          });
          globalRowIdx++;
          return;
        }

        students.forEach((st, sIdx) => {
          const code = st.studentCode || this.getStudentCode(st);
          const row = detailSheet.addRow({
            index: globalRowIdx,
            term: termLabel,
            subject: subName,
            subjectTotal: countText,
            subjectSeq: `${sIdx + 1} / ${count}`,
            grade: st.grade ? `${st.grade}` : '',
            classNum: st.classNum ? `${st.classNum}` : '',
            studentNum: st.studentNum ? `${st.studentNum}` : '',
            studentCode: code,
            name: st.name || '',
            status: statusText
          });

          row.height = 24;
          row.eachCell((cell, colNum) => {
            cell.alignment = { vertical: 'middle', horizontal: colNum === 3 ? 'left' : 'center' };
            cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF1E293B' } };
            this.applyBorder(cell);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.rowBg } };

            // 학기 열 (col 2)
            if (colNum === 2) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
            }
            // 과목명 열 (col 3)
            if (colNum === 3) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.subText }, size: 10 };
            }
            // 과목 수강인원 열 (col 4)
            if (colNum === 4) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
            }
            // 과목내 번호 열 (col 5)
            if (colNum === 5) {
              cell.font = { name: '맑은 고딕', size: 9, color: { argb: 'FF64748B' } };
            }
            // 성명 열 (col 10)
            if (colNum === 10) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF0F172A' } };
            }
            // 상태 열 (col 11)
            if (colNum === 11) {
              if (subItem.status === 'canceled') {
                cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFDC2626' } };
              } else {
                cell.font = { name: '맑은 고딕', color: { argb: 'FF059669' } };
              }
            }
          });

          globalRowIdx++;
        });
      });
    };

    // 1학기 먼저 출력, 이어서 2학기 연속 출력
    processSemesterData('1', sem1Subjects);
    processSemesterData('2', sem2Subjects);

    detailSheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 11 }
    };

    // -----------------------------------------------------------
    // 시트 2: 과목별 수강인원 한눈에 보는 요약 통계
    // -----------------------------------------------------------
    const summarySheet = workbook.addWorksheet(`${academicYear}학년도_수강인원요약`, { views: [{ showGridLines: true }] });
    summarySheet.columns = [
      { header: '순번', key: 'idx', width: 8 },
      { header: '과목명', key: 'subject', width: 28 },
      { header: '1학기 수강인원', key: 'sem1Count', width: 16 },
      { header: '2학기 수강인원', key: 'sem2Count', width: 16 },
      { header: '연간 총 수강건수', key: 'totalCount', width: 18 },
      { header: '개설 상태', key: 'status', width: 14 }
    ];

    const headerRow2 = summarySheet.getRow(1);
    headerRow2.height = 30;
    headerRow2.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } }; // 진한 에메랄드/틸
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    // 1학기와 2학기 과목 통합 맵 구성
    const summaryMap = new Map();
    (sem1Subjects || []).forEach(s => {
      if (!summaryMap.has(s.subject)) {
        summaryMap.set(s.subject, { subject: s.subject, sem1: s.count || 0, sem2: 0, status: s.status || 'confirmed' });
      } else {
        summaryMap.get(s.subject).sem1 = s.count || 0;
      }
    });
    (sem2Subjects || []).forEach(s => {
      if (!summaryMap.has(s.subject)) {
        summaryMap.set(s.subject, { subject: s.subject, sem1: 0, sem2: s.count || 0, status: s.status || 'confirmed' });
      } else {
        summaryMap.get(s.subject).sem2 = s.count || 0;
      }
    });

    const sortedSummary = Array.from(summaryMap.values()).sort((a, b) => a.subject.localeCompare(b.subject, 'ko'));
    let totalSem1 = 0;
    let totalSem2 = 0;

    sortedSummary.forEach((item, sIdx) => {
      const sum = item.sem1 + item.sem2;
      totalSem1 += item.sem1;
      totalSem2 += item.sem2;
      const statusLabel = item.status === 'canceled' ? '폐강' : '정상 개설';

      const row = summarySheet.addRow({
        idx: sIdx + 1,
        subject: item.subject,
        sem1Count: `${item.sem1}명`,
        sem2Count: `${item.sem2}명`,
        totalCount: `${sum}명`,
        status: statusLabel
      });

      row.height = 24;
      row.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: colNum === 2 ? 'left' : 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);

        // 줄무늬 배경
        if (sIdx % 2 === 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }
        // 1학기 셀 (연하늘 폰트)
        if (colNum === 3 && item.sem1 > 0) {
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF0369A1' } };
        }
        // 2학기 셀 (연보라 폰트)
        if (colNum === 4 && item.sem2 > 0) {
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF6D28D9' } };
        }
        // 총 수강건수 셀
        if (colNum === 5) {
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF0F172A' } };
        }
      });
    });

    // 합계 요약 행
    const sumRow = summarySheet.addRow({
      idx: '합계',
      subject: `총 ${sortedSummary.length}개 과목`,
      sem1Count: `${totalSem1}명`,
      sem2Count: `${totalSem2}명`,
      totalCount: `${totalSem1 + totalSem2}명`,
      status: '-'
    });
    sumRow.height = 28;
    sumRow.eachCell((cell, colNum) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.font = { name: '맑은 고딕', bold: true, size: 10, color: { argb: 'FF0F172A' } };
      cell.alignment = { vertical: 'middle', horizontal: colNum === 2 ? 'left' : 'center' };
      this.applyBorder(cell);
    });

    summarySheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 6 }
    };

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_${this.getTimestamp()}.xlsx`);
  },

  applyBorder(cell) {
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };
  },

  async saveFile(buffer, fileName) {
    if (window.FileStorageService && typeof window.FileStorageService.saveFile === 'function') {
      try {
        const res = await window.FileStorageService.saveFile(buffer, fileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        if (res && res.savedToFolder) {
          if (window.app && typeof window.app.showToast === 'function') {
            window.app.showToast(`[${res.folderName}] 폴더에 '${res.fileName}' 파일이 안전하게 저장되었습니다.`, true);
          }
          return;
        }
      } catch (err) {
        console.warn('FileStorageService save failed, fallback to standard download:', err);
      }
    }
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 100);
  },

  formatTimeString(str) {
    if (!str) return '';
    const clean = str.replace(/[^0-9:]/g, '');
    const parts = clean.split(':');
    if (parts.length >= 2) {
      const h = parts[0].padStart(2, '0');
      const m = parts[1].padStart(2, '0');
      return `${h}:${m}`;
    }
    return str;
  },

  getTimestamp() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    return `${y}${m}${d}_${hh}${mm}`;
  }
};
