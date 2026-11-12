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
    return `${sym} ${sub}(${codeName})`;
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
    const yearLabel = `${academicYear}학년도 ${semester}학기`;
    const filePrefix = `${academicYear}학년도_${semester}학기_${schoolName}`;
    workbook.creator = `${schoolName} 시간표 시스템`;

    // 5-1. 종합 시간표 시트 생성
    const totalSheet = workbook.addWorksheet('종합 시간표', { views: [{ showGridLines: true }] });
    this.buildTimetableSheet(totalSheet, periods, days, students, subjectColorMap, `${yearLabel} ${schoolName} 전체 종합 시간표`.trim());

    // 5-2. 학년-반별 개별 시트 자동 생성
    const classSet = new Set();
    students.forEach(s => {
      if (s.grade && s.classNum) {
        classSet.add(`${s.grade}-${s.classNum}`);
      }
    });

    const sortedClasses = Array.from(classSet).sort((a, b) => {
      const [gA, cA] = a.split('-').map(Number);
      const [gB, cB] = b.split('-').map(Number);
      return gA === gB ? cA - cB : gA - gB;
    });

    for (const cls of sortedClasses) {
      const [grade, classNum] = cls.split('-').map(Number);
      const sheetName = `${grade}학년 ${classNum}반`;
      const classStudents = students.filter(s => s.grade === grade && s.classNum === classNum);
      const classSheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });
      this.buildTimetableSheet(
        classSheet,
        periods,
        days,
        classStudents,
        subjectColorMap,
        `${yearLabel} ${schoolName} [${grade}학년 ${classNum}반] 시간표`.trim()
      );
    }

    // 5-3. [수강자 전체 명단] 시트 구성 (항목: 학년도, 학기, 과목명, 학번, 학생 이름, 개설/조정/폐강)
    const rosterSheet = workbook.addWorksheet('수강자 전체 명단', { views: [{ showGridLines: true }] });
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

    // 과목 상태 텍스트 판정 헬퍼
    const getStatusLabel = (sub) => {
      if (!courseStatus) return '개설';
      const st = courseStatus[sub] || 'open';
      if (st === 'closed') return '폐강';
      if (st === 'adjust') return '조정';
      if (st === 'excluded') return '폐강'; // '제외'는 폐강으로 표기
      return '개설';
    };

    // 수강자 목록 데이터 작성
    const rosterRows = [];
    const studentMasterMap = new Map();
    if (masterStudents && Array.isArray(masterStudents)) {
      masterStudents.forEach(m => studentMasterMap.set(m.id, m));
    }

    if (courseEnrollments && typeof courseEnrollments === 'object') {
      Object.entries(courseEnrollments).forEach(([subject, studentIds]) => {
        if (!Array.isArray(studentIds)) return;
        const statusText = getStatusLabel(subject);
        studentIds.forEach(id => {
          const m = studentMasterMap.get(id);
          if (m) {
            const code = m.studentCode || this.getStudentCode(m);
            rosterRows.push({
              academicYear: `${academicYear}학년도`,
              semester: `${semester}학기`,
              subject: subject,
              studentCode: code,
              studentName: m.name,
              statusText: statusText
            });
          }
        });
      });
    }

    // 만약 courseEnrollments가 비어있거나 없는 경우 students 배열에서 유니크하게 추출
    if (rosterRows.length === 0 && students && Array.isArray(students)) {
      const seen = new Set();
      students.forEach(s => {
        if (!s.subject || !s.name) return;
        const code = s.studentCode || this.getStudentCode(s);
        const key = `${code}_${s.subject}`;
        if (!seen.has(key)) {
          seen.add(key);
          const statusText = getStatusLabel(s.subject);
          rosterRows.push({
            academicYear: `${academicYear}학년도`,
            semester: `${semester}학기`,
            subject: s.subject,
            studentCode: code,
            studentName: s.name,
            statusText: statusText
          });
        }
      });
    }

    // 정렬: 1순위 과목명 가나다순, 2순위 학번 오름차순 정렬
    rosterRows.sort((a, b) => {
      const subComp = (a.subject || '').localeCompare(b.subject || '', 'ko');
      if (subComp !== 0) return subComp;
      return (a.studentCode || '').localeCompare(b.studentCode || '', 'ko', { numeric: true });
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

      row.eachCell((cell, colNum) => {
        cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF1E293B' } };
        cell.alignment = { vertical: 'middle', horizontal: colNum === 3 ? 'left' : 'center' };

        // 6열: 개설/조정/폐강 상태별 텍스트 색상
        if (colNum === 6) {
          if (rData.statusText === '개설') {
            cell.font = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FF059669' } }; // 진한 초록
          } else if (rData.statusText === '조정') {
            cell.font = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FFD97706' } }; // 진한 주황/노랑
          } else if (rData.statusText === '폐강') {
            cell.font = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FFDC2626' } }; // 진한 빨강
          }
        }

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isEven ? 'FFF8FAFC' : 'FFFFFFFF' }
        };
        this.applyBorder(cell);
      });
    });

    // 5-4. [신규 탭] 1,2학기 통합시간표 시트 생성 (1학기/2학기 다운로드 무관 항상 포함)
    if (yearData && typeof yearData === 'object') {
      const unifiedSheet = workbook.addWorksheet('1,2학기 통합시간표', { views: [{ showGridLines: true }] });
      this.buildUnifiedSemesterTimetableSheet(unifiedSheet, periods, academicYear, schoolName, yearData, noticeText, subjectColorMap);
    }

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_학급_시간표_${this.getTimestamp()}.xlsx`);
  },

  // 개별 시간표 시트 디자인 생성 헬퍼
            // =========================================================================
  // 5-4. [신규] 1,2학기 통합시간표 시트 디자인 생성 헬퍼
  // 월화수목금 각 4개 열 (1학기 2개열: 과목명/(학번) 학생명, 2학기 2개열: 과목명/(학번) 학생명)
  // - [인쇄 설정] 여백: 좁게, 용지 방향: 가로, 크기: B4, 인쇄제목 반복할 행: 1~5행
  // - [상태 표기] 과목명 오른쪽에 (개설/조정/폐강) 표기 (글자 크기 10pt, 과목명 13pt)
  // - 1학기/2학기 어떤 탭에서 다운로드하든 양 학기 데이터를 100% 온전히 추출 및 표시
  // - 타이틀 제외 데이터 글자 크기 13pt, '(학번)' 부분은 9pt (Rich Text 적용)
  // - 과목명에 연한 파스텔톤 배경색을 채워 시각적 구별성 강화 (과목별 고유 파스텔톤)
  // - '과목명'은 세로 병합을 먼저 수행한 후 파스텔 배경 및 왼쪽 정렬/들여쓰기 1(indent: 1) 적용
  // - 안내 사항 기입 기능 지원 (타이틀 아래 돋보이는 디자인 박스)
  // =========================================================================
  buildUnifiedSemesterTimetableSheet(sheet, periods, academicYear, schoolName, yearData, noticeText = '', subjectColorMap = {}) {
    const sem1Data = (yearData && yearData['1']) || {};
    const sem2Data = (yearData && yearData['2']) || {};

    // -------------------------------------------------------------------------
    // ★ 페이지 인쇄 설정: 여백 좁게, 가로 방향, B4 크기, 인쇄제목 1~5행 반복
    // -------------------------------------------------------------------------
    sheet.pageSetup = {
      orientation: 'landscape',  // 용지 방향: 가로
      paperSize: 12,            // 크기: B4 (JIS)
      margins: {                // 여백: 좁게 (Narrow)
        left: 0.25,
        right: 0.25,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3
      },
      printTitlesRow: '1:5',    // 인쇄제목 - 반복할 행 : 1~5행
      fitToPage: true,          // 페이지 맞춤
      fitToWidth: 1,            // 가로 1페이지 폭에 맞춤
      fitToHeight: 0            // 세로는 내용에 따라 자동 분할
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

    // 열 너비 설정 (A열: 교시, 각 요일당 4개 열씩 총 21개 열, 13pt 글자에 맞게 넉넉히 설정)
    const columnsConfig = [
      { width: 16 } // A: 교시
    ];
    targetDays.forEach(() => {
      columnsConfig.push({ width: 23 }); // 1학기 과목명 (상태 포함)
      columnsConfig.push({ width: 20 }); // 1학기 (학번) 학생명
      columnsConfig.push({ width: 23 }); // 2학기 과목명 (상태 포함)
      columnsConfig.push({ width: 20 }); // 2학기 (학번) 학생명
    });
    sheet.columns = columnsConfig;

    // 1행: 타이틀 (A1:U1 병합)
    sheet.mergeCells(1, 1, 1, 21);
    const titleCell = sheet.getCell(1, 1);
    titleCell.value = `${academicYear}학년도 ${schoolName} 1·2학기 통합시간표`;
    titleCell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF1E293B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(1).height = 42;

    let rowDayIdx = 2;
    // 2행: 안내 사항이 있는 경우 강조 박스로 삽입
    const trimmedNotice = (noticeText || '').trim();
    if (trimmedNotice) {
      sheet.mergeCells(2, 1, 2, 21);
      const noticeCell = sheet.getCell(2, 1);
      noticeCell.value = `📢 [안내 사항]\n${trimmedNotice}`;
      noticeCell.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FF1E293B' } };
      noticeCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true, indent: 1 };
      
      const lines = trimmedNotice.split('\n').length;
      sheet.getRow(2).height = Math.max(34, (lines + 1) * 20 + 8);

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

    // 요일 대분류 헤더 (A열 교시, 각 요일 4개열씩 병합)
    const rowDay = sheet.getRow(rowDayIdx);
    rowDay.height = 30;
    targetDays.forEach((day, idx) => {
      const startCol = 2 + idx * 4;
      const endCol = startCol + 3;
      sheet.mergeCells(rowDayIdx, startCol, rowDayIdx, endCol);
      const dayCell = sheet.getCell(rowDayIdx, startCol);
      dayCell.value = day.name;
      dayCell.font = { name: '맑은 고딕', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      dayCell.alignment = { vertical: 'middle', horizontal: 'center' };
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(rowDayIdx, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        this.applyBorder(cell);
      }
    });

    // 학기 중분류 헤더 (1학기 2개열, 2학기 2개열 병합)
    const rowSem = sheet.getRow(rowSemIdx);
    rowSem.height = 26;
    targetDays.forEach((day, idx) => {
      const colSem1 = 2 + idx * 4;
      const colSem2 = colSem1 + 2;

      // 1학기 헤더 (2개 열 병합)
      sheet.mergeCells(rowSemIdx, colSem1, rowSemIdx, colSem1 + 1);
      const sem1Cell = sheet.getCell(rowSemIdx, colSem1);
      sem1Cell.value = '1학기';
      sem1Cell.font = { name: '맑은 고딕', size: 12, bold: true, color: { argb: 'FF1D4ED8' } };
      sem1Cell.alignment = { vertical: 'middle', horizontal: 'center' };
      for (let c = colSem1; c <= colSem1 + 1; c++) {
        const cell = sheet.getCell(rowSemIdx, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
        this.applyBorder(cell);
      }

      // 2학기 헤더 (2개 열 병합)
      sheet.mergeCells(rowSemIdx, colSem2, rowSemIdx, colSem2 + 1);
      const sem2Cell = sheet.getCell(rowSemIdx, colSem2);
      sem2Cell.value = '2학기';
      sem2Cell.font = { name: '맑은 고딕', size: 12, bold: true, color: { argb: 'FF4338CA' } };
      sem2Cell.alignment = { vertical: 'middle', horizontal: 'center' };
      for (let c = colSem2; c <= colSem2 + 1; c++) {
        const cell = sheet.getCell(rowSemIdx, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF2FF' } };
        this.applyBorder(cell);
      }
    });

    // 소분류 헤더 (과목명, (학번) 학생명)
    const rowSubj = sheet.getRow(rowSubjIdx);
    rowSubj.height = 26;
    targetDays.forEach((day, idx) => {
      const colBase = 2 + idx * 4;
      const headers = ['과목명', '(학번) 학생명', '과목명', '(학번) 학생명'];
      headers.forEach((h, hIdx) => {
        const cell = sheet.getCell(rowSubjIdx, colBase + hIdx);
        cell.value = h;
        cell.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FF334155' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        this.applyBorder(cell);
      });
    });

    // A열 교시 헤더 병합 (rowDayIdx ~ rowSubjIdx)
    sheet.mergeCells(rowDayIdx, 1, rowSubjIdx, 1);
    const colPeriodCell = sheet.getCell(rowDayIdx, 1);
    colPeriodCell.value = '교시';
    colPeriodCell.font = { name: '맑은 고딕', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    colPeriodCell.alignment = { vertical: 'middle', horizontal: 'center' };
    for (let r = rowDayIdx; r <= rowSubjIdx; r++) {
      const cell = sheet.getCell(r, 1);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      this.applyBorder(cell);
    }

    const statusMap1 = sem1Data.courseStatus || {};
    const statusMap2 = sem2Data.courseStatus || {};

    // =========================================================================
    // ★ 특정 학기 데이터(semData)에서 특정 요일/교시 슬롯의 과목 및 학생 목록 추출 헬퍼
    // =========================================================================
    const getSlotItems = (semData, dayId, period) => {
      if (!semData || typeof semData !== 'object') return [];

      const periodIdStr = String(period.id);
      const periodNum = period.period;

      // 1) 우선 semData.students 배열(시간표에 매핑된 최종 평탄화 데이터)이 있는 경우 최우선 활용
      if (Array.isArray(semData.students) && semData.students.length > 0) {
        const matching = semData.students.filter(s => {
          if (!s || !s.subject) return false;
          if (s.day !== dayId) return false;
          const matchPeriodId = s.periodId && String(s.periodId) === periodIdStr;
          const matchPeriodNum = s.period !== undefined && s.period !== null && s.period === periodNum;
          return matchPeriodId || matchPeriodNum;
        });

        if (matching.length > 0) {
          // 과목별로 학생 그룹화
          const subjectMap = new Map();
          matching.forEach(s => {
            if (!subjectMap.has(s.subject)) {
              subjectMap.set(s.subject, []);
            }
            if (s.name) {
              const code = s.studentCode || (s.grade && s.classNum && s.studentNum ? `${s.grade}${String(s.classNum).padStart(2,'0')}${String(s.studentNum).padStart(2,'0')}` : '');
              subjectMap.get(s.subject).push({
                stId: code,
                stName: s.name
              });
            }
          });

          const items = [];
          const sortedSubjects = Array.from(subjectMap.keys()).sort((a, b) => a.localeCompare(b, 'ko'));
          sortedSubjects.forEach(sub => {
            const stList = subjectMap.get(sub);
            // 학번/이름순 정렬
            stList.sort((a, b) => {
              if (a.stId && b.stId) return a.stId.localeCompare(b.stId, 'ko', { numeric: true });
              return a.stName.localeCompare(b.stName, 'ko');
            });

            if (stList.length === 0) {
              items.push({ subject: sub, stId: '', stName: '-' });
            } else {
              stList.forEach(st => {
                items.push({ subject: sub, stId: st.stId, stName: st.stName });
              });
            }
          });
          return items;
        }
      }

      // 2) semData.courseSchedules (배열 및 객체 모두 수용) 및 courseEnrollments 에서 직접 추출
      const schedules = semData.courseSchedules || {};
      const enrollments = semData.courseEnrollments || {};
      const masterStudents = semData.masterStudents || [];

      // masterStudents 맵 구축 (id 기준 및 name 기준 이중 맵)
      const masterIdMap = new Map();
      const masterNameMap = new Map();
      masterStudents.forEach(m => {
        if (!m) return;
        const code = m.studentCode || (m.grade && m.classNum && m.studentNum ? `${m.grade}${String(m.classNum).padStart(2,'0')}${String(m.studentNum).padStart(2,'0')}` : '');
        if (m.id) masterIdMap.set(m.id, { id: m.id, code, name: m.name });
        if (m.name) masterNameMap.set(m.name, { id: m.id, code, name: m.name });
      });

      const matchedSubjects = [];

      Object.entries(schedules).forEach(([subject, scheduleSlots]) => {
        if (!scheduleSlots) return;

        // scheduleSlots 가 배열인 경우([ {day, periodId}, ... ])와 단일 객체인 경우 모두 지원
        const slotsArray = Array.isArray(scheduleSlots) ? scheduleSlots : [scheduleSlots];
        const isMatched = slotsArray.some(slot => {
          if (!slot || slot.day !== dayId) return false;
          const matchId = slot.periodId && String(slot.periodId) === periodIdStr;
          const matchNum = slot.period !== undefined && slot.period !== null && slot.period === periodNum;
          return matchId || matchNum;
        });

        if (isMatched) {
          matchedSubjects.push(subject);
        }
      });

      matchedSubjects.sort((a, b) => a.localeCompare(b, 'ko'));

      const items = [];
      matchedSubjects.forEach(subject => {
        const enrolled = enrollments[subject] || [];
        const studentItems = [];

        enrolled.forEach(stRef => {
          // stRef 는 학생 id 일 수도 있고 name 일 수도 있음
          const m = masterIdMap.get(stRef) || masterNameMap.get(stRef);
          if (m) {
            studentItems.push({ stId: m.code || '', stName: m.name });
          } else {
            studentItems.push({ stId: '', stName: String(stRef) });
          }
        });

        // 정렬
        studentItems.sort((a, b) => {
          if (a.stId && b.stId) return a.stId.localeCompare(b.stId, 'ko', { numeric: true });
          return a.stName.localeCompare(b.stName, 'ko');
        });

        if (studentItems.length === 0) {
          items.push({ subject, stId: '', stName: '-' });
        } else {
          studentItems.forEach(st => {
            items.push({ subject, stId: st.stId, stName: st.stName });
          });
        }
      });

      return items;
    };

    // 교시별 행 생성
    (periods || []).forEach(period => {
      // 월~금 각각의 1학기/2학기 데이터 추출
      const daySlotsData = targetDays.map(day => {
        const items1 = getSlotItems(sem1Data, day.id, period);
        const items2 = getSlotItems(sem2Data, day.id, period);
        return { day, items1, items2 };
      });

      // 이 교시 블록의 최대 행 수 계산
      let maxBlockRows = 1;
      daySlotsData.forEach(d => {
        const count = Math.max(d.items1.length, d.items2.length);
        if (count > maxBlockRows) maxBlockRows = count;
      });

      const blockStartRow = currentRow;
      const blockEndRow = blockStartRow + maxBlockRows - 1;

      // 행 높이 설정 (글자 크기 13pt에 맞춰 넉넉하게 29로 설정)
      for (let r = blockStartRow; r <= blockEndRow; r++) {
        sheet.getRow(r).height = 29;
      }

      // 교시 (A열) 세로 병합
      if (blockEndRow > blockStartRow) {
        sheet.mergeCells(blockStartRow, 1, blockEndRow, 1);
      }
      const periodCell = sheet.getCell(blockStartRow, 1);
      periodCell.value = `${period.name}\n(${period.startTime}~${period.endTime})`;
      periodCell.font = { name: '맑은 고딕', size: 13, bold: true, color: { argb: 'FF1E293B' } };
      periodCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

      for (let r = blockStartRow; r <= blockEndRow; r++) {
        const cell = sheet.getCell(r, 1);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        this.applyBorder(cell);
      }

      // 각 요일별 열 채우기
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

          if (item) {
            if (item.stId) {
              // 학번은 9pt, 학생명은 13pt
              cSt.value = {
                richText: [
                  { text: `(${item.stId}) `, font: { name: '맑은 고딕', size: 9, color: { argb: 'FF475569' } } },
                  { text: item.stName, font: { name: '맑은 고딕', size: 13, color: { argb: 'FF1E293B' } } }
                ]
              };
            } else {
              cSt.value = item.stName;
              cSt.font = { name: '맑은 고딕', size: 13, color: { argb: 'FF1E293B' } };
            }
          } else {
            cSt.value = '';
            cSt.font = { name: '맑은 고딕', size: 13 };
          }

          // (학번) 학생명: 왼쪽 정렬 + 들여쓰기 1
          cSt.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

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
            subSpans1.push({ start: spanStart, end: spanEnd, subject: d.items1[spanStart].subject });
            if (spanEnd > spanStart) {
              sheet.mergeCells(blockStartRow + spanStart, colSub1, blockStartRow + spanEnd, colSub1);
            }
            spanStart = spanEnd + 1;
          }
        }

        // 1학기: 3) 병합 완료 후 -> 과목명 13pt + (개설/조정/폐강) 10pt + 파스텔 배경 + 좌측정렬/들여쓰기 1
        subSpans1.forEach(span => {
          const topRow = blockStartRow + span.start;
          const endRow = blockStartRow + span.end;
          const cell = sheet.getCell(topRow, colSub1);
          const color = getSubjectColor(span.subject);

          if (span.subject) {
            const stText = getStatusText(statusMap1, span.subject);
            let stColor = 'FF059669'; // 개설: 진한 초록
            if (stText === '조정') stColor = 'FFD97706'; // 조정: 진한 주황
            else if (stText === '폐강') stColor = 'FFDC2626'; // 폐강: 진한 빨강

            cell.value = {
              richText: [
                {
                  text: span.subject + ' ',
                  font: {
                    name: '맑은 고딕',
                    size: 13,
                    bold: true,
                    color: { argb: color ? color.text : 'FF1E293B' }
                  }
                },
                {
                  text: `(${stText})`,
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

          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

          // 파스텔톤 배경색 채우기
          if (color) {
            for (let r = topRow; r <= endRow; r++) {
              const c = sheet.getCell(r, colSub1);
              c.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color.bg }
              };
              this.applyBorder(c);
            }
          } else {
            this.applyBorder(cell);
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

          if (item) {
            if (item.stId) {
              // 학번은 9pt, 학생명은 13pt
              cSt.value = {
                richText: [
                  { text: `(${item.stId}) `, font: { name: '맑은 고딕', size: 9, color: { argb: 'FF475569' } } },
                  { text: item.stName, font: { name: '맑은 고딕', size: 13, color: { argb: 'FF1E293B' } } }
                ]
              };
            } else {
              cSt.value = item.stName;
              cSt.font = { name: '맑은 고딕', size: 13, color: { argb: 'FF1E293B' } };
            }
          } else {
            cSt.value = '';
            cSt.font = { name: '맑은 고딕', size: 13 };
          }

          // (학번) 학생명: 왼쪽 정렬 + 들여쓰기 1
          cSt.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

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
            subSpans2.push({ start: spanStart, end: spanEnd, subject: d.items2[spanStart].subject });
            if (spanEnd > spanStart) {
              sheet.mergeCells(blockStartRow + spanStart, colSub2, blockStartRow + spanEnd, colSub2);
            }
            spanStart = spanEnd + 1;
          }
        }

        // 2학기: 3) 병합 완료 후 -> 과목명 13pt + (개설/조정/폐강) 10pt + 파스텔 배경 + 좌측정렬/들여쓰기 1
        subSpans2.forEach(span => {
          const topRow = blockStartRow + span.start;
          const endRow = blockStartRow + span.end;
          const cell = sheet.getCell(topRow, colSub2);
          const color = getSubjectColor(span.subject);

          if (span.subject) {
            const stText = getStatusText(statusMap2, span.subject);
            let stColor = 'FF059669'; // 개설: 진한 초록
            if (stText === '조정') stColor = 'FFD97706'; // 조정: 진한 주황
            else if (stText === '폐강') stColor = 'FFDC2626'; // 폐강: 진한 빨강

            cell.value = {
              richText: [
                {
                  text: span.subject + ' ',
                  font: {
                    name: '맑은 고딕',
                    size: 13,
                    bold: true,
                    color: { argb: color ? color.text : 'FF1E293B' }
                  }
                },
                {
                  text: `(${stText})`,
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

          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

          // 파스텔톤 배경색 채우기
          if (color) {
            for (let r = topRow; r <= endRow; r++) {
              const c = sheet.getCell(r, colSub2);
              c.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color.bg }
              };
              this.applyBorder(c);
            }
          } else {
            this.applyBorder(cell);
          }
        });
      });

      currentRow = blockEndRow + 1;
    });
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
