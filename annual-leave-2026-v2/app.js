/* 2026학년도 연차유급휴가 부여일수 계산기 (정적/브라우저 JS-only)
 * 기준: 2025-03-01 ~ 2026-02-28 근무실적, 2026-03-01 부여.
 * 주의: 자동분류는 보조용이며, 최종 분류는 사용자가 수정 가능.
 */
(function(){
  'use strict';

  // ===== 고정 기준 =====
  const PERIOD_START = '2025-03-01';
  const PERIOD_END   = '2026-02-28';
  const GRANT_DATE   = '2026-03-01';

  // ===== DOM =====
  const el = (id) => document.getElementById(id);

  // Calendar
  const summerStartEl = el('summerStart');
  const summerEndEl = el('summerEnd');
  const winterStartEl = el('winterStart');
  const winterEndEl = el('winterEnd');
  const excludeSaturdayEl = el('excludeSaturday');
  const btnApplyCalendar = el('btnApplyCalendar');
  const calendarSummaryEl = el('calendarSummary');

  // HR input
  const hrUploadPane = el('hrUploadPane');
  const hrManualPane = el('hrManualPane');
  const hrFileEl = el('hrFile');
  const btnLoadHr = el('btnLoadHr');
  const hrUploadStatusEl = el('hrUploadStatus');

  // Manual HR fields
  const mPersonalIdEl = el('mPersonalId');
  const mNameEl = el('mName');
  const mJobEl = el('mJob');
  const mJobGroupEl = el('mJobGroup');
  const mWorkFormEl = el('mWorkForm');
  const mWorkTypeEl = el('mWorkType');
  const mWeeklyHoursEl = el('mWeeklyHours');
  const mHireDateEl = el('mHireDate');
  const btnAddEmployee = el('btnAddEmployee');
  const hrManualStatusEl = el('hrManualStatus');

  // tables
  const tblEmployeesBody = el('tblEmployees').querySelector('tbody');

  // Work input
  const workUploadPane = el('workUploadPane');
  const workManualPane = el('workManualPane');
  const workFilesEl = el('workFiles');
  const btnLoadWork = el('btnLoadWork');
  const workUploadStatusEl = el('workUploadStatus');

  // Manual work fields
  const mEmpSelectEl = el('mEmpSelect');
  const mRecStartEl = el('mRecStart');
  const mRecEndEl = el('mRecEnd');
  const mRecTypeEl = el('mRecType');
  const mRecReasonEl = el('mRecReason');
  const mRecClassEl = el('mRecClass');
  const mRecVacationCreditEl = el('mRecVacationCredit');
  const btnAddRecord = el('btnAddRecord');

  const tblRecordsBody = el('tblRecords').querySelector('tbody');

  // Calculate/export
  const btnCalculate = el('btnCalculate');
  const btnExport = el('btnExport');
  const calcStatusEl = el('calcStatus');
  const tblResultsBody = el('tblResults').querySelector('tbody');

  // ===== 상태 =====
  const state = {
    rules: null,
    calendar: {
      summerStart: '',
      summerEnd: '',
      winterStart: '',
      winterEnd: '',
      excludeSaturday: true,
    },
    daySets: {
      fullYearDays: new Set(),   // epochDay (UTC day index)
      semesterDays: new Set(),   // fullYearDays minus vacation
      vacationDays: new Set(),
      totalFullDays: 0,
      totalSemesterDays: 0
    },
    employees: new Map(), // key -> employee
    records: [], // list of records
    results: []  // computed
  };

  // ===== 날짜 유틸(UTC day index) =====
  function ymdToEpochDay(ymd){
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ymd);
    if(!m) return null;
    const y = Number(m[1]); const mo = Number(m[2]); const d = Number(m[3]);
    return Math.floor(Date.UTC(y, mo-1, d) / 86400000);
  }
  function epochDayToYMD(n){
    return new Date(n * 86400000).toISOString().slice(0,10);
  }
  function dayOfWeekEpochDay(n){
    return new Date(n * 86400000).getUTCDay(); // 0 Sun ... 6 Sat
  }
  function clampToPeriod(ymd){
    const a = ymdToEpochDay(ymd);
    const ps = ymdToEpochDay(PERIOD_START);
    const pe = ymdToEpochDay(PERIOD_END);
    if(a === null) return null;
    if(a < ps) return epochDayToYMD(ps);
    if(a > pe) return epochDayToYMD(pe);
    return ymd;
  }
  function enumerateEpochDays(startYMD, endYMD){
    const s = ymdToEpochDay(startYMD);
    const e = ymdToEpochDay(endYMD);
    if(s === null || e === null) return [];
    const out = [];
    const step = s <= e ? 1 : -1;
    for(let n=s; step>0 ? n<=e : n>=e; n+=step){
      out.push(n);
    }
    return out;
  }

  // ===== 문자열 파서 =====
  function normalizeHeader(h){
    if(h === null || h === undefined) return '';
    return String(h).replace(/\r?\n/g,'').replace(/\s+/g,'').trim();
  }
  function parseNameAndId(raw){
    const s = String(raw ?? '').trim();
    // e.g. "고남향\r\n(K109050178)"
    const idMatch = /\((K\d+)\)/.exec(s);
    const personalId = idMatch ? idMatch[1] : '';
    const name = s.replace(/\r?\n/g,' ').replace(/\(K\d+\)/,'').trim();
    return { name, personalId };
  }
  function parseWeeklyMinutes(value){
    // "40시간00분" or number
    if(value === null || value === undefined) return null;
    if(typeof value === 'number') return Math.round(value * 60);
    const s = String(value).trim();
    const hm = /(\d+)\s*시간\s*(\d+)\s*분/.exec(s);
    if(hm) return Number(hm[1])*60 + Number(hm[2]);
    const hOnly = /(\d+)\s*시간/.exec(s);
    if(hOnly) return Number(hOnly[1])*60;
    const num = Number(s);
    if(!Number.isNaN(num)) return Math.round(num * 60);
    return null;
  }
  function parseKoreanDuration(s){
    // "6일 0시간 30분" or "0일 1시간 0분"
    const str = String(s ?? '').trim();
    const m = /(\d+)\s*일\s*(\d+)\s*시간\s*(\d+)\s*분/.exec(str);
    if(m){
      return {
        days: Number(m[1]),
        hours: Number(m[2]),
        mins: Number(m[3]),
        totalMinutes: (Number(m[1])*8 + Number(m[2]))*60 + Number(m[3]) // 기본 1일=8시간 환산(보고용)
      };
    }
    // fallback: try hours/mins only
    const hm = /(\d+)\s*시간\s*(\d+)\s*분/.exec(str);
    if(hm){
      return {days:0, hours:Number(hm[1]), mins:Number(hm[2]), totalMinutes:(Number(hm[1])*60+Number(hm[2]))};
    }
    return {days:0, hours:0, mins:0, totalMinutes:0};
  }
  function parsePeriodRange(s){
    // "2025-05-08 14:30 ~ 2025-05-08 15:30"
    const str = String(s ?? '').trim();
    const m = /(\d{4}-\d{2}-\d{2})\s*\d{2}:\d{2}\s*~\s*(\d{4}-\d{2}-\d{2})\s*\d{2}:\d{2}/.exec(str);
    if(m) return {start: m[1], end: m[2]};
    const m2 = /(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})/.exec(str);
    if(m2) return {start: m2[1], end: m2[2]};
    const m3 = /(\d{4}-\d{2}-\d{2})/.exec(str);
    if(m3) return {start: m3[1], end: m3[1]};
    return {start:'', end:''};
  }

  // ===== 근속 =====
  function completedYears(hireYMD, refYMD){
    const h = /^(\d{4})-(\d{2})-(\d{2})$/.exec(hireYMD);
    const r = /^(\d{4})-(\d{2})-(\d{2})$/.exec(refYMD);
    if(!h || !r) return 0;
    const hy = Number(h[1]), hm = Number(h[2]), hd = Number(h[3]);
    const ry = Number(r[1]), rm = Number(r[2]), rd = Number(r[3]);
    let years = ry - hy;
    if (rm < hm || (rm === hm && rd < hd)) years--;
    return Math.max(0, years);
  }

  // ===== 분류(룰 기반) =====
  function classifyRecordAuto(leaveType, reason){
    const t = String(leaveType ?? '').trim();
    const r = String(reason ?? '').trim();
    const tNorm = t.replace(/\s+/g,'');
    const rNorm = r;

    let cls = 'review';
    let needsReview = false;
    let note = '';

    if(state.rules){
      for(const m of state.rules.matchers || []){
        const typeExact = (m.type_exact || []).map(x => String(x).replace(/\s+/g,''));
        const reasonKeywords = (m.reason_keywords || []);
        const typeHit = typeExact.length ? typeExact.includes(tNorm) : false;
        const reasonHit = reasonKeywords.some(k => rNorm.includes(k));
        if(typeHit || reasonHit){
          cls = m.class || 'review';
          if(m.needs_review_if_type && m.needs_review_if_type.map(x=>String(x).replace(/\s+/g,'')).includes(tNorm)){
            needsReview = true;
            note = '상한/예외 규정 검토 필요';
          }
          break;
        }
      }
    }

    // 방학근무크레딧(상시기준 출근율 산정용)
    const vacationCredit = (state.rules?.vacation_credit_keywords || []).some(k => rNorm.includes(k));

    // 학교장 재량휴업 등은 통상 유급휴일로 많이 처리(자동은 deemed로 승격)
    const paidHoliday = (state.rules?.paid_holiday_keywords || []).some(k => rNorm.includes(k));
    if(cls === 'review' && paidHoliday){
      cls = 'deemed';
      needsReview = false;
      note = '약정 유급휴일(학교별 확인)';
    }

    if(cls === 'review' && vacationCredit){
      needsReview = true;
      note = note || '방학근무/유급처리 여부 확인';
    }

    return { cls, vacationCredit, needsReview, note };
  }

  function resolveFinalClassification(rec){
    const cls = rec.override?.cls || rec.auto.cls;
    const vacationCredit = (rec.override?.vacationCredit !== undefined && rec.override?.vacationCredit !== null)
      ? rec.override.vacationCredit
      : rec.auto.vacationCredit;
    return { cls, vacationCredit };
  }

  // ===== 캘린더(분모) 계산 =====
  function rebuildDaySets(){
    const ps = ymdToEpochDay(PERIOD_START);
    const pe = ymdToEpochDay(PERIOD_END);
    const excludeSat = !!state.calendar.excludeSaturday;

    const fullYear = new Set();
    for(let d=ps; d<=pe; d++){
      const dow = dayOfWeekEpochDay(d);
      if(excludeSat && dow === 6) continue; // 토요일 제외
      fullYear.add(d);
    }

    const vacation = new Set();
    function addVacationRange(sYMD, eYMD){
      if(!sYMD || !eYMD) return;
      const s = ymdToEpochDay(clampToPeriod(sYMD));
      const e = ymdToEpochDay(clampToPeriod(eYMD));
      if(s === null || e === null) return;
      const step = s <= e ? 1 : -1;
      for(let d=s; step>0 ? d<=e : d>=e; d+=step){
        if(fullYear.has(d)) vacation.add(d);
      }
    }
    addVacationRange(state.calendar.summerStart, state.calendar.summerEnd);
    addVacationRange(state.calendar.winterStart, state.calendar.winterEnd);

    const semester = new Set();
    for(const d of fullYear){
      if(!vacation.has(d)) semester.add(d);
    }

    state.daySets.fullYearDays = fullYear;
    state.daySets.vacationDays = vacation;
    state.daySets.semesterDays = semester;
    state.daySets.totalFullDays = fullYear.size;
    state.daySets.totalSemesterDays = semester.size;

    calendarSummaryEl.textContent =
      `연간(토 제외) 일수: ${state.daySets.totalFullDays}일 / 학기(방학 제외) 일수: ${state.daySets.totalSemesterDays}일`;
  }

  // ===== UI 모드 전환 =====
  function getSelectedRadio(name){
    const els = document.querySelectorAll(`input[name="${name}"]`);
    for(const e of els){
      if(e.checked) return e.value;
    }
    return null;
  }
  function updateMethodPanes(){
    const hrMethod = getSelectedRadio('hrMethod');
    hrUploadPane.classList.toggle('hidden', hrMethod !== 'upload');
    hrManualPane.classList.toggle('hidden', hrMethod !== 'manual');

    const workMethod = getSelectedRadio('workMethod');
    workUploadPane.classList.toggle('hidden', workMethod !== 'upload');
    workManualPane.classList.toggle('hidden', workMethod !== 'manual');
  }

  // ===== 대상자 관리 =====
  function employeeKey(personalId, name){
    if(personalId) return personalId.trim();
    return `NAME:${(name||'').trim()}`;
  }
  function upsertEmployee(emp){
    const key = employeeKey(emp.personalId, emp.name);
    emp.key = key;
    state.employees.set(key, emp);
    refreshEmployeeTable();
    refreshManualEmployeeSelect();
  }
  function removeEmployee(key){
    state.employees.delete(key);
    // orphan records remain, but we can keep for review; here we also remove records bound to that key
    state.records = state.records.filter(r => r.empKey !== key);
    refreshEmployeeTable();
    refreshRecordTable();
    refreshManualEmployeeSelect();
  }

  function refreshEmployeeTable(){
    tblEmployeesBody.innerHTML = '';
    const rows = Array.from(state.employees.values()).sort((a,b)=> (a.name||'').localeCompare(b.name||''));
    for(const emp of rows){
      const tr = document.createElement('tr');
      const serviceYears = completedYears(emp.hireDate || '', GRANT_DATE);
      emp.serviceYears = serviceYears;

      tr.innerHTML = `
        <td>${escapeHtml(emp.personalId||'')}</td>
        <td>${escapeHtml(emp.name||'')}</td>
        <td>${escapeHtml(emp.job||'')}</td>
        <td>${escapeHtml(emp.jobGroup||'')}</td>
        <td>${escapeHtml(emp.workForm||'')}</td>
        <td>${escapeHtml(emp.workType||'')}</td>
        <td>${escapeHtml(emp.hireDate||'')}</td>
        <td>${escapeHtml(String(emp.weeklyHours ?? ''))}</td>
        <td>${serviceYears}</td>
        <td><button class="danger" data-act="delEmp" data-key="${escapeAttr(emp.key)}">삭제</button></td>
      `;
      tblEmployeesBody.appendChild(tr);
    }
  }

  function refreshManualEmployeeSelect(){
    mEmpSelectEl.innerHTML = '';
    const opt0 = document.createElement('option');
    opt0.value = '';
    opt0.textContent = '대상자 선택';
    mEmpSelectEl.appendChild(opt0);

    const rows = Array.from(state.employees.values()).sort((a,b)=> (a.name||'').localeCompare(b.name||''));
    for(const emp of rows){
      const opt = document.createElement('option');
      opt.value = emp.key;
      opt.textContent = `${emp.name} (${emp.personalId || emp.workForm || '식별정보 없음'})`;
      mEmpSelectEl.appendChild(opt);
    }
  }

  // ===== 복무 기록 관리 =====
  function addRecord(rec){
    state.records.push(rec);
    refreshRecordTable();
  }

  function refreshRecordTable(){
    tblRecordsBody.innerHTML = '';
    for(const rec of state.records){
      const final = resolveFinalClassification(rec);
      const tr = document.createElement('tr');

      const emp = state.employees.get(rec.empKey);
      const displayName = emp?.name || rec.name || '';
      const displayId = emp?.personalId || rec.personalId || '';

      const period = rec.startDate && rec.endDate ? `${rec.startDate} ~ ${rec.endDate}` : '';

      const autoLabel = labelClass(rec.auto.cls, rec.auto.needsReview);
      const finalLabel = labelClass(final.cls, rec.auto.needsReview);

      tr.innerHTML = `
        <td>${escapeHtml(displayName)}</td>
        <td>${escapeHtml(displayId)}</td>
        <td>${escapeHtml(period)}</td>
        <td>${escapeHtml(rec.leaveType||'')}</td>
        <td>${escapeHtml(rec.reason||'')}</td>
        <td>${escapeHtml(autoLabel)}</td>
        <td>
          <select data-act="setCls" data-id="${rec.id}">
            <option value="auto"${(!rec.override || rec.override.cls===null || rec.override.cls===undefined) ? ' selected' : ''}>자동</option>
            <option value="deemed"${(rec.override?.cls==='deemed') ? ' selected' : ''}>출근간주</option>
            <option value="excluded"${(rec.override?.cls==='excluded') ? ' selected' : ''}>산정제외</option>
            <option value="absence"${(rec.override?.cls==='absence') ? ' selected' : ''}>결근성</option>
            <option value="review"${(rec.override?.cls==='review') ? ' selected' : ''}>검토</option>
          </select>
        </td>
        <td>
          <label class="checkbox" style="margin:0;">
            <input type="checkbox" data-act="setVac" data-id="${rec.id}" ${(final.vacationCredit ? 'checked' : '')}/>
            <span class="subtle">적용</span>
          </label>
        </td>
        <td>${escapeHtml(rec.approvalStatus || '')}</td>
        <td><button class="danger" data-act="delRec" data-id="${rec.id}">삭제</button></td>
      `;
      tblRecordsBody.appendChild(tr);
    }
  }

  function labelClass(cls, needsReview){
    const base = (cls === 'deemed') ? '출근간주'
      : (cls === 'excluded') ? '산정제외'
      : (cls === 'absence') ? '결근성'
      : '검토';
    return needsReview ? `${base} (확인필요)` : base;
  }

  // ===== 출근율/연차 계산 =====
  function intersectionSize(setA, setB){
    let n=0;
    for(const v of setA){
      if(setB.has(v)) n++;
    }
    return n;
  }

  function recordCoveredDays(rec){
    const days = [];
    if(!rec.startDate || !rec.endDate) return days;
    const start = clampToPeriod(rec.startDate);
    const end = clampToPeriod(rec.endDate);
    const all = enumerateEpochDays(start, end);
    for(const d of all){
      if(state.daySets.fullYearDays.has(d)) days.push(d); // 분모 정의와 동일하게 토 제외 반영
    }
    return days;
  }

  function computeEmployeeMetrics(emp){
    const isSpecial = String(emp.jobGroup||'').includes('특수운영');
    const workForm = String(emp.workForm||'');
    const isRegular = workForm.includes('상시');
    const isVacationOff = workForm.includes('방학중비근무');
    const isUltraShort = workForm.includes('초단시간');

    const fullYear = state.daySets.fullYearDays;
    const semester = state.daySets.semesterDays;

    // 수기 입력은 "예외사항만 등록" → 미등록 날짜는 정상 근무로 간주
    const excludedDays = new Set();
    const absenceDays = new Set();
    const vacationCreditDays = new Set();

    const recs = state.records.filter(r => r.empKey === emp.key);
    for(const rec of recs){
      const final = resolveFinalClassification(rec);
      const days = recordCoveredDays(rec);
      if(final.cls === 'excluded'){
        for(const d of days) excludedDays.add(d);
      }else if(final.cls === 'absence'){
        for(const d of days) absenceDays.add(d);
      }

      if(final.vacationCredit){
        for(const d of days) vacationCreditDays.add(d);
      }
    }

    // 공통: 연간 분모
    const totalFull = state.daySets.totalFullDays;

    // 상시 기준 (raw/recalc)
    const excludedFull = intersectionSize(excludedDays, fullYear);
    const absenceFull = intersectionSize(absenceDays, fullYear);
    const attendedFull_regular = Math.max(0, totalFull - excludedFull - absenceFull);
    const rateFull_raw = safeRate(attendedFull_regular, totalFull);
    const rateFull_recalc = safeRate(attendedFull_regular, Math.max(0, totalFull - excludedFull));

    // 비상시 기준(학기) raw/recalc
    const totalSem = state.daySets.totalSemesterDays;
    const excludedSem = intersectionSize(excludedDays, semester);
    const absenceSem = intersectionSize(absenceDays, semester);
    const attendedSem = Math.max(0, totalSem - excludedSem - absenceSem);
    const rateSem_raw = safeRate(attendedSem, totalSem);
    const rateSem_recalc = safeRate(attendedSem, Math.max(0, totalSem - excludedSem));

    // 비상시 상시기준(연간) raw: (학기 출근분 + 방학근무크레딧) / 연간분모
    // vacationCreditDays 는 학기/방학 구분 없이 잡히므로, "학기 외"만 카운트
    let creditOutsideSemester = 0;
    for(const d of vacationCreditDays){
      if(fullYear.has(d) && !semester.has(d)) creditOutsideSemester++;
    }
    const attendedFull_vacationOff = attendedSem + creditOutsideSemester;
    const rateFull_vacationOff = safeRate(attendedFull_vacationOff, totalFull);

    // 근속/가산
    const serviceYears = completedYears(emp.hireDate || '', GRANT_DATE);
    const addDays = Math.floor(Math.max(0, serviceYears - 1) / 2);
    const baseRegular = 15;
    const baseVacationEdu = 12;
    const baseVacationSpecial = 11;

    // 기준 출근율 경로
    let route = '';
    let baseDays = 0;
    let grantedDays = 0;
    let grantedHours = 0;
    let basePlusAdd = 0;

    // 초단시간은 대상 제외 안내(근기법 적용 제외 가능성이 높음)
    if(isUltraShort){
      route = '대상 제외(초단시간)';
      return buildResult(emp, {
        isSpecial, isRegular, isVacationOff, serviceYears,
        rateSem_raw, rateSem_recalc, rateFull_raw: rateFull_vacationOff, rateFull_recalc,
        baseDays: 0, addDays: 0, basePlusAdd: 0,
        grantedDays: 0, grantedHours: 0, route,
        excludedFull, absenceFull, excludedSem, absenceSem, creditOutsideSemester
      });
    }

    if(serviceYears < 1){
      // 1년 미만: 1개월 개근 시 1일(학교 방식은 실제로 월별 발생/사용기한 관리가 필요하지만, 여기서는 총합만 제공)
      // 개근월수는 "결근성(absence)"가 있는 달은 제외. 산정제외 기간은 스케줄에서 빠진 것으로 취급(개근 깨지지 않음).
      const maxDays = isVacationOff ? 9 : 11;
      const months = countPerfectMonths(emp, absenceDays, excludedDays, isVacationOff);
      grantedDays = Math.min(maxDays, months);
      route = '1년 미만: 개근월수(1개월 1일)';
      return buildResult(emp, {
        isSpecial, isRegular, isVacationOff, serviceYears,
        rateSem_raw, rateSem_recalc, rateFull_raw: (isVacationOff ? rateFull_vacationOff : rateFull_raw), rateFull_recalc,
        baseDays: 0, addDays: 0, basePlusAdd: 0,
        grantedDays, grantedHours: 0, route,
        excludedFull, absenceFull, excludedSem, absenceSem, creditOutsideSemester
      });
    }

    // 1년 이상: 기본일수 결정
    if(isRegular){
      baseDays = baseRegular;
    }else if(isVacationOff){
      // 비상시: 기본 12/11, 단 상시기준 출근율(연간) 80% 이상이면 15로 전환
      if(isSpecial){
        baseDays = (rateFull_vacationOff >= 0.8) ? 15 : baseVacationSpecial;
      }else{
        baseDays = (rateFull_vacationOff >= 0.8) ? 15 : baseVacationEdu;
      }
    }else{
      // 근무형태 미확정: 상시로 가정(업무상 안전)
      baseDays = baseRegular;
    }

    basePlusAdd = Math.min(25, baseDays + addDays);

    // 출근율 분기
    if(isRegular){
      if(rateFull_raw >= 0.8){
        route = '정상부여(상시 80% 이상)';
        ({days: grantedDays, hours: grantedHours} = toDayHour(basePlusAdd));
      }else if(rateFull_recalc >= 0.8){
        route = '비례부여(상시 재산정 80% 이상)';
        const ratio = safeRate(Math.max(0, totalFull - excludedFull), totalFull);
        ({days: grantedDays, hours: grantedHours} = prorateToDayHour(basePlusAdd, ratio));
      }else{
        route = '개근월수(상시 80% 미달)';
        const months = countPerfectMonths(emp, absenceDays, excludedDays, false);
        ({days: grantedDays, hours: grantedHours} = toDayHour(months)); // 개근월수 = 일수
      }
    }else if(isVacationOff){
      if(rateSem_raw >= 0.8){
        route = '정상부여(비상시 학기 80% 이상)';
        ({days: grantedDays, hours: grantedHours} = toDayHour(basePlusAdd));
      }else if(rateSem_recalc >= 0.8){
        route = '비례부여(비상시 재산정 80% 이상)';
        // 사용자 정리 기준: 비례식의 분모는 2025학년도 총 일수(연간)로 고정
        const ratio = safeRate(Math.max(0, totalFull - excludedFull), totalFull);
        ({days: grantedDays, hours: grantedHours} = prorateToDayHour(basePlusAdd, ratio));
      }else{
        route = '개근월수(비상시 80% 미달)';
        const months = countPerfectMonths(emp, absenceDays, excludedDays, true);
        ({days: grantedDays, hours: grantedHours} = toDayHour(months));
      }
    }else{
      // 알 수 없음 → 상시 경로로 처리
      if(rateFull_raw >= 0.8){
        route = '정상부여(근무형태 미확정: 상시 가정)';
        ({days: grantedDays, hours: grantedHours} = toDayHour(basePlusAdd));
      }else if(rateFull_recalc >= 0.8){
        route = '비례부여(근무형태 미확정: 상시 가정)';
        const ratio = safeRate(Math.max(0, totalFull - excludedFull), totalFull);
        ({days: grantedDays, hours: grantedHours} = prorateToDayHour(basePlusAdd, ratio));
      }else{
        route = '개근월수(근무형태 미확정: 상시 가정)';
        const months = countPerfectMonths(emp, absenceDays, excludedDays, false);
        ({days: grantedDays, hours: grantedHours} = toDayHour(months));
      }
    }

    return buildResult(emp, {
      isSpecial, isRegular, isVacationOff, serviceYears,
      rateSem_raw, rateSem_recalc, rateFull_raw: (isVacationOff ? rateFull_vacationOff : rateFull_raw), rateFull_recalc,
      baseDays, addDays, basePlusAdd,
      grantedDays, grantedHours, route,
      excludedFull, absenceFull, excludedSem, absenceSem, creditOutsideSemester
    });
  }

  function safeRate(num, den){
    if(!den || den <= 0) return 0;
    return num / den;
  }

  function toDayHour(daysFloat){
    // daysFloat는 정수 기대. 소수면 시간 환산
    const d = Math.floor(daysFloat);
    const frac = daysFloat - d;
    const hours = Math.round(frac * 8);
    if(hours >= 8) return {days: d+1, hours: 0};
    return {days: d, hours};
  }

  function prorateToDayHour(basePlusAdd, ratio){
    // 소수 둘째 자리에서 반올림 → 0.1일 단위로 표현 후 시간 환산
    const raw = basePlusAdd * ratio;
    const rounded1 = Math.round(raw * 10) / 10; // 소수 첫째 자리까지
    const d = Math.floor(rounded1);
    let hours = Math.round((rounded1 - d) * 8);
    let days = d;
    if(hours >= 8){
      days += 1;
      hours = 0;
    }
    return {days, hours};
  }

  function countPerfectMonths(emp, absenceDays, excludedDays, isVacationOff){
    // 개근월수(1개월 개근 시 1일): 결근성(absence)이 있는 달은 제외.
    // 산정제외(excluded)는 스케줄에서 제거되어 개근을 깨지 않는 것으로 처리(보수적).
    // 비상시 특례: 여름방학(7~8) 1개월, 겨울방학(12~2) 1개월로 간주.
    const fullYear = state.daySets.fullYearDays;
    const semester = state.daySets.semesterDays;

    // 근무 개시일이 기간 중간이면 그 이전 달은 계산 대상에서 제외(근무 전이므로)
    const hire = ymdToEpochDay(emp.hireDate || PERIOD_START) ?? ymdToEpochDay(PERIOD_START);
    const ps = ymdToEpochDay(PERIOD_START);

    const startDay = Math.max(ps, hire);

    function monthKey(epochDay){
      const ymd = epochDayToYMD(epochDay);
      return ymd.slice(0,7); // YYYY-MM
    }

    // 그룹 정의
    const groups = [];
    if(isVacationOff){
      // 3,4,5,6,9,10,11 개별 + 7~8 여름 + 12~2 겨울
      groups.push({label:'2025-03', months:['2025-03']});
      groups.push({label:'2025-04', months:['2025-04']});
      groups.push({label:'2025-05', months:['2025-05']});
      groups.push({label:'2025-06', months:['2025-06']});
      groups.push({label:'여름방학(7~8)', months:['2025-07','2025-08']});
      groups.push({label:'2025-09', months:['2025-09']});
      groups.push({label:'2025-10', months:['2025-10']});
      groups.push({label:'2025-11', months:['2025-11']});
      groups.push({label:'겨울방학(12~2)', months:['2025-12','2026-01','2026-02']});
    }else{
      // 12개월 개별
      const months = ['2025-03','2025-04','2025-05','2025-06','2025-07','2025-08','2025-09','2025-10','2025-11','2025-12','2026-01','2026-02'];
      for(const m of months) groups.push({label:m, months:[m]});
    }

    // 스케줄: 상시는 fullYear, 비상시는 semester(방학 제외)로 보는 것이 원칙이나,
    // 개근월수는 "근로제공 정지기간(방학)" 특례가 존재 → 위 그룹화로 처리.
    // 여기서는 결근성 여부만 본다(미등록 날짜는 근무로 간주).
    let count = 0;

    for(const g of groups){
      // 그룹에 해당하는 "스케줄 일"을 산정
      let scheduled = 0;
      let hasAbsence = false;

      for(const d of fullYear){
        if(d < startDay) continue;
        const mk = monthKey(d);
        if(!g.months.includes(mk)) continue;

        // 비상시의 경우 방학 달은 원칙적으로 스케줄이 없으나, 그룹 자체를 1개월로 간주하므로
        // semesterDays에 포함되는 날짜만 스케줄로 인정(방학기간은 스케줄 0)
        if(isVacationOff){
          // summer/winter 그룹에서 semesterDays가 0이면 scheduled=0 → 그 달은 개근으로 세지지 않음
          if(!semester.has(d)) continue;
        }

        // 산정제외는 스케줄에서 제거
        if(excludedDays.has(d)) continue;

        scheduled++;
        if(absenceDays.has(d)) hasAbsence = true;
      }

      if(scheduled > 0 && !hasAbsence) count++;
    }

    return count;
  }

  function buildResult(emp, m){
    return {
      empKey: emp.key,
      name: emp.name || '',
      personalId: emp.personalId || '',
      job: emp.job || '',
      jobGroup: emp.jobGroup || '',
      workForm: emp.workForm || '',
      workType: emp.workType || '',
      hireDate: emp.hireDate || '',
      weeklyHours: emp.weeklyHours ?? null,

      serviceYears: m.serviceYears,
      isRegular: m.isRegular,
      isVacationOff: m.isVacationOff,
      isSpecial: m.isSpecial,

      rateSem_raw: m.rateSem_raw,
      rateFull_raw: m.rateFull_raw,
      rateSem_recalc: m.rateSem_recalc,
      rateFull_recalc: m.rateFull_recalc,

      baseDays: m.baseDays,
      addDays: m.addDays,
      basePlusAdd: m.basePlusAdd,

      grantedDays: m.grantedDays,
      grantedHours: m.grantedHours,
      route: m.route,

      debug: {
        excludedFull: m.excludedFull,
        absenceFull: m.absenceFull,
        excludedSem: m.excludedSem,
        absenceSem: m.absenceSem,
        creditOutsideSemester: m.creditOutsideSemester
      }
    };
  }

  function renderResults(results){
    tblResultsBody.innerHTML = '';
    for(const r of results){
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(r.name)}</td>
        <td>${escapeHtml(r.personalId)}</td>
        <td>${escapeHtml(r.job)}</td>
        <td>${escapeHtml(r.workForm)}</td>
        <td>${escapeHtml(r.jobGroup)}</td>
        <td>${r.serviceYears}</td>
        <td>${r.isVacationOff ? pct(r.rateSem_raw) : '-'}</td>
        <td>${pct(r.rateFull_raw)}</td>
        <td>${r.isVacationOff ? pct(r.rateSem_recalc) : '-'}</td>
        <td>${escapeHtml(r.route)}</td>
        <td>${(r.baseDays || 0)} / ${(r.addDays || 0)}</td>
        <td><strong>${r.grantedDays}일 ${r.grantedHours}시간</strong></td>
      `;
      tblResultsBody.appendChild(tr);
    }
  }

  function pct(x){
    if(x === null || x === undefined) return '-';
    return (x*100).toFixed(1) + '%';
  }

  // ===== 엑셀/CSV 로딩 =====
  async function readFileAsArrayBuffer(file){
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result);
      fr.onerror = () => reject(fr.error);
      fr.readAsArrayBuffer(file);
    });
  }

  async function loadWorkbookFromFile(file){
    if(typeof XLSX === 'undefined'){
      throw new Error('XLSX 라이브러리가 로딩되지 않았습니다. (lib/xlsx.full.min.js 또는 CDN 필요)');
    }
    const buf = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(buf, {type:'array', cellDates:true});
    return wb;
  }

  function sheetToObjects(ws){
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''});
    if(!rows.length) return [];
    const headers = rows[0].map(normalizeHeader);
    const out = [];
    for(let i=1; i<rows.length; i++){
      const row = rows[i];
      if(row.every(v => String(v||'').trim()==='')) continue;
      const obj = {};
      for(let c=0; c<headers.length; c++){
        obj[headers[c]] = row[c];
      }
      out.push(obj);
    }
    return out;
  }

  async function loadHrFromFile(){
    const file = hrFileEl.files?.[0];
    if(!file){
      hrUploadStatusEl.textContent = '인사기록 파일을 선택하세요.';
      return;
    }
    hrUploadStatusEl.textContent = '불러오는 중...';

    try{
      const wb = await loadWorkbookFromFile(file);
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const items = sheetToObjects(ws);

      let count = 0;
      for(const it of items){
        // HR 헤더(예시): 개인번호, 성명, 직종, 직종구분, 근무시작일, 근무유형, 근무형태, 주소정근로시간
        const personalId = String(it['개인번호'] || '').trim();
        const name = String(it['성명'] || '').trim();
        if(!name && !personalId) continue;

        const hireDate = normalizeDate(it['근무시작일'] || it['근무시작'] || it['근무시작일자'] || it['근무시작일(YYYY-MM-DD)']);
        const job = String(it['직종'] || '').trim();
        const jobGroup = String(it['직종구분'] || '').trim();
        const workType = String(it['근무유형'] || '').trim();
        const workForm = String(it['근무형태'] || '').trim();

        const weeklyMin = parseWeeklyMinutes(it['주소정근로시간'] || it['주당근무시간'] || '');
        const weeklyHours = weeklyMin ? (weeklyMin / 60) : null;

        upsertEmployee({
          personalId, name, job, jobGroup, workType, workForm,
          hireDate: hireDate || '',
          weeklyHours: weeklyHours ?? 40
        });
        count++;
      }

      hrUploadStatusEl.textContent = `인사기록 로딩 완료: ${count}명`;
    }catch(err){
      console.error(err);
      hrUploadStatusEl.textContent = `인사기록 로딩 실패: ${err.message || err}`;
    }
  }

  function normalizeDate(v){
    if(!v) return '';
    // XLSX may give Date object or string
    if(v instanceof Date){
      const y = v.getFullYear();
      const m = String(v.getMonth()+1).padStart(2,'0');
      const d = String(v.getDate()).padStart(2,'0');
      return `${y}-${m}-${d}`;
    }
    const s = String(v).trim();
    const m = /(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if(m) return `${m[1]}-${m[2]}-${m[3]}`;
    // fallback for 2023. 3. 1.
    const m2 = /(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/.exec(s);
    if(m2){
      const y=m2[1], mo=String(m2[2]).padStart(2,'0'), d=String(m2[3]).padStart(2,'0');
      return `${y}-${mo}-${d}`;
    }
    return '';
  }

  async function loadWorkFromFiles(){
    const files = Array.from(workFilesEl.files || []);
    if(!files.length){
      workUploadStatusEl.textContent = '근무상황 파일을 선택하세요.';
      return;
    }
    workUploadStatusEl.textContent = '불러오는 중...';

    let totalRecs = 0;
    try{
      for(const file of files){
        const wb = await loadWorkbookFromFile(file);
        const sheetName = wb.SheetNames.find(n => String(n).includes('근무상황')) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const items = sheetToObjects(ws);

        for(const it of items){
          const nameRaw = it[normalizeHeader('성명')] || it['성명'] || it['성명(K개인번호)'] || '';
          const {name, personalId} = parseNameAndId(nameRaw);
          if(!name && !personalId) continue;

          // skip total row
          if(String(nameRaw).includes('총계')) continue;

          const empKeyGuess = employeeKey(personalId, name);
          if(!state.employees.has(empKeyGuess)){
            // 인사파일이 없을 수도 있으므로 placeholder 생성
            upsertEmployee({
              personalId,
              name,
              job: String(it[normalizeHeader('직급(직종)')] || it['직급(직종)'] || '').trim(),
              jobGroup: '미확인',
              workType: '',
              workForm: '',
              hireDate: '',
              weeklyHours: 40
            });
          }

          const period = parsePeriodRange(it[normalizeHeader('기간')] || it['기간'] || '');
          const leaveType = String(it[normalizeHeader('종별')] || it['종별'] || '').trim();
          const reason = String(it[normalizeHeader('사유또는용무')] || it['사유또는용무'] || it['사유또는용무'] || it['사유또는용무'] || it['사유 또는 용무'] || '').trim();
          const approvalStatus = String(it[normalizeHeader('결재상태')] || it['결재상태'] || '').trim();

          const auto = classifyRecordAuto(leaveType, reason);
          const rec = {
            id: cryptoRandomId(),
            empKey: empKeyGuess,
            name,
            personalId,
            startDate: period.start,
            endDate: period.end,
            leaveType,
            reason,
            approvalStatus,
            auto,
            override: {
              cls: null,
              vacationCredit: null
            },
            source: file.name
          };
          addRecord(rec);
          totalRecs++;
        }
      }

      workUploadStatusEl.textContent = `근무상황 로딩 완료: ${files.length}개 파일 / ${totalRecs}건`;
    }catch(err){
      console.error(err);
      workUploadStatusEl.textContent = `근무상황 로딩 실패: ${err.message || err}`;
    }
  }

  // ===== 수기 입력 핸들러 =====
  function addEmployeeManual(){
    const name = String(mNameEl.value || '').trim();
    const personalId = String(mPersonalIdEl.value || '').trim();
    const job = String(mJobEl.value || '').trim();
    const jobGroup = String(mJobGroupEl.value || '').trim();
    const workForm = String(mWorkFormEl.value || '').trim();
    const workType = String(mWorkTypeEl.value || '').trim();
    const weeklyHours = Number(mWeeklyHoursEl.value || 0);
    const hireDate = String(mHireDateEl.value || '').trim();

    if(!name){
      hrManualStatusEl.textContent = '성명은 필수입니다.';
      return;
    }
    if(!hireDate){
      hrManualStatusEl.textContent = '최초임용일(근무시작일)은 필수입니다.';
      return;
    }

    upsertEmployee({
      personalId, name, job, jobGroup, workForm, workType, hireDate,
      weeklyHours: weeklyHours || 40
    });

    hrManualStatusEl.textContent = `대상자 추가 완료: ${name}`;
    // reset minimal
    mPersonalIdEl.value = '';
    mNameEl.value = '';
    mJobEl.value = '';
  }

  function addRecordManual(){
    const empKey = String(mEmpSelectEl.value || '').trim();
    if(!empKey){
      alert('대상자를 선택하세요.');
      return;
    }
    const emp = state.employees.get(empKey);
    if(!emp){
      alert('대상자 정보를 찾을 수 없습니다.');
      return;
    }

    const startDate = String(mRecStartEl.value || '').trim();
    const endDate = String(mRecEndEl.value || '').trim() || startDate;
    const leaveType = String(mRecTypeEl.value || '').trim();
    const reason = String(mRecReasonEl.value || '').trim();
    const classSel = String(mRecClassEl.value || 'auto').trim();
    const vacCredit = !!mRecVacationCreditEl.checked;

    if(!startDate || !endDate || !leaveType){
      alert('기간(시작/종료)과 종별은 필수입니다.');
      return;
    }

    const auto = classifyRecordAuto(leaveType, reason);
    const override = {cls: null, vacationCredit: null};
    if(classSel !== 'auto') override.cls = classSel;
    override.vacationCredit = vacCredit;

    addRecord({
      id: cryptoRandomId(),
      empKey: empKey,
      name: emp.name,
      personalId: emp.personalId,
      startDate,
      endDate,
      leaveType,
      reason,
      approvalStatus: '수기입력',
      auto,
      override,
      source: 'manual'
    });

    // keep values for quick repeated entry; clear minimal
    mRecTypeEl.value = '';
    mRecReasonEl.value = '';
    mRecVacationCreditEl.checked = false;
  }

  // ===== 계산/엑셀 출력 =====
  function calculateAll(){
    rebuildDaySets(); // ensure
    const results = [];
    for(const emp of state.employees.values()){
      results.push(computeEmployeeMetrics(emp));
    }
    state.results = results;
    renderResults(results);
    calcStatusEl.textContent = `계산 완료: ${results.length}명 (연간 ${state.daySets.totalFullDays}일 / 학기 ${state.daySets.totalSemesterDays}일)`;
  }

  function exportXlsx(){
    if(typeof XLSX === 'undefined'){
      alert('XLSX 라이브러리가 없어 엑셀 다운로드를 생성할 수 없습니다. (lib/xlsx.full.min.js 또는 CDN 필요)');
      return;
    }

    const wb = XLSX.utils.book_new();

    // Sheet 1: results
    const sheet1 = [
      ['성명','개인번호','직종','직종구분','근무형태','근무유형','근무시작일','근속(완료연수)',
       '출근율(비상시)','출근율(상시)','출근율(비상시 재산정)','산정경로','기본','가산','최종부여(일)','최종부여(시간)',
       '연간일수(토제외)','학기일수(방학제외)','제외(연간)','결근성(연간)','제외(학기)','결근성(학기)','방학근무크레딧(학기외)']
    ];
    for(const r of state.results){
      sheet1.push([
        r.name, r.personalId, r.job, r.jobGroup, r.workForm, r.workType, r.hireDate, r.serviceYears,
        r.isVacationOff ? pctNum(r.rateSem_raw) : '',
        pctNum(r.rateFull_raw),
        r.isVacationOff ? pctNum(r.rateSem_recalc) : '',
        r.route,
        r.baseDays, r.addDays, r.grantedDays, r.grantedHours,
        state.daySets.totalFullDays,
        state.daySets.totalSemesterDays,
        r.debug.excludedFull, r.debug.absenceFull, r.debug.excludedSem, r.debug.absenceSem,
        r.debug.creditOutsideSemester
      ]);
    }
    const ws1 = XLSX.utils.aoa_to_sheet(sheet1);
    XLSX.utils.book_append_sheet(wb, ws1, '연차부여결과');

    // Sheet 2: records (full)
    const sheet2 = [
      ['성명','개인번호','기간시작','기간종료','종별','사유','자동분류','자동확인필요','최종분류','방학근무크레딧','결재상태','출처파일']
    ];
    for(const rec of state.records){
      const emp = state.employees.get(rec.empKey);
      const displayName = emp?.name || rec.name || '';
      const displayId = emp?.personalId || rec.personalId || '';
      const final = resolveFinalClassification(rec);
      sheet2.push([
        displayName, displayId, rec.startDate, rec.endDate, rec.leaveType, rec.reason,
        rec.auto.cls, rec.auto.needsReview ? 'Y' : '',
        final.cls, final.vacationCredit ? 'Y' : '',
        rec.approvalStatus, rec.source
      ]);
    }
    const ws2 = XLSX.utils.aoa_to_sheet(sheet2);
    XLSX.utils.book_append_sheet(wb, ws2, '복무취합전체');

    const filename = `annual-leave-2026-results_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  function pctNum(x){
    if(x === null || x === undefined) return '';
    return (x*100).toFixed(1) + '%';
  }

  // ===== 보안/표시 =====
  function escapeHtml(s){
    return String(s ?? '')
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#39;');
  }
  function escapeAttr(s){ return escapeHtml(s).replace(/"/g,'&quot;'); }

  function cryptoRandomId(){
    // simple unique id
    if(window.crypto && crypto.getRandomValues){
      const arr = new Uint32Array(2);
      crypto.getRandomValues(arr);
      return 'R' + arr[0].toString(16) + arr[1].toString(16);
    }
    return 'R' + Math.random().toString(16).slice(2) + Date.now().toString(16);
  }

  // ===== 이벤트 바인딩 =====
  function bindEvents(){
    // method toggle
    document.querySelectorAll('input[name="hrMethod"]').forEach(r => r.addEventListener('change', updateMethodPanes));
    document.querySelectorAll('input[name="workMethod"]').forEach(r => r.addEventListener('change', updateMethodPanes));

    btnApplyCalendar.addEventListener('click', () => {
      state.calendar.summerStart = summerStartEl.value;
      state.calendar.summerEnd = summerEndEl.value;
      state.calendar.winterStart = winterStartEl.value;
      state.calendar.winterEnd = winterEndEl.value;
      state.calendar.excludeSaturday = !!excludeSaturdayEl.checked;
      rebuildDaySets();
    });

    btnLoadHr.addEventListener('click', loadHrFromFile);
    btnAddEmployee.addEventListener('click', addEmployeeManual);

    btnLoadWork.addEventListener('click', loadWorkFromFiles);
    btnAddRecord.addEventListener('click', addRecordManual);

    btnCalculate.addEventListener('click', calculateAll);
    btnExport.addEventListener('click', exportXlsx);

    // table actions (event delegation)
    tblEmployeesBody.addEventListener('click', (e) => {
      const t = e.target;
      if(!(t instanceof HTMLElement)) return;
      const act = t.getAttribute('data-act');
      if(act === 'delEmp'){
        const key = t.getAttribute('data-key');
        if(key && confirm('대상자를 삭제하면 관련 복무 기록도 삭제됩니다. 진행하시겠습니까?')){
          removeEmployee(key);
        }
      }
    });

    tblRecordsBody.addEventListener('click', (e) => {
      const t = e.target;
      if(!(t instanceof HTMLElement)) return;
      const act = t.getAttribute('data-act');
      if(act === 'delRec'){
        const id = t.getAttribute('data-id');
        if(!id) return;
        state.records = state.records.filter(r => r.id !== id);
        refreshRecordTable();
      }
    });

    tblRecordsBody.addEventListener('change', (e) => {
      const t = e.target;
      if(!(t instanceof HTMLElement)) return;
      const act = t.getAttribute('data-act');
      const id = t.getAttribute('data-id');
      if(!id) return;

      const rec = state.records.find(r => r.id === id);
      if(!rec) return;

      if(act === 'setCls' && t instanceof HTMLSelectElement){
        const v = t.value;
        if(v === 'auto'){
          rec.override = rec.override || {};
          rec.override.cls = null;
        }else{
          rec.override = rec.override || {};
          rec.override.cls = v;
        }
      }else if(act === 'setVac' && t instanceof HTMLInputElement){
        rec.override = rec.override || {};
        rec.override.vacationCredit = t.checked;
      }
    });
  }

  // ===== 초기화 =====
  async function init(){
    updateMethodPanes();

    // default dates (typical)
    if(!summerStartEl.value) summerStartEl.value = '2025-07-19';
    if(!summerEndEl.value) summerEndEl.value = '2025-08-17';
    if(!winterStartEl.value) winterStartEl.value = '2025-12-27';
    if(!winterEndEl.value) winterEndEl.value = '2026-02-28';

    // load rules
    try{
      const resp = await fetch('rules_2026.json', {cache:'no-store'});
      state.rules = await resp.json();
    }catch(err){
      console.warn('rules_2026.json 로딩 실패', err);
      state.rules = null;
    }

    // initial calendar build
    state.calendar.summerStart = summerStartEl.value;
    state.calendar.summerEnd = summerEndEl.value;
    state.calendar.winterStart = winterStartEl.value;
    state.calendar.winterEnd = winterEndEl.value;
    state.calendar.excludeSaturday = !!excludeSaturdayEl.checked;
    rebuildDaySets();

    bindEvents();
    refreshEmployeeTable();
    refreshRecordTable();
    refreshManualEmployeeSelect();

    // XLSX ready event (when loaded from CDN after init)
    document.addEventListener('xlsx-ready', () => {
      hrUploadStatusEl.textContent = '엑셀 라이브러리 로딩 완료(CDN).';
      workUploadStatusEl.textContent = '엑셀 라이브러리 로딩 완료(CDN).';
    });
  }

  init();

})();
