/* /after-school-monthly-wage/script.js */
/* eslint-disable no-alert */
(function () {
  // ----------------------------
  // 0) 라이브러리 준비(PDF.js)
  // ----------------------------
  const pdfjsLib = window.pdfjsLib;
  if (!pdfjsLib) {
    alert("PDF.js 로딩 실패. 네트워크 상태를 확인하세요.");
    return;
  }
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.js";

  const $ = (id) => document.getElementById(id);
  const fmt = (n) => (Number.isFinite(n) ? Math.round(n).toLocaleString("ko-KR") : "0");

  // ----------------------------
  // 1) 필요경비율(기간별) - 데이터로 유지
  // ----------------------------
  const EXPENSE_PERIODS = [
    { label: "22년7월~23년6월", start: "2022-07-01", end: "2023-06-30", rate: 0.235 },
    { label: "23년7월~24년6월", start: "2023-07-01", end: "2024-06-30", rate: 0.165 },
    { label: "24년7월~25년6월", start: "2024-07-01", end: "2025-06-30", rate: 0.165 },
    { label: "25년7월~26년6월", start: "2025-07-01", end: "2026-06-30", rate: 0.149 },
  ];

  // ----------------------------
  // 2) UI 초기화
  // ----------------------------
  const periodSel = $("periodSel");
  EXPENSE_PERIODS.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p.label;
    opt.textContent = `${p.label} (공제율 ${Math.round(p.rate * 1000) / 10}%)`;
    periodSel.appendChild(opt);
  });
  periodSel.value = EXPENSE_PERIODS[EXPENSE_PERIODS.length - 1].label;

  const dropZone = $("dropZone");
  const pdfInput = $("pdfInput");
  const fileNames = $("fileNames");
  const status = $("status");

  const state = {
    files: [],
    rawItems: [],
    summary: [],
    uncertain: [],
    workbook: null,
  };

  function setFiles(files) {
    state.files = Array.from(files || []).filter((f) => f.type === "application/pdf");
    fileNames.textContent = state.files.length
      ? state.files.map((f) => f.name).join(", ")
      : "선택된 파일 없음";
    $("btnDownload").disabled = true;
  }

  dropZone.addEventListener("click", () => pdfInput.click());
  pdfInput.addEventListener("change", (e) => setFiles(e.target.files));

  ["dragenter", "dragover"].forEach((evt) => {
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((evt) => {
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("dragover");
    });
  });
  dropZone.addEventListener("drop", (e) => setFiles(e.dataTransfer.files));

  function getSelectedExpenseRate() {
    const custom = parseFloat(($("customRate").value || "").trim());
    if (Number.isFinite(custom) && custom > 0) return custom / 100;

    const label = $("periodSel").value;
    const p = EXPENSE_PERIODS.find((x) => x.label === label);
    return p ? p.rate : EXPENSE_PERIODS[EXPENSE_PERIODS.length - 1].rate;
  }

  // ----------------------------
  // 3) PDF 텍스트 추출 (로컬 처리)
  //    - arrayBuffer -> pdfjsLib.getDocument() 패턴
  // ----------------------------
  async function extractTextFromPdf(file) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(buf).promise; // 김동규 코드와 동일 패턴7

    const pages = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      // MVP: 아이템 str join (표 좌표 복원은 2차에서)
      const raw = tc.items.map((it) => it.str).join(" ");
      pages.push({ page: p, text: normalize(raw) });
    }
    return pages;
  }

  function normalize(s) {
    return (s || "")
      .replace(/\u00a0/g, " ")
      .replace(/[ \t]+/g, " ")
      .trim();
  }

  // 품목내역 라인 후보 분절: " 12 " 같은 번호 패턴 앞에 개행 삽입
  function injectRowBreaks(text) {
    return text.replace(/(\s)(\d{1,3})\s+/g, "\n$2 ");
  }

  // "번호 내용 수량 단가 금액" 최대한 복원
  const ROW_RE = /^\s*(\d+)\s+(.+?)\s+(\d+(?:\.\d+)?)\s+([\d,]+)\s+([\d,]+)\s*$/;

  function parseRow(line) {
    const m = line.match(ROW_RE);
    if (!m) return null;

    const seq = parseInt(m[1], 10);
    const desc = normalize(m[2]);
    const qty = parseFloat(m[3]);
    const unit = parseInt(m[4].replace(/,/g, ""), 10);
    const amt = parseInt(m[5].replace(/,/g, ""), 10);

    let instructor = null;
    let course = "";
    let type = "기타";
    let match = "미확정";

    // 1) "... 강사료 김태희"
    if (desc.includes("강사료")) {
      const parts = desc.split(" ");
      const idx = parts.indexOf("강사료");
      if (idx >= 0 && idx + 1 < parts.length && /^[가-힣]{2,4}$/.test(parts[idx + 1])) {
        instructor = parts[idx + 1];
        course = parts.slice(0, idx).join(" ").trim();
        type = "강사료";
        match = "확정";
      }
    }

    // 2) 라인 끝 "… 김태희"
    if (!instructor) {
      const parts = desc.split(" ");
      const last = parts[parts.length - 1];
      if (/^[가-힣]{2,4}$/.test(last)) {
        instructor = last;
        course = parts.slice(0, -1).join(" ").trim();
        type = "강사료";
        match = "확정";
      } else {
        course = desc;
        if (desc.includes("끝전") || desc.includes("충당금")) {
          type = "끝전";
          match = "미확정";
        }
      }
    }

    return { seq, desc, qty, unit_price: unit, amount: amt, instructor, course, type, match_status: match };
  }

  function tokenize(s) {
    const STOP = new Set(["강사료", "강사수당", "끝전", "끝전충당금", "충당금", "지급", "요구", "(목)", "[목]"]);
    const cleaned = (s || "")
      .replace(/[\(\)\[\]\{\}<>]/g, " ")
      .replace(/[^0-9A-Za-z가-힣\s]/g, " ");
    return cleaned
      .split(/\s+/)
      .filter((t) => t && !STOP.has(t) && !/^\d+$/.test(t));
  }

  // 강사명 없는 끝전 행 자동부착(단일 후보만)
  function attachUnassigned(items) {
    const byInst = new Map();
    items.forEach((it) => {
      if (!it.instructor) return;
      if (!byInst.has(it.instructor)) byInst.set(it.instructor, []);
      byInst.get(it.instructor).push(it);
    });

    const instTokens = new Map();
    for (const [inst, rows] of byInst.entries()) {
      const set = new Set();
      rows.forEach((r) => tokenize(r.course).forEach((t) => set.add(t)));
      instTokens.set(inst, set);
    }

    items.forEach((it) => {
      if (it.instructor) return;
      if (it.type !== "끝전") return;

      const toks = new Set(tokenize(it.course || it.desc));
      if (toks.size === 0) return;

      const cands = [];
      for (const [inst, set] of instTokens.entries()) {
        let hit = false;
        for (const t of toks) {
          if (set.has(t)) { hit = true; break; }
        }
        if (hit) cands.push(inst);
      }

      if (cands.length === 1) {
        it.instructor = cands[0];
        it.match_status = "자동부착";
      } else {
        it.match_status = "미확정";
      }
    });

    return items;
  }

  function summarize(items) {
    const map = new Map();
    items.forEach((it) => {
      if (!it.instructor) return;
      if (!map.has(it.instructor)) map.set(it.instructor, { gross: 0, courses: new Set() });
      const d = map.get(it.instructor);
      d.gross += it.amount;
      if (it.course) d.courses.add(it.course);
    });

    return Array.from(map.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([inst, d]) => ({
        instructor: inst,
        courses: Array.from(d.courses).sort().join(" / "),
        gross_amount: d.gross,
      }));
  }

  // ----------------------------
  // 4) 엑셀 생성(SheetJS)
  // ----------------------------
  function buildWorkbook({ rawItems, summaryRows, uncertainRows, settings }) {
    const wb = XLSX.utils.book_new();

    // RAW
    const rawHeader = ["pdf", "page", "seq", "desc", "course", "instructor", "qty", "unit_price", "amount", "type", "match_status"];
    const rawAoA = [rawHeader, ...rawItems.map((r) => rawHeader.map((k) => (r[k] ?? "")))];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawAoA), "RAW_품목내역");

    // 강사별 월보수액
    const rate = settings.expenseRate;
    const sumHeader = ["강사명", "강좌(추출)", "총품의금액(원)", "비과세소득(원)", "필요경비율", "필요경비(원)", "월보수액(원)"];
    const sumAoA = [sumHeader];
    summaryRows.forEach((r) => {
      const nt = settings.nontaxable;
      const expense = (r.gross_amount - nt) * rate;
      const wage = (r.gross_amount - nt) - expense;
      sumAoA.push([r.instructor, r.courses, r.gross_amount, nt, rate, expense, wage]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sumAoA), "강사별_월보수액");

    // 보험료
    const insHeader = ["강사명", "월보수액(원)", "인정월보수액(하한반영)", "산재보험료(총)", "산재(사업주)", "산재(노무)", "고용보험료(총)", "고용(사업주)", "고용(노무)"];
    const insAoA = [insHeader];
    for (let i = 1; i < sumAoA.length; i++) {
      const inst = sumAoA[i][0];
      const wage = sumAoA[i][6];
      const base = Math.max(wage, settings.minBase);
      const wciTotal = base * settings.rateWCI;
      const empTotal = base * settings.rateEMP;
      insAoA.push([inst, wage, base, wciTotal, wciTotal * 0.5, wciTotal * 0.5, empTotal, empTotal * 0.5, empTotal * 0.5]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(insAoA), "보험료_산정");

    // 미확정
    const unHeader = ["pdf", "desc", "amount"];
    const unAoA = [unHeader, ...uncertainRows.map((r) => [r.pdf, r.desc, r.amount])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(unAoA), "미확정_검토");

    // 설정(참고)
    const setAoA = [
      ["지급월(YYYY-MM)", settings.payMonth],
      ["필요경비 적용기간", settings.periodLabel],
      ["필요경비율 직접입력(%)", settings.customRatePercent ?? ""],
      ["비과세소득(합계)", settings.nontaxable],
      ["기준보수액 하한", settings.minBase],
      ["산재보험료율(총)", settings.rateWCI],
      ["고용보험료율(총)", settings.rateEMP],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(setAoA), "설정");

    return wb;
  }

  // ----------------------------
  // 5) 화면 렌더
  // ----------------------------
  function renderTables(summaryRows, uncertainRows, settings) {
    const tb = $("tblSummary").querySelector("tbody");
    const ub = $("tblUncertain").querySelector("tbody");
    tb.innerHTML = "";
    ub.innerHTML = "";

    const nt = settings.nontaxable;
    const rate = settings.expenseRate;

    summaryRows.forEach((r) => {
      const expense = (r.gross_amount - nt) * rate;
      const wage = (r.gross_amount - nt) - expense;

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.instructor}</td>
        <td>${r.courses || ""}</td>
        <td class="right">${fmt(r.gross_amount)}</td>
        <td class="right">${fmt(wage)}</td>
      `;
      tb.appendChild(tr);
    });

    uncertainRows.forEach((r) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.pdf}</td>
        <td>${r.desc}</td>
        <td class="right">${fmt(r.amount)}</td>
      `;
      ub.appendChild(tr);
    });
  }

  // ----------------------------
  // 6) 메인 실행
  // ----------------------------
  async function run() {
    if (!state.files.length) {
      status.innerHTML = `<span class="warn">PDF 파일을 선택하십시오.</span>`;
      return;
    }

    $("btnRun").disabled = true;
    $("btnDownload").disabled = true;
    status.textContent = "처리 중…";

    try {
      const rawItems = [];

      for (let i = 0; i < state.files.length; i++) {
        const f = state.files[i];
        status.textContent = `PDF 분석 중… (${i + 1}/${state.files.length}) ${f.name}`;

        const pages = await extractTextFromPdf(f);
        pages.forEach(({ page, text }) => {
          // 품목내역 이후만(간단 컷) — “품목내역”이 안 잡히면 전체에서라도 파싱 시도
          const cut = text.includes("품목내역") ? text.split("품목내역").slice(1).join(" ") : text;
          const injected = injectRowBreaks(cut);
          const lines = injected.split("\n").map((s) => s.trim()).filter(Boolean);

          lines.forEach((line) => {
            const parsed = parseRow(line);
            if (!parsed) return;

            rawItems.push({
              pdf: f.name,
              page,
              ...parsed,
            });
          });
        });
      }

      attachUnassigned(rawItems);

      const summaryRows = summarize(rawItems);
      const uncertainRows = rawItems.filter(
        (it) => (it.type === "끝전") && (!it.instructor || it.match_status === "미확정")
      );

      const settings = {
        payMonth: ($("payMonth").value || "").trim(),
        periodLabel: $("periodSel").value,
        customRatePercent: ($("customRate").value || "").trim() === "" ? null : parseFloat($("customRate").value),
        expenseRate: getSelectedExpenseRate(),
        nontaxable: parseFloat($("nontaxable").value || "0") || 0,
        minBase: parseFloat($("minBase").value || "1330000") || 1330000,
        rateWCI: parseFloat($("rateWCI").value || "0.0066") || 0.0066,
        rateEMP: parseFloat($("rateEMP").value || "0.016") || 0.016,
      };

      const wb = buildWorkbook({ rawItems, summaryRows, uncertainRows, settings });

      state.rawItems = rawItems;
      state.summary = summaryRows;
      state.uncertain = uncertainRows;
      state.workbook = wb;

      renderTables(summaryRows, uncertainRows, settings);

      status.textContent = `완료. (추출행 ${rawItems.length} / 강사 ${summaryRows.length} / 미확정 ${uncertainRows.length})`;
      $("btnDownload").disabled = false;
    } catch (e) {
      console.error(e);
      status.innerHTML = `<span class="warn">오류:</span> ${String(e.message || e)}`;
    } finally {
      $("btnRun").disabled = false;
    }
  }

  function downloadExcel() {
    if (!state.workbook) {
      alert("먼저 계산을 실행하세요.");
      return;
    }
    const pm = ($("payMonth").value || "").trim().replace(/[^0-9-]/g, "");
    const fname = `방과후강사_월보수액_산출_${pm || "결과"}.xlsx`;
    XLSX.writeFile(state.workbook, fname, { compression: true });
  }

  $("btnRun").addEventListener("click", run);
  $("btnDownload").addEventListener("click", downloadExcel);
})();