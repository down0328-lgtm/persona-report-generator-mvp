let originalData = [];
let headers = [];
let uploadedFileName = "";
let selectedColumns = {};
let personaResults = [];
let currentChart = null;
let latestChartRows = [];
let latestChartMeta = {};

const columnRoles = [
  { key: "age", label: "연령대" },
  { key: "gender", label: "성별" },
  { key: "nationality", label: "국적" },
  { key: "visa", label: "체류자격" },
  { key: "amount", label: "거래금액" },
  { key: "count", label: "거래건수" },
  { key: "product", label: "상품" },
  { key: "channel", label: "채널" },
];

const groupPriority = ["nationality", "visa", "age", "gender", "product", "channel"];
const chartColors = ["#0F766E", "#2563EB", "#D97706", "#7C3AED", "#DC2626", "#0891B2", "#65A30D", "#BE185D", "#475569", "#9333EA"];
const GEMINI_SYSTEM_PROMPT = `
너는 업로드된 엑셀 데이터를 기반으로 차트와 주요 수치를 해석하는 데이터 분석 전략가다.
너의 역할은 단순 요약이 아니라, 데이터 안에서 의미 있는 패턴, 추세, 이상치, 고객군 특징, 비즈니스 기회, 리스크 요인을 찾아내는 것이다.

분석 시 다음 관점을 반드시 적용해라.
1. 추세 관점: 월별 증가·감소, 급증·급감, 평균 대비 특이 구간, 계절성 가능성을 확인한다.
2. 구성비 관점: 상위 항목, 점유율, 특정 항목 쏠림, 기타 그룹 비중을 확인한다.
3. 비교 관점: 국가별, 체류자격별, 연령대별, 상품별, 채널별, 고객군별 차이를 비교한다.
4. 이상치 관점: 특정 월만 튀는 현상, 특정 그룹 집중, 거래건수와 거래금액 불일치, 고객 수 대비 금액 차이를 확인한다.
5. 비즈니스 활용 관점: 타겟 마케팅, 상품 제안, 수수료 전략, 채널 전략, 고객 세분화, 운영 리스크 관점에서 활용 방안을 제시한다.

답변은 반드시 아래 형식으로 작성해라.
1. 데이터 요약
전체 데이터 규모, 주요 컬럼, 분석 기준, 핵심 지표를 요약한다.
2. 주요 차트 해석
차트에서 가장 눈에 띄는 패턴, 증가·감소 흐름, 상위 항목, 특이 구간을 설명한다.
3. 핵심 인사이트
최소 3개 이상 작성한다. 각 인사이트는 "인사이트 제목:", "분석 내용:", "비즈니스 의미:", "활용 방안:" 구조로 작성한다.
4. 이상치 및 주의할 점
급증 구간, 급감 구간, 데이터 편중, 추가 확인이 필요한 항목을 정리한다.
5. 추천 액션
실무자가 바로 실행할 수 있는 액션을 제안한다.
6. 보고서용 요약 문장
PDF 보고서에 바로 넣을 수 있는 문장으로 3~5문장 작성한다.

주의사항:
데이터에 없는 내용은 단정하지 마라.
원인 추정은 “가능성이 있다”, “추정된다”, “추가 확인이 필요하다”라고 표현해라.
숫자는 가능한 한 구체적으로 언급해라.
단순 설명이 아니라 의미를 해석해라.
보고서 문체로 작성해라.
실무자가 바로 활용할 수 있는 표현으로 작성해라.
`.trim();

const elements = {
  fileInput: document.getElementById("excelFile"),
  fileName: document.getElementById("fileName"),
  rowCount: document.getElementById("rowCount"),
  columnCount: document.getElementById("columnCount"),
  message: document.getElementById("message"),
  previewSection: document.getElementById("previewSection"),
  previewMeta: document.getElementById("previewMeta"),
  previewTable: document.getElementById("previewTable"),
  chartSection: document.getElementById("chartSection"),
  chartCategoryColumn: document.getElementById("chartCategoryColumn"),
  chartValueColumn: document.getElementById("chartValueColumn"),
  chartMetric: document.getElementById("chartMetric"),
  chartType: document.getElementById("chartType"),
  drawChartButton: document.getElementById("drawChartButton"),
  dataChart: document.getElementById("dataChart"),
  selectorSection: document.getElementById("selectorSection"),
  columnSelectors: document.getElementById("columnSelectors"),
  generateButton: document.getElementById("generateButton"),
  reportSection: document.getElementById("reportSection"),
  reportContent: document.getElementById("reportContent"),
  downloadButton: document.getElementById("downloadButton"),
  insightSection: document.getElementById("insightSection"),
  geminiApiKey: document.getElementById("geminiApiKey"),
  geminiModel: document.getElementById("geminiModel"),
  generateInsightButton: document.getElementById("generateInsightButton"),
  insightOutput: document.getElementById("insightOutput"),
  resetButton: document.getElementById("resetButton"),
};

elements.fileInput.addEventListener("change", handleFileUpload);
elements.drawChartButton.addEventListener("click", renderChart);
elements.chartCategoryColumn.addEventListener("change", renderChart);
elements.chartValueColumn.addEventListener("change", renderChart);
elements.chartMetric.addEventListener("change", renderChart);
elements.chartType.addEventListener("change", renderChart);
elements.generateButton.addEventListener("click", generatePersonas);
elements.downloadButton.addEventListener("click", downloadPDF);
elements.generateInsightButton.addEventListener("click", generateGeminiInsights);
elements.resetButton.addEventListener("click", resetApp);

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    showMessage("엑셀 파일만 업로드할 수 있습니다.", "error");
    elements.fileInput.value = "";
    return;
  }

  uploadedFileName = file.name;
  parseExcelFile(file);
}

function parseExcelFile(file) {
  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const firstSheetName = workbook.SheetNames[0];

      if (!firstSheetName) {
        showMessage("데이터가 없습니다.", "warn");
        return;
      }

      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      const normalizedRows = rows.filter((row) => row.some((cell) => String(cell).trim() !== ""));

      if (normalizedRows.length < 2) {
        showMessage("데이터가 없습니다.", "warn");
        return;
      }

      headers = normalizedRows[0].map((header, index) => String(header || `컬럼 ${index + 1}`).trim());
      originalData = normalizedRows.slice(1).map((row) => {
        const record = {};
        headers.forEach((header, index) => {
          record[header] = row[index] ?? "";
        });
        return record;
      });

      elements.fileName.textContent = uploadedFileName;
      elements.rowCount.textContent = originalData.length.toLocaleString("ko-KR");
      elements.columnCount.textContent = headers.length.toLocaleString("ko-KR");

      renderPreviewTable();
      renderColumnSelectors();
      renderChartSelectors();
      renderChart();
      elements.previewSection.classList.remove("hidden");
      elements.chartSection.classList.remove("hidden");
      elements.selectorSection.classList.remove("hidden");
      elements.reportSection.classList.add("hidden");
      elements.insightSection.classList.add("hidden");
      showMessage("엑셀 파일을 정상적으로 읽었습니다.");
    } catch (error) {
      console.error(error);
      showMessage("파일을 읽는 중 오류가 발생했습니다.", "error");
    }
  };

  reader.onerror = () => showMessage("파일을 읽는 중 오류가 발생했습니다.", "error");
  reader.readAsArrayBuffer(file);
}

function renderPreviewTable() {
  const previewRows = originalData.slice(0, 10);
  elements.previewMeta.textContent = `총 ${originalData.length.toLocaleString("ko-KR")}행 중 상위 ${previewRows.length}행 표시`;

  const head = `<thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>`;
  const body = previewRows
    .map((row) => `<tr>${headers.map((header) => `<td title="${escapeHtml(row[header])}">${escapeHtml(row[header])}</td>`).join("")}</tr>`)
    .join("");

  elements.previewTable.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderChartSelectors() {
  const categoryGuess = guessCategoryColumn();
  const valueGuess = guessNumericColumn();
  const headerOptions = headers.map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`).join("");
  const valueOptions = [`<option value="">값 컬럼 없음</option>`, headerOptions].join("");

  elements.chartCategoryColumn.innerHTML = headerOptions;
  elements.chartValueColumn.innerHTML = valueOptions;
  elements.chartCategoryColumn.value = categoryGuess || headers[0] || "";
  elements.chartValueColumn.value = valueGuess || "";
  elements.chartMetric.value = valueGuess ? "sum" : "count";
}

function renderChart() {
  if (!originalData.length || !window.Chart) return;

  const categoryColumn = elements.chartCategoryColumn.value || headers[0];
  const valueColumn = elements.chartValueColumn.value;
  const metric = elements.chartMetric.value;
  const chartType = elements.chartType.value;
  const chartRows = buildChartRows(categoryColumn, valueColumn, metric);
  latestChartRows = chartRows;
  latestChartMeta = { categoryColumn, valueColumn, metric, chartType };

  if (!chartRows.length) {
    showMessage("차트로 표시할 데이터가 없습니다.", "warn");
    return;
  }

  if (currentChart) currentChart.destroy();

  const labels = chartRows.map((row) => row.label);
  const values = chartRows.map((row) => row.value);
  const isCircular = ["pie", "doughnut", "polarArea"].includes(chartType);

  currentChart = new Chart(elements.dataChart, {
    type: chartType,
    data: {
      labels,
      datasets: [
        {
          label: chartDatasetLabel(categoryColumn, valueColumn, metric),
          data: values,
          borderColor: "#0F766E",
          backgroundColor: isCircular ? labels.map((_, index) => chartColors[index % chartColors.length]) : "rgba(15, 118, 110, 0.72)",
          pointBackgroundColor: "#0F766E",
          fill: chartType === "radar",
          tension: chartType === "line" ? 0.28 : 0,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: isCircular || chartType === "radar" },
        title: {
          display: true,
          text: chartDatasetLabel(categoryColumn, valueColumn, metric),
          color: "#17211f",
          font: { size: 16, weight: "bold" },
        },
      },
      scales: isCircular ? {} : { y: { beginAtZero: true } },
    },
  });
}

function buildChartRows(categoryColumn, valueColumn, metric) {
  const groups = new Map();

  originalData.forEach((row) => {
    const label = normalizeText(row[categoryColumn]) || "미분류";
    if (!groups.has(label)) groups.set(label, []);
    groups.get(label).push(row);
  });

  return Array.from(groups.entries())
    .map(([label, rows]) => {
      const numbers = valueColumn ? rows.map((row) => toNumber(row[valueColumn])).filter(Number.isFinite) : [];
      const value = metric === "count" || !valueColumn ? rows.length : metric === "avg" ? average(numbers) : sum(numbers);
      return { label, value };
    })
    .filter((row) => Number.isFinite(row.value))
    .sort((a, b) => b.value - a.value)
    .slice(0, 10);
}

function renderColumnSelectors() {
  elements.columnSelectors.innerHTML = columnRoles
    .map((role) => {
      const guessedHeader = guessHeader(role.label);
      const options = [
        `<option value="">선택 안 함</option>`,
        ...headers.map((header) => {
          const selected = header === guessedHeader ? "selected" : "";
          return `<option value="${escapeHtml(header)}" ${selected}>${escapeHtml(header)}</option>`;
        }),
      ].join("");

      return `
        <div class="field">
          <label for="select-${role.key}">${role.label} 컬럼</label>
          <select id="select-${role.key}" data-role="${role.key}">
            ${options}
          </select>
        </div>
      `;
    })
    .join("");
}

function generatePersonas() {
  selectedColumns = {};
  document.querySelectorAll("[data-role]").forEach((select) => {
    if (select.value) selectedColumns[select.dataset.role] = select.value;
  });

  if (Object.keys(selectedColumns).length === 0) {
    showMessage("페르소나 생성을 위해 최소 1개 이상의 컬럼을 선택해주세요.", "warn");
    return;
  }

  const groupKey = groupPriority.find((key) => selectedColumns[key]) || Object.keys(selectedColumns)[0];
  personaResults = analyzeGroupData(groupKey);

  if (personaResults.length === 0) {
    showMessage("데이터가 없습니다.", "warn");
    return;
  }

  renderPersonaReport(groupKey);
  elements.reportSection.classList.remove("hidden");
  elements.insightSection.classList.remove("hidden");
  showMessage("페르소나 보고서가 생성되었습니다.");
}

function analyzeGroupData(groupKey) {
  const groupColumn = selectedColumns[groupKey];
  const groups = new Map();

  originalData.forEach((row) => {
    const name = normalizeText(row[groupColumn]) || "미분류";
    if (!groups.has(name)) {
      groups.set(name, { name, rows: [], amountValues: [], countValues: [], products: [], channels: [] });
    }

    const group = groups.get(name);
    group.rows.push(row);

    if (selectedColumns.amount) group.amountValues.push(toNumber(row[selectedColumns.amount]));
    if (selectedColumns.count) group.countValues.push(toNumber(row[selectedColumns.count]));
    if (selectedColumns.product) group.products.push(normalizeText(row[selectedColumns.product]));
    if (selectedColumns.channel) group.channels.push(normalizeText(row[selectedColumns.channel]));
  });

  return Array.from(groups.values())
    .map((group) => {
      const amountNumbers = group.amountValues.filter(Number.isFinite);
      const countNumbers = group.countValues.filter(Number.isFinite);

      return {
        ...group,
        customerCount: group.rows.length,
        share: group.rows.length / originalData.length,
        amountTotal: sum(amountNumbers),
        amountAverage: average(amountNumbers),
        countTotal: sum(countNumbers),
        countAverage: average(countNumbers),
        topProduct: mode(group.products.filter(Boolean)),
        topChannel: mode(group.channels.filter(Boolean)),
      };
    })
    .sort((a, b) => b.customerCount - a.customerCount)
    .slice(0, 3);
}

function renderPersonaReport(groupKey) {
  const selectedList = Object.entries(selectedColumns)
    .map(([role, column]) => `${roleLabel(role)}: ${column}`)
    .join(", ");
  const createdAt = new Date().toLocaleDateString("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit" });
  const groupLabel = roleLabel(groupKey);

  const cards = personaResults
    .map((persona, index) => {
      const traits = buildTraits(persona);
      return `
        <article class="persona-card">
          <h3>페르소나 ${index + 1}. ${personaName(persona, index)}</h3>
          <dl>
            <dt>대표 고객군</dt>
            <dd>${escapeHtml(groupLabel)}: ${escapeHtml(persona.name)}</dd>
            <dt>고객 비중</dt>
            <dd>전체의 ${formatPercent(persona.share)} (${persona.customerCount.toLocaleString("ko-KR")}명)</dd>
            <dt>주요 특징</dt>
            <dd>${traits.feature}</dd>
            <dt>거래 특성</dt>
            <dd>${traits.trade}</dd>
            <dt>주요 니즈</dt>
            <dd>${traits.need}</dd>
            <dt>추천 접근 전략</dt>
            <dd>${traits.strategy}</dd>
            <dt>예상 활용 포인트</dt>
            <dd>${traits.usePoint}</dd>
          </dl>
        </article>
      `;
    })
    .join("");

  elements.reportContent.innerHTML = `
    <h2>페르소나 분석 보고서</h2>
    <div class="report-meta">
      <div><span>생성일자</span><strong>${createdAt}</strong></div>
      <div><span>업로드 파일명</span><strong>${escapeHtml(uploadedFileName)}</strong></div>
      <div><span>분석 대상 행 수</span><strong>${originalData.length.toLocaleString("ko-KR")}행</strong></div>
      <div><span>그룹 기준</span><strong>${escapeHtml(groupLabel)}</strong></div>
    </div>
    <p><strong>선택한 컬럼</strong>: ${escapeHtml(selectedList)}</p>
    <p class="summary-box">${escapeHtml(groupLabel)} 기준 고객 수가 많은 상위 ${personaResults.length}개 그룹을 추출했습니다.</p>
    <div class="persona-grid">${cards}</div>
  `;
}

function downloadPDF() {
  if (!personaResults.length) return;

  const baseName = uploadedFileName.replace(/\.(xlsx|xls)$/i, "") || "persona";
  const options = {
    margin: 10,
    filename: `${baseName}_persona_report.pdf`,
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
  };

  html2pdf()
    .set(options)
    .from(elements.reportContent)
    .save()
    .catch(() => showMessage("PDF 생성 중 오류가 발생했습니다.", "error"));
}

async function generateGeminiInsights() {
  const apiKey = elements.geminiApiKey.value.trim();
  const model = elements.geminiModel.value;

  if (!apiKey) {
    showMessage("Gemini API 키를 입력해주세요.", "warn");
    elements.geminiApiKey.focus();
    return;
  }

  if (!personaResults.length) {
    showMessage("먼저 페르소나를 생성해주세요.", "warn");
    return;
  }

  elements.generateInsightButton.disabled = true;
  elements.insightOutput.classList.add("loading");
  elements.insightOutput.textContent = "Gemini가 차트와 페르소나를 분석하는 중입니다...";

  try {
    const response = await fetch("https://generativelanguage.googleapis.com/v1beta/interactions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey,
      },
      body: JSON.stringify({
        model,
        system_instruction: GEMINI_SYSTEM_PROMPT,
        input: buildGeminiInput(),
        generation_config: {
          temperature: 0.4,
        },
      }),
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result?.error?.message || "Gemini API 호출에 실패했습니다.");
    }

    const text = extractGeminiText(result);
    elements.insightOutput.classList.remove("loading");
    elements.insightOutput.innerHTML = formatInsightText(text || "인사이트 응답이 비어 있습니다.");
    showMessage("Gemini 인사이트 분석이 완료되었습니다.");
  } catch (error) {
    console.error(error);
    elements.insightOutput.classList.remove("loading");
    elements.insightOutput.textContent = buildInsightErrorMessage(error.message);
    showMessage("Gemini 인사이트 생성 중 오류가 발생했습니다.", "error");
  } finally {
    elements.generateInsightButton.disabled = false;
  }
}

function buildInsightErrorMessage(message) {
  const rawMessage = message || "알 수 없는 오류가 발생했습니다.";
  const isHighDemand = /high demand|try again later|overloaded|unavailable|503/i.test(rawMessage);
  const suggestion = isHighDemand
    ? "\n\n선택한 모델이 일시적으로 혼잡합니다. Gemini 모델을 gemini-2.5-flash 또는 gemini-2.0-flash로 바꾼 뒤 다시 눌러보세요."
    : "\n\nAPI 키, 모델 권한, 네트워크 상태를 확인한 뒤 다시 시도해주세요.";

  return `인사이트 생성 중 오류가 발생했습니다.\n${rawMessage}${suggestion}`;
}

function resetApp() {
  originalData = [];
  headers = [];
  uploadedFileName = "";
  selectedColumns = {};
  personaResults = [];
  latestChartRows = [];
  latestChartMeta = {};

  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  elements.fileInput.value = "";
  elements.fileName.textContent = "-";
  elements.rowCount.textContent = "-";
  elements.columnCount.textContent = "-";
  elements.previewTable.innerHTML = "";
  elements.columnSelectors.innerHTML = "";
  elements.reportContent.innerHTML = "";
  elements.insightOutput.textContent = "페르소나 생성 후 Gemini API 키를 입력하고 인사이트를 받아보세요.";
  elements.previewSection.classList.add("hidden");
  elements.chartSection.classList.add("hidden");
  elements.selectorSection.classList.add("hidden");
  elements.reportSection.classList.add("hidden");
  elements.insightSection.classList.add("hidden");
  showMessage("업로드 전 엑셀 파일을 업로드해주세요.");
}

function showMessage(text, type = "success") {
  elements.message.textContent = text;
  elements.message.className = `message ${type === "success" ? "" : type}`.trim();
}

function guessHeader(label) {
  const compactLabel = label.replace(/\s/g, "");
  const exact = headers.find((header) => header.replace(/\s/g, "") === compactLabel);
  if (exact) return exact;
  return headers.find((header) => header.includes(label.replace("거래", ""))) || "";
}

function guessCategoryColumn() {
  return headers.find((header) => originalData.some((row) => !Number.isFinite(toNumber(row[header])))) || headers[0] || "";
}

function guessNumericColumn() {
  return headers.find((header) => {
    const sample = originalData.slice(0, 30).map((row) => toNumber(row[header]));
    return sample.filter(Number.isFinite).length >= Math.max(2, Math.floor(sample.length * 0.5));
  });
}

function chartDatasetLabel(categoryColumn, valueColumn, metric) {
  const metricLabel = metric === "avg" ? "평균" : metric === "sum" ? "합계" : "건수";
  const target = valueColumn && metric !== "count" ? `${valueColumn} ${metricLabel}` : metricLabel;
  return `${categoryColumn}별 ${target}`;
}

function buildGeminiInput() {
  const numericSummary = headers
    .map((header) => {
      const values = originalData.map((row) => toNumber(row[header])).filter(Number.isFinite);
      if (!values.length) return null;
      return {
        column: header,
        count: values.length,
        sum: Math.round(sum(values)),
        average: Math.round(average(values) * 100) / 100,
        min: Math.min(...values),
        max: Math.max(...values),
      };
    })
    .filter(Boolean)
    .slice(0, 10);

  const sampleRows = originalData.slice(0, 8).map((row) => {
    const compact = {};
    headers.slice(0, 12).forEach((header) => {
      compact[header] = row[header];
    });
    return compact;
  });

  const payload = {
    fileName: uploadedFileName,
    rowCount: originalData.length,
    columnCount: headers.length,
    headers,
    selectedColumns,
    chart: {
      ...latestChartMeta,
      topRows: latestChartRows,
    },
    personas: personaResults.map((persona) => ({
      name: persona.name,
      customerCount: persona.customerCount,
      share: formatPercent(persona.share),
      amountTotal: persona.amountTotal,
      amountAverage: persona.amountAverage,
      countTotal: persona.countTotal,
      countAverage: persona.countAverage,
      topProduct: persona.topProduct,
      topChannel: persona.topChannel,
    })),
    numericSummary,
    sampleRows,
  };

  return `아래 JSON은 사용자가 업로드한 엑셀 데이터에서 계산한 차트 분석 결과와 페르소나 요약입니다. 첨부 페르소나 지침의 형식에 맞춰 한국어 보고서 문체로 인사이트를 작성하세요.\n\n${JSON.stringify(payload, null, 2)}`;
}

function extractGeminiText(result) {
  if (typeof result.output_text === "string") return result.output_text;
  const texts = [];
  (result.steps || []).forEach((step) => {
    (step.content || step.contents || []).forEach((content) => {
      if (typeof content.text === "string") texts.push(content.text);
    });
  });
  return texts.join("\n").trim();
}

function formatInsightText(text) {
  const safe = escapeHtml(text);
  return safe
    .replace(/^(\d+\.\s*[^\n]+)/gm, "<h3>$1</h3>")
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/\n/g, "<br>");
}

function buildTraits(persona) {
  const amountAverage = average(personaResults.map((item) => item.amountAverage).filter(Number.isFinite));
  const countAverage = average(personaResults.map((item) => item.countAverage).filter(Number.isFinite));
  const hasHighAmount = persona.amountAverage > amountAverage;
  const hasHighCount = persona.countAverage > countAverage;
  const channelText = persona.topChannel ? `${persona.topChannel} 채널 이용 경향이 뚜렷합니다.` : "선택된 채널 컬럼이 없어 채널 특성은 제외했습니다.";
  const productText = persona.topProduct ? `${persona.topProduct} 상품 선호가 가장 높습니다.` : "선택된 상품 컬럼이 없어 상품 특성은 제외했습니다.";

  return {
    feature: hasHighAmount
      ? "거래 규모가 크고 수익 기여도가 높은 핵심 고객군입니다."
      : hasHighCount
        ? "반복 거래 가능성이 높고 정기 관리가 필요한 고객군입니다."
        : "전체 데이터에서 높은 비중을 차지하는 대표 고객군입니다.",
    trade: `${formatMoneyText(persona.amountTotal, "거래금액 합계")} ${formatMoneyText(persona.amountAverage, "평균 거래금액")} ${formatCountText(persona.countAverage)} ${productText} ${channelText}`,
    need: hasHighCount ? "빠른 처리, 간편한 재이용, 정기 알림에 대한 니즈가 높습니다." : "명확한 혜택, 맞춤형 안내, 안정적인 이용 경험이 중요합니다.",
    strategy: persona.topChannel
      ? `${persona.topChannel} 중심의 맞춤 혜택과 재거래 유도 메시지를 제안합니다.`
      : "수수료 우대, 맞춤 상품 제안, 이용 단계별 안내를 제안합니다.",
    usePoint: "마케팅 타깃 세분화, 상담 스크립트, 상품 추천 기준으로 활용할 수 있습니다.",
  };
}

function personaName(persona, index) {
  if (persona.countAverage > 0) return "고빈도 생활거래형 고객";
  if (persona.amountAverage > 0) return "고가치 거래중심형 고객";
  const names = ["대표 성장형 고객", "집중 관리형 고객", "맞춤 제안형 고객"];
  return names[index] || "핵심 고객";
}

function roleLabel(key) {
  return columnRoles.find((role) => role.key === key)?.label || key;
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function toNumber(value) {
  if (typeof value === "number") return value;
  const cleaned = String(value ?? "").replace(/[^\d.-]/g, "");
  return cleaned ? Number(cleaned) : NaN;
}

function sum(values) {
  return values.reduce((total, value) => total + value, 0);
}

function average(values) {
  return values.length ? sum(values) / values.length : 0;
}

function mode(values) {
  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] || "";
}

function formatPercent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function formatMoneyText(value, label) {
  if (!Number.isFinite(value) || value === 0) return `${label}: -`;
  return `${label}: ${Math.round(value).toLocaleString("ko-KR")}`;
}

function formatCountText(value) {
  if (!Number.isFinite(value) || value === 0) return "평균 거래건수: -";
  return `평균 거래건수: ${value.toFixed(1)}건`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
