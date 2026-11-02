const SVG_NS = "http://www.w3.org/2000/svg";

const categories = [
  {
    short: "空",
    name: "空港委",
    metrics: [
      { name: "航空运输", value: 0.7 },
      { name: "航空客货", value: 0.7 },
    ],
  },
  {
    short: "港",
    name: "港务局",
    metrics: [
      { name: "水路运输", value: 0.5 },
      { name: "管道运输", value: 0.1 },
    ],
  },
  {
    short: "发",
    name: "发改委",
    metrics: [
      { name: "采矿业", value: 0 },
      { name: "公管工资", value: 3.8 },
    ],
  },
  {
    short: "住",
    name: "住建局",
    metrics: [
      { name: "建安工程", value: 2.4 },
      { name: "专包企业", value: 2.4 },
      { name: "商品房销售", value: 1.8 },
      { name: "房产企业", value: 2.9 },
    ],
  },
  {
    short: "邮",
    name: "邮政局",
    metrics: [{ name: "邮政行业", value: 0.2 }],
  },
  {
    short: "文",
    name: "文广旅局",
    metrics: [
      { name: "文娱营收", value: 0.5 },
      { name: "文娱工资", value: 0.5 },
    ],
  },
  {
    short: "卫",
    name: "卫健委",
    metrics: [{ name: "卫生工资", value: 3.2 }],
  },
  {
    short: "金",
    name: "金融局",
    metrics: [
      { name: "证券交易", value: 1.2 },
      { name: "保费收入", value: 1.1 },
      { name: "利佣净收入", value: 7.5 },
    ],
  },
  {
    short: "商",
    name: "商务局",
    metrics: [
      { name: "批发业", value: 8.6 },
      { name: "租商营收", value: 6.5 },
      { name: "零售业", value: 4.5 },
      { name: "住宿业", value: 0.3 },
      { name: "餐饮业", value: 1.6 },
      { name: "服务营收", value: 1.7 },
    ],
  },
  {
    short: "水",
    name: "水务局",
    metrics: [
      { name: "水利营收", value: 0.3 },
      { name: "水利工资", value: 0.3 },
    ],
  },
  {
    short: "农",
    name: "农业农村局",
    metrics: [{ name: "农业", value: 1.2 }],
  },
  {
    short: "科",
    name: "科技局",
    metrics: [
      { name: "科技营收", value: 2.1 },
      { name: "科技工资", value: 2.1 },
    ],
  },
  {
    short: "教",
    name: "教育局",
    metrics: [{ name: "教育工资", value: 4.2 }],
  },
  {
    short: "交",
    name: "交通局",
    metrics: [
      { name: "铁路运输", value: 0.2 },
      { name: "公路运输", value: 2.4 },
      { name: "多式联运", value: 1.2 },
      { name: "装卸仓储", value: 0.3 },
    ],
  },
  {
    short: "工",
    name: "工信局",
    metrics: [
      { name: "制造业", value: 15.9 },
      { name: "互软信业", value: 7.6 },
      { name: "规下工业", value: 3.3 },
      { name: "水电燃气", value: 1.6 },
      { name: "电信业", value: 0.9 },
    ],
  },
];

// Source: 指标亮灯统计表.xlsx, Sheet1. Each string covers months 2-9.
const monthlyLightData = {
  建安工程: "RRRRRRRR",
  商品房销售: "RGRYRRRR",
  房产企业: "NRRRRRRR",
  邮政行业: "NRRRRRRR",
  文娱营收: "NRRRRRRR",
  文娱工资: "NGNNRNNR",
  卫生工资: "NRNNGNNG",
  证券交易: "GGGGGGGG",
  保费收入: "RRRRRRRR",
  利佣净收入: "GGGNRGGY",
  批发业: "RYYRYRRY",
  租商营收: "NGGGGGGG",
  零售业: "GGRRRRRR",
  住宿业: "GGGGGGGG",
  餐饮业: "RRRRRRRR",
  服务营收: "NRRRRRRR",
  水利营收: "NRRRRRRR",
  水利工资: "NGNNRNNR",
  农业: "NGNNGNNG",
  科技营收: "NGRRRRRR",
  科技工资: "NGNNRNNY",
  教育工资: "NGNNGNNG",
  铁路运输: "GGGGGYGG",
  公路运输: "RRRRRYGG",
  多式联运: "NYGGGGGG",
  装卸仓储: "NGRRRRRY",
  公管工资: "NRNNRNNR",
  采矿业: "NGNGGRYG",
  互软信业: "NNRNNNNN",
  制造业: "NYNRYRRR",
  水电燃气: "NRNRRRRR",
  规下工业: "NRNNRNNR",
  电信业: "RRRRRRRR",
  航空运输: "GGRRRRRR",
  航空客货: "GYRRRRRR",
  水路运输: "GGGGGYYY",
  管道运输: "NRNNNNNN",
  专包企业: "NGNNRNNR",
};

const lightStatuses = {
  G: { key: "增长", light: "绿灯", color: "#57cc88", weight: 1.2 },
  Y: { key: "关注", light: "黄灯", color: "#f5c451", weight: 1.0 },
  R: { key: "回落", light: "红灯", color: "#ef6b6b", weight: 0.82 },
  N: { key: "暂无", light: "暂无", color: "#d3d6da", weight: 0.66 },
};

function monthStatusFromWorkbook(metricName, month) {
  if (month === 1) return lightStatuses.N;
  const light = monthlyLightData[metricName]?.[month - 2] || "N";
  return lightStatuses[light] || lightStatuses.N;
}

function barLengthRatioFromStatuses(statuses) {
  const averageWeight =
    statuses.reduce((sum, status) => sum + status.weight, 0) / statuses.length;
  const normalized = (averageWeight - 0.66) / (1.2 - 0.66);
  return 0.72 + clamp(normalized, 0, 1) * 0.28;
}

function dominantStatus(statuses) {
  const counts = new Map();
  statuses.forEach((status) => {
    counts.set(status.key, (counts.get(status.key) || 0) + 1);
  });
  const key = [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
  return statuses.find((status) => status.key === key);
}

categories.forEach((category) => {
  category.metrics.forEach((metric) => {
    const rawMonths = Array.from({ length: 9 }, (_, monthIndex) => {
      const status = monthStatusFromWorkbook(metric.name, monthIndex + 1);
      return {
        status,
        weight: status.weight,
      };
    });
    const weights = rawMonths.map((month) => month.weight);
    const weightTotal = weights.reduce((sum, weight) => sum + weight, 0);

    metric.months = rawMonths.map(({ status }, monthIndex) => {
      const weight = weights[monthIndex];
      const normalizedWeight = weight / weightTotal;
      return {
        month: monthIndex + 1,
        value: metric.value * normalizedWeight,
        weight: normalizedWeight,
        status,
      };
    });
  });

  category.months = Array.from({ length: 9 }, (_, monthIndex) => {
    const totalValue = category.metrics.reduce(
      (sum, metric) => sum + metric.months[monthIndex].value,
      0,
    );
    const status = dominantStatus(
      category.metrics.map((metric) => metric.months[monthIndex].status),
    );

    return {
      month: monthIndex + 1,
      value: totalValue,
      status,
    };
  });
});

const state = {
  filter: "全部",
  selectedCategory: null,
  focusedMetric: null,
  zoom: {
    scale: 1,
    x: 0,
    y: 0,
  },
};

const chart = document.querySelector("#indicatorChart");
const chartCanvas = document.querySelector("#chartCanvas");
const tooltip = document.querySelector("#chartTooltip");
const chartZoomStatus = document.querySelector("#chartZoomStatus");
const chartZoomDetail = document.querySelector("#chartZoomDetail");
const filterButton = document.querySelector("#filterButton");
const filterLabel = document.querySelector("#filterLabel");
const filterMenu = document.querySelector("#filterMenu");
const chartReset = document.querySelector("#chartReset");
const summary = document.querySelector("#chartSummary");
const toast = document.querySelector("#toast");
const appShell = document.querySelector(".app-shell");

const center = { x: 610, y: 430 };
const gapDegrees = 1.2;
const fullCircle = 360;
const totalMetricCount = categories.reduce((sum, category) => sum + category.metrics.length, 0);
const metricArc = fullCircle / totalMetricCount;
const maxMetricValue = Math.max(
  ...categories.flatMap((category) => category.metrics.map((metric) => metric.value)),
);
const uniformStackLength = maxMetricValue * 11.2;
const quarterlyMetricPeriods = new Map([
  ["公管工资", "季度"],
  ["文娱工资", "季度"],
  ["卫生工资", "季度"],
  ["利佣净收入", "季度"],
  ["水利工资", "季度"],
  ["农业", "季度"],
  ["科技工资", "季度"],
  ["教育工资", "季度"],
  ["规下工业", "季度"],
  ["管道运输", "季度"],
  ["专包企业", "季度"],
]);
const metricBarLengthRatios = new Map(
  categories
    .flatMap((category) => category.metrics)
    .map((metric) => [
      metric.name,
      barLengthRatioFromStatuses(metric.months.map((month) => month.status)),
    ]),
);
const quarterlyBarProfiles = new Map(
  [...quarterlyMetricPeriods.entries()].map(([metricName, periodName]) => {
    const metric = categories
      .flatMap((category) => category.metrics)
      .find((item) => item.name === metricName);
    const quarterMonths = [
      { month: 3, name: "一季度" },
      { month: 6, name: "二季度" },
      { month: 9, name: "三季度" },
    ];
    const rawMonths = quarterMonths.map(({ month, name }) => {
      const status = monthStatusFromWorkbook(metricName, month);
      return {
        month,
        quarterName: name,
        status,
        weight: status.weight,
      };
    });
    const weightTotal = rawMonths.reduce((sum, month) => sum + month.weight, 0);
    const months = rawMonths.map(({ status, weight }, index) => {
      const normalizedWeight = weight / weightTotal;
      return {
        ...rawMonths[index],
        weight: normalizedWeight,
        value: metric.value * normalizedWeight,
        status,
      };
    });

    return [
      metricName,
      {
        months,
        lengthRatio: barLengthRatioFromStatuses(months.map((month) => month.status)),
      },
    ];
  }),
);
const chartStartAngle = -55;
const zoomLimits = {
  min: 0.65,
  max: 8,
};
const detailZoomThreshold = 1.45;
const radius = {
  center: 40,
  categoryInner: 64,
  categoryOuter: 94,
  stackTrackInner: 104,
  stackStart: 116,
  labelGap: 20,
};
let chartViewport = null;
let chartDetailLabels = [];
let chartDrag = null;
let chartDragged = false;

function polar(angleDeg, r) {
  const angle = ((angleDeg - 90) * Math.PI) / 180;
  return {
    x: center.x + Math.cos(angle) * r,
    y: center.y + Math.sin(angle) * r,
  };
}

function donutPath(startAngle, endAngle, innerRadius, outerRadius) {
  const startOuter = polar(startAngle, outerRadius);
  const endOuter = polar(endAngle, outerRadius);
  const startInner = polar(endAngle, innerRadius);
  const endInner = polar(startAngle, innerRadius);
  const largeArc = endAngle - startAngle > 180 ? 1 : 0;

  return [
    `M ${startOuter.x} ${startOuter.y}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${endOuter.x} ${endOuter.y}`,
    `L ${startInner.x} ${startInner.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${endInner.x} ${endInner.y}`,
    "Z",
  ].join(" ");
}

function makeSvg(tag, attributes = {}) {
  const node = document.createElementNS(SVG_NS, tag);
  Object.entries(attributes).forEach(([key, value]) => {
    node.setAttribute(key, String(value));
  });
  return node;
}

function addText(parent, text, x, y, className, attributes = {}) {
  const node = makeSvg("text", { x, y, class: className, ...attributes });
  node.textContent = text;
  parent.appendChild(node);
  return node;
}

function radialLabel(parent, text, angle, distance, className, options = {}) {
  const point = polar(angle, distance);
  const normalized = ((angle % 360) + 360) % 360;
  const outwardAngle = normalized - 90;
  const flip = normalized > 180;
  const rotation = flip ? outwardAngle + 180 : outwardAngle;
  const attributes = {
    transform: `rotate(${rotation} ${point.x} ${point.y})`,
    "text-anchor": flip ? "end" : "start",
    "dominant-baseline": "middle",
    ...options,
  };
  return addText(parent, text, point.x, point.y, className, attributes);
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function applyChartZoom() {
  if (!chartViewport) return;
  const { scale, x, y } = state.zoom;
  chartViewport.setAttribute("transform", `translate(${x} ${y}) scale(${scale})`);
  updateChartDetailLabels();
  chartZoomStatus.textContent = `${Math.round(scale * 100)}%`;
  chartCanvas.classList.add("is-pannable");
}

function updateChartDetailLabels() {
  const { scale } = state.zoom;
  const detailEnabled = scale >= detailZoomThreshold;
  let visibleCount = 0;

  chartDetailLabels.forEach(({ node, thickness }) => {
    node.style.fontSize = `${9.8 / scale}px`;
    node.style.strokeWidth = `${0.65 / scale}px`;
    node.style.letterSpacing = `${-0.08 / scale}px`;
    const visible = detailEnabled && thickness * scale >= 10.8;
    node.classList.toggle("is-visible", visible);
    node.setAttribute("aria-hidden", String(!visible));
    if (visible) visibleCount += 1;
  });

  chartZoomDetail.textContent = detailEnabled
    ? `灯色 ${visibleCount}/${chartDetailLabels.length}段`
    : "放大看灯色";
}

function defaultChartScale() {
  return window.matchMedia("(min-width: 821px)").matches ? 1.5 : 1;
}

function resetChartZoom() {
  const scale = defaultChartScale();
  state.zoom = {
    scale,
    x: center.x * (1 - scale),
    y: center.y * (1 - scale),
  };
  applyChartZoom();
}

function handleChartWheel(event) {
  event.preventDefault();

  const deltaUnit = event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? 100 : 1;
  const factor = Math.exp(-event.deltaY * deltaUnit * 0.0015);
  const oldScale = state.zoom.scale;
  const newScale = clamp(oldScale * factor, zoomLimits.min, zoomLimits.max);
  if (Math.abs(newScale - oldScale) < 0.001) return;

  const screenMatrix = chart.getScreenCTM();
  if (!screenMatrix) return;
  const pointer = new DOMPoint(event.clientX, event.clientY).matrixTransform(screenMatrix.inverse());
  const ratio = newScale / oldScale;

  state.zoom.scale = newScale;
  state.zoom.x = pointer.x - (pointer.x - state.zoom.x) * ratio;
  state.zoom.y = pointer.y - (pointer.y - state.zoom.y) * ratio;

  clampChartTranslation();
  applyChartZoom();
}

function chartPointFromEvent(event) {
  const screenMatrix = chart.getScreenCTM();
  if (!screenMatrix) return null;
  return new DOMPoint(event.clientX, event.clientY).matrixTransform(screenMatrix.inverse());
}

function clampChartTranslation() {
  const { scale } = state.zoom;
  const matrix = chart.getScreenCTM();
  if (!matrix) return;

  const canvasRect = chartCanvas.getBoundingClientRect();
  const inverse = matrix.inverse();
  const topLeft = new DOMPoint(canvasRect.left, canvasRect.top).matrixTransform(inverse);
  const bottomRight = new DOMPoint(canvasRect.right, canvasRect.bottom).matrixTransform(inverse);
  const minX = Math.min(topLeft.x, bottomRight.x);
  const maxX = Math.max(topLeft.x, bottomRight.x);
  const minY = Math.min(topLeft.y, bottomRight.y);
  const maxY = Math.max(topLeft.y, bottomRight.y);

  const visibleCenterX = clamp(center.x * scale + state.zoom.x, minX, maxX);
  const visibleCenterY = clamp(center.y * scale + state.zoom.y, minY, maxY);
  state.zoom.x = visibleCenterX - center.x * scale;
  state.zoom.y = visibleCenterY - center.y * scale;
}

function handleChartPointerDown(event) {
  if (event.button !== 0) return;
  const point = chartPointFromEvent(event);
  if (!point) return;

  chartDrag = {
    pointerId: event.pointerId,
    startX: point.x,
    startY: point.y,
    originX: state.zoom.x,
    originY: state.zoom.y,
    moved: false,
    target: event.target,
  };
  chartCanvas.setPointerCapture(event.pointerId);
}

function handleChartPointerMove(event) {
  if (!chartDrag || event.pointerId !== chartDrag.pointerId) return;
  const point = chartPointFromEvent(event);
  if (!point) return;

  const dx = point.x - chartDrag.startX;
  const dy = point.y - chartDrag.startY;
  if (Math.hypot(dx, dy) > 3) {
    chartDrag.moved = true;
    chartCanvas.classList.add("is-panning");
  }
  if (!chartDrag.moved) return;

  state.zoom.x = chartDrag.originX + dx;
  state.zoom.y = chartDrag.originY + dy;
  clampChartTranslation();
  applyChartZoom();
}

function handleChartPointerUp(event) {
  if (!chartDrag || event.pointerId !== chartDrag.pointerId) return;
  const { moved, target } = chartDrag;
  chartDragged = moved;
  chartDrag = null;
  chartCanvas.classList.remove("is-panning");
  if (chartCanvas.hasPointerCapture(event.pointerId)) {
    chartCanvas.releasePointerCapture(event.pointerId);
  }
  if (!moved) {
    const sector = target.closest?.(".sector-group");
    selectCategory(sector?.dataset.category || null);
  }
  if (chartDragged) window.setTimeout(() => {
    chartDragged = false;
  }, 0);
}

function metricMatchesFilter(metric) {
  const quarterPeriod = quarterlyMetricPeriods.get(metric.name);
  const months = quarterPeriod ? getQuarterSummary(metric, quarterPeriod).months : metric.months;
  const comparableMonths = months.filter((month) => month.month > 1);
  return state.filter === "全部" || comparableMonths.some((month) => month.status.key === state.filter);
}

function getQuarterSummary(metric, periodName) {
  const quarterIndex = ["一季度", "二季度", "三季度", "四季度"].indexOf(periodName);
  const profile = quarterlyBarProfiles.get(metric.name);
  const months = profile?.months || metric.months.slice(quarterIndex * 3, quarterIndex * 3 + 3);
  const statusCounts = new Map();
  months.forEach((month) => {
    statusCounts.set(month.status.key, (statusCounts.get(month.status.key) || 0) + 1);
  });
  const statusKey = [...statusCounts.entries()].sort((a, b) => b[1] - a[1])[0][0];
  const status = months.find((month) => month.status.key === statusKey).status;
  return {
    name: periodName,
    months,
    status,
    value: months.reduce((sum, month) => sum + month.value, 0),
  };
}

function categoryMatchesFilter(category) {
  return category.metrics.some(metricMatchesFilter);
}

function buildChart() {
  chart.replaceChildren();
  chartViewport = makeSvg("g", { class: "chart-viewport" });
  chartDetailLabels = [];

  const defs = makeSvg("defs");
  const filter = makeSvg("filter", { id: "softShadow", x: "-30%", y: "-30%", width: "160%", height: "160%" });
  filter.append(
    makeSvg("feDropShadow", {
      dx: 0,
      dy: 2,
      stdDeviation: 3,
      "flood-color": "#4b5d70",
      "flood-opacity": 0.1,
    }),
  );
  defs.appendChild(filter);
  chart.appendChild(defs);

  const background = makeSvg("g", { "aria-hidden": "true", "pointer-events": "none" });
  background.appendChild(
    makeSvg("circle", {
      cx: center.x,
      cy: center.y,
      r: 218,
      fill: "none",
      stroke: "rgba(170,177,184,0.18)",
      "stroke-dasharray": "2 7",
      "stroke-width": 1,
    }),
  );
  chartViewport.appendChild(background);

  const categoryLayer = makeSvg("g", { "aria-label": "部门分组" });
  const sectorLayer = makeSvg("g", { "aria-label": "指标扇区" });
  const labelLayer = makeSvg("g", { "aria-label": "指标名称与增速" });
  const centerLayer = makeSvg("g", { "aria-label": "GDP 中心" });

  let categoryCursor = chartStartAngle;
  categories.forEach((category, categoryIndex) => {
    const categorySpan = category.metrics.length * metricArc;
    const categoryStart = categoryCursor + gapDegrees / 2;
    const categoryEnd = categoryCursor + categorySpan - gapDegrees / 2;
    const categoryMid = (categoryStart + categoryEnd) / 2;
    categoryCursor += categorySpan;
    const matchesCategory = categoryMatchesFilter(category);
    const selected = state.selectedCategory === category.name;
    const dimmed = (state.filter !== "全部" && !matchesCategory) || (state.selectedCategory && !selected);

    const group = makeSvg("g", {
      class: `sector-group${dimmed ? " is-dimmed" : ""}${selected ? " is-selected" : ""}`,
      tabindex: "0",
      role: "button",
      "aria-label": `${category.name}，${category.metrics.length} 项指标`,
      "data-category": category.name,
    });

    const categoryPath = makeSvg("path", {
      d: donutPath(categoryStart, categoryEnd, radius.categoryInner, radius.categoryOuter),
      fill: selected ? "#d8e9ff" : "#e5e6e7",
      class: "sector-fill",
    });
    group.appendChild(categoryPath);

    const metrics = category.metrics;
    const metricSpan = (categoryEnd - categoryStart) / metrics.length;

    metrics.forEach((metric, metricIndex) => {
      const start = categoryStart + metricIndex * metricSpan + 0.5;
      const end = categoryStart + (metricIndex + 1) * metricSpan - 0.5;
      const mid = (start + end) / 2;
      const metricDimmed = !metricMatchesFilter(metric);
      const stackLength = uniformStackLength;
      const outerRadius = radius.stackStart + stackLength;

      const metricGroup = makeSvg("g", {
        class: metricDimmed ? "metric is-dimmed" : "metric",
        opacity: metricDimmed ? 0.16 : 1,
        "data-metric": metric.name,
      });

      metricGroup.appendChild(
        makeSvg("path", {
          d: donutPath(start, end, radius.stackTrackInner, radius.stackStart),
          fill: "#d5d7d9",
          opacity: 0.72,
          class: "stack-track",
        }),
      );

      const quarterPeriod = quarterlyMetricPeriods.get(metric.name);
      if (quarterPeriod) {
        const quarter = getQuarterSummary(metric, quarterPeriod);
        let stackCursor = radius.stackStart;
        quarter.months.forEach((month) => {
          const nextRadius = stackCursor + stackLength * month.weight;
          const monthDimmed = state.filter !== "全部" && month.status.key !== state.filter;
          const segmentPath = donutPath(start, end, stackCursor, nextRadius);
          const segmentGroup = makeSvg("g", { class: "stack-segment-group" });
          const hoverAngle = ((mid - 90) * Math.PI) / 180;
          segmentGroup.style.setProperty("--hover-x", `${Math.cos(hoverAngle) * 8}px`);
          segmentGroup.style.setProperty("--hover-y", `${Math.sin(hoverAngle) * 8}px`);
          segmentGroup.style.setProperty("--sector-mid-x", `${center.x + Math.cos(hoverAngle) * ((stackCursor + nextRadius) / 2)}px`);
          segmentGroup.style.setProperty("--sector-mid-y", `${center.y + Math.sin(hoverAngle) * ((stackCursor + nextRadius) / 2)}px`);
          const segment = makeSvg("path", {
            d: segmentPath,
            fill: month.status.color,
            opacity: monthDimmed ? 0.12 : 0.97,
            stroke: "rgba(250,250,250,0.94)",
            "stroke-width": 1.15,
            class: "stack-segment",
            "pointer-events": "none",
          });
          const hitArea = makeSvg("path", {
            d: segmentPath,
            fill: "transparent",
            class: "stack-segment-hit",
          });
          hitArea.setAttribute(
            "aria-label",
            `${metric.name} ${month.quarterName} ${month.status.light}`,
          );
          hitArea.addEventListener("pointerenter", (event) =>
            showQuarterMonthTooltip(event, category, metric, month),
          );
          hitArea.addEventListener("pointermove", moveTooltip);
          hitArea.addEventListener("pointerleave", hideTooltip);
          segmentGroup.append(segment, hitArea);
          metricGroup.appendChild(segmentGroup);

          const detailPoint = polar(mid, (stackCursor + nextRadius) / 2);
          const detailLabel = addText(
            metricGroup,
            month.quarterName,
            detailPoint.x,
            detailPoint.y,
            "month-value-label",
            {
              "aria-hidden": "true",
              "text-anchor": "middle",
              "dominant-baseline": "middle",
              transform: `rotate(${((mid % 360) + 360) % 360} ${detailPoint.x} ${detailPoint.y})`,
            },
          );
          chartDetailLabels.push({
            node: detailLabel,
            thickness: nextRadius - stackCursor,
            kind: "month",
          });

          stackCursor = nextRadius;
        });
      } else {
        let stackCursor = radius.stackStart;
        metric.months.forEach((month) => {
          const nextRadius = stackCursor + stackLength * month.weight;
          const monthDimmed = state.filter !== "全部" && month.status.key !== state.filter;
          const segmentPath = donutPath(start, end, stackCursor, nextRadius);
          const segmentGroup = makeSvg("g", { class: "stack-segment-group" });
          const hoverAngle = ((mid - 90) * Math.PI) / 180;
          segmentGroup.style.setProperty("--hover-x", `${Math.cos(hoverAngle) * 8}px`);
          segmentGroup.style.setProperty("--hover-y", `${Math.sin(hoverAngle) * 8}px`);
          segmentGroup.style.setProperty("--sector-mid-x", `${center.x + Math.cos(hoverAngle) * ((stackCursor + nextRadius) / 2)}px`);
          segmentGroup.style.setProperty("--sector-mid-y", `${center.y + Math.sin(hoverAngle) * ((stackCursor + nextRadius) / 2)}px`);
          const segment = makeSvg("path", {
            d: segmentPath,
            fill: month.status.color,
            opacity: monthDimmed ? 0.12 : 0.97,
            stroke: "rgba(250,250,250,0.9)",
            "stroke-width": 0.7,
            class: "stack-segment",
            "pointer-events": "none",
          });
          const hitArea = makeSvg("path", {
            d: segmentPath,
            fill: "transparent",
            class: "stack-segment-hit",
          });
          hitArea.setAttribute(
            "aria-label",
            `${metric.name} ${month.month}月 ${month.status.light}`,
          );
          hitArea.addEventListener("pointerenter", (event) =>
            showStackTooltip(event, category, metric, month),
          );
          hitArea.addEventListener("pointermove", moveTooltip);
          hitArea.addEventListener("pointerleave", hideTooltip);
          segmentGroup.append(segment, hitArea);
          metricGroup.appendChild(segmentGroup);

          const detailPoint = polar(mid, (stackCursor + nextRadius) / 2);
          const detailLabel = addText(
            metricGroup,
            `${month.month}月 ${month.status.light}`,
            detailPoint.x,
            detailPoint.y,
            "month-value-label",
            {
              "aria-hidden": "true",
              "text-anchor": "middle",
              "dominant-baseline": "middle",
              transform: `rotate(${((mid % 360) + 360) % 360} ${detailPoint.x} ${detailPoint.y})`,
            },
          );
          chartDetailLabels.push({
            node: detailLabel,
            thickness: nextRadius - stackCursor,
            kind: "month",
          });

          stackCursor = nextRadius;
        });
      }

      group.appendChild(metricGroup);

      const labelRadius = outerRadius + radius.labelGap;
      const label = radialLabel(
        labelLayer,
        `${metric.name}${quarterPeriod ? ` · ${quarterPeriod}` : ""}`,
        mid,
        labelRadius,
        "sector-label",
        { "data-label-metric": metric.name },
      );
      if (metricDimmed) label.setAttribute("opacity", "0.16");
    });

    group.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectCategory(category.name);
      }
    });
    sectorLayer.appendChild(group);

    const categoryLabelPoint = polar(categoryMid, 79);
    const normalized = ((categoryMid % 360) + 360) % 360;
    const flip = normalized > 90 && normalized < 270;
    addText(
      categoryLayer,
      category.short,
      categoryLabelPoint.x,
      categoryLabelPoint.y,
      "category-label",
      {
        "text-anchor": "middle",
        "dominant-baseline": "middle",
        transform: `rotate(${flip ? normalized + 180 : normalized} ${categoryLabelPoint.x} ${categoryLabelPoint.y})`,
      },
    );
  });

  centerLayer.append(
    makeSvg("circle", {
      cx: center.x,
      cy: center.y,
      r: radius.center + 8,
      fill: "#d9d9d9",
      stroke: "#c8c8c8",
      "stroke-width": 1,
      filter: "url(#softShadow)",
    }),
    makeSvg("circle", {
      cx: center.x,
      cy: center.y,
      r: radius.center,
      fill: "#fbfbfb",
      stroke: "#cfd1d3",
      "stroke-width": 1,
    }),
  );
  addText(centerLayer, "GDP", center.x, center.y + 2, "center-label", {
    "text-anchor": "middle",
    "dominant-baseline": "middle",
  });
  addText(centerLayer, "经济总量", center.x, center.y + 20, "center-sub", {
    "text-anchor": "middle",
    "dominant-baseline": "middle",
  });

  chartViewport.append(sectorLayer, categoryLayer, labelLayer, centerLayer);
  chart.appendChild(chartViewport);
  applyChartZoom();

  updateSummary();
  chartReset.hidden = !state.selectedCategory;
}

function showStackTooltip(event, category, metric, month) {
  tooltip.innerHTML = `<strong>${category.name} · ${metric.name}</strong>
    <span>${month.month} 月：<b>${month.status.light}</b></span><br />
    <span>状态：${month.status.key} · 径向第 ${month.month} 层</span>`;
  tooltip.classList.add("is-visible");
  moveTooltip(event);
}

function showQuarterMonthTooltip(event, category, metric, month) {
  tooltip.innerHTML = `<strong>${category.name} · ${metric.name}</strong>
    <span>${month.quarterName}：<b>${month.status.light}</b></span><br />
    <span>状态：${month.status.key} · 季度指标</span>`;
  tooltip.classList.add("is-visible");
  moveTooltip(event);
}

function moveTooltip(event) {
  if (!tooltip.classList.contains("is-visible")) return;
  const margin = 18;
  const rect = tooltip.getBoundingClientRect();
  let x = event.clientX + margin;
  let y = event.clientY + margin;
  if (x + rect.width > window.innerWidth - 12) x = event.clientX - rect.width - margin;
  if (y + rect.height > window.innerHeight - 12) y = event.clientY - rect.height - margin;
  tooltip.style.left = `${x}px`;
  tooltip.style.top = `${y}px`;
}

function hideTooltip() {
  tooltip.classList.remove("is-visible");
}

function selectCategory(name) {
  const nextCategory = state.selectedCategory === name ? null : name;
  if (state.selectedCategory === nextCategory) return;
  state.selectedCategory = nextCategory;
  hideTooltip();
  buildChart();
}

function updateSummary() {
  if (state.selectedCategory) {
    const category = categories.find((item) => item.name === state.selectedCategory);
    summary.innerHTML = `<span class="summary-kicker">部门聚焦</span>
      <strong>${category.name} · ${category.metrics.length} 项指标</strong>
      <span>1–9 月亮灯状态</span>
      <span class="month-strip" aria-label="1到9月状态">${category.months
        .map(
          (month) =>
            `<i style="--month-color:${month.status.color}" title="${month.month}月：${month.status.light}">${month.month}</i>`,
        )
        .join("")}</span>`;
    return;
  }

  const metricCount = categories.reduce((sum, category) => sum + category.metrics.length, 0);
  const filterText = state.filter === "全部" ? "全市 15 个部门" : `筛选：${state.filter}`;
  summary.innerHTML = `<span class="summary-kicker">当前视图</span>
    <strong>${filterText}</strong>
    <span>${metricCount} 项核心指标 · 来源：指标亮灯统计表</span>
    <span class="month-order"><b>${quarterlyMetricPeriods.size}</b> 项季度指标（Q1–Q3） · 其余按月份展示</span>`;
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("is-visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("is-visible"), 1900);
}

filterButton.addEventListener("click", () => {
  const open = filterMenu.classList.toggle("is-open");
  filterButton.setAttribute("aria-expanded", String(open));
});

filterMenu.addEventListener("click", (event) => {
  const option = event.target.closest("[data-filter]");
  if (!option) return;
  state.filter = option.dataset.filter;
  filterLabel.textContent = state.filter;
  filterMenu.querySelectorAll("[data-filter]").forEach((item) => {
    item.classList.toggle("is-selected", item === option);
  });
  filterMenu.classList.remove("is-open");
  filterButton.setAttribute("aria-expanded", "false");
  state.selectedCategory = null;
  buildChart();
});

document.addEventListener("click", (event) => {
  if (!event.target.closest(".header-tools")) {
    filterMenu.classList.remove("is-open");
    filterButton.setAttribute("aria-expanded", "false");
  }
});

chartReset.addEventListener("click", () => selectCategory(null));

document.querySelector("#collapseButton").addEventListener("click", () => {
  appShell.classList.toggle("sidebar-collapsed");
  const collapsed = appShell.classList.contains("sidebar-collapsed");
  document.querySelector("#collapseButton").setAttribute("aria-label", collapsed ? "展开侧边栏" : "收起侧边栏");
});

document.querySelectorAll(".nav-item, .side-item").forEach((button) => {
  button.addEventListener("click", () => {
    if (button.classList.contains("is-active")) return;
    showToast("演示版本：该入口暂未接入业务页面");
  });
});

document.querySelector(".report-tab").addEventListener("click", () => {
  showToast("汇报材料已准备，可在此接入导出或预览流程");
});

document.querySelector(".year-button").addEventListener("click", () => {
  showToast("演示版本：当前展示 2026 年 1-9 月数据");
});

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    filterMenu.classList.remove("is-open");
    filterButton.setAttribute("aria-expanded", "false");
    if (state.selectedCategory) selectCategory(null);
  }
});

resetChartZoom();
buildChart();

chartCanvas.addEventListener("wheel", handleChartWheel, { passive: false });
chartCanvas.addEventListener("pointerdown", handleChartPointerDown);
chartCanvas.addEventListener("pointermove", handleChartPointerMove);
chartCanvas.addEventListener("pointerup", handleChartPointerUp);
chartCanvas.addEventListener("pointercancel", handleChartPointerUp);
chart.addEventListener("click", (event) => {
  if (!chartDragged) return;
  event.preventDefault();
  event.stopPropagation();
}, true);
chartCanvas.addEventListener("dblclick", (event) => {
  event.preventDefault();
  resetChartZoom();
});
