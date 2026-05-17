"use strict";

/* =========================================================
   全局状态
========================================================= */

let currentMode = "xy";
let columns = [];
let dataRows = [];
let customTraceNames = [];
let customColors = [
    "#253f5f",
    "#b24a3b",
    "#2f6f4e",
    "#7a4fa3",
    "#c47f1f",
    "#3b7ea1"
];


/* =========================================================
   基础工具
========================================================= */

function $(id) {
    return document.getElementById(id);
}

function getValue(id, fallback = "") {
    const el = $(id);
    if (!el) return fallback;
    return el.value ?? fallback;
}

function getNumber(id, fallback = 0) {
    const n = Number(getValue(id));
    return Number.isFinite(n) ? n : fallback;
}

function setStatus(text) {
    const box = $("statusBox");
    if (box) box.textContent = text;
}

function selectedValues(id) {
    const el = $(id);
    if (!el) return [];
    return Array.from(el.selectedOptions).map(opt => opt.value);
}

function columnIndex(name) {
    return columns.indexOf(name);
}

function getColumnData(name) {
    const idx = columnIndex(name);
    if (idx < 0) return [];

    return dataRows
        .map(row => Number(row[idx]))
        .filter(v => Number.isFinite(v));
}

function getPairedColumnData(xName, yName) {
    const xi = columnIndex(xName);
    const yi = columnIndex(yName);

    const x = [];
    const y = [];

    if (xi < 0 || yi < 0) return { x, y };

    dataRows.forEach(row => {
        const xv = Number(row[xi]);
        const yv = Number(row[yi]);

        if (Number.isFinite(xv) && Number.isFinite(yv)) {
            x.push(xv);
            y.push(yv);
        }
    });

    return { x, y };
}

function sanitizeHtmlLabel(input) {
    if (!input) return "";

    let s = String(input);

    s = s
        .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, "")
        .replace(/on\w+="[^"]*"/gi, "")
        .replace(/on\w+='[^']*'/gi, "");

    const allowed = ["sub", "sup", "i", "b", "br"];

    s = s.replace(/<\/?([a-zA-Z0-9]+)(\s[^>]*)?>/g, (match, tag) => {
        tag = tag.toLowerCase();
        return allowed.includes(tag) ? match.replace(/\s[^>]*/g, "") : "";
    });

    return s;
}


/* =========================================================
   数据解析
========================================================= */

function parseTextData(text) {
    const lines = text
        .replace(/\r/g, "")
        .split("\n")
        .map(line => line.trim())
        .filter(Boolean);

    if (lines.length < 2) {
        throw new Error("数据行数不足。");
    }

    const first = lines[0];

    let delimiter = ",";
    if (first.includes("\t")) delimiter = "\t";
    else if (first.includes(";")) delimiter = ";";
    else if (first.includes(",")) delimiter = ",";
    else delimiter = /\s+/;

    const splitLine = line => {
        if (delimiter instanceof RegExp) {
            return line.split(delimiter).map(v => v.trim());
        }
        return line.split(delimiter).map(v => v.trim());
    };

    let raw = lines.map(splitLine);

    const firstRow = raw[0];
    const firstRowNumbers = firstRow.map(v => Number(v));
    const hasHeader = firstRowNumbers.some(v => !Number.isFinite(v));

    if (hasHeader) {
        columns = firstRow.map((v, i) => v || `列 ${i + 1}`);
        dataRows = raw.slice(1);
    } else {
        columns = firstRow.map((_, i) => `列 ${i + 1}`);
        dataRows = raw;
    }

    const width = columns.length;

    dataRows = dataRows
        .filter(row => row.length >= width)
        .map(row => row.slice(0, width));

    if (dataRows.length === 0) {
        throw new Error("没有有效数据。");
    }

    afterDataLoaded();
}

function parseExcelData(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    const json = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        defval: ""
    });

    const rows = json.filter(row => row.some(cell => String(cell).trim() !== ""));

    if (rows.length < 2) {
        throw new Error("Excel 数据行数不足。");
    }

    const firstRow = rows[0].map(v => String(v).trim());
    const firstRowNumbers = firstRow.map(v => Number(v));
    const hasHeader = firstRowNumbers.some(v => !Number.isFinite(v));

    if (hasHeader) {
        columns = firstRow.map((v, i) => v || `列 ${i + 1}`);
        dataRows = rows.slice(1);
    } else {
        columns = firstRow.map((_, i) => `列 ${i + 1}`);
        dataRows = rows;
    }

    const width = columns.length;

    dataRows = dataRows
        .filter(row => row.length >= width)
        .map(row => row.slice(0, width));

    afterDataLoaded();
}

function afterDataLoaded() {
    customTraceNames = [];

    populateAllSelects();
    updatePreviewTable();
    updateTraceNameInputs();
    updateColorPickers();

    setStatus(`已读取 ${dataRows.length} 行、${columns.length} 列数据。`);

    drawPlot();
}


/* =========================================================
   选择框和预览表
========================================================= */

function populateSelect(id, multiple = false) {
    const el = $(id);
    if (!el) return;

    el.innerHTML = "";

    columns.forEach((col, index) => {
        const opt = document.createElement("option");
        opt.value = col;
        opt.textContent = col;

        if (!multiple) {
            if (index === 0) opt.selected = true;
        } else {
            if (index > 0) opt.selected = true;
        }

        el.appendChild(opt);
    });
}

function populateAllSelects() {
    [
        "dualX",
        "fitX",
        "fitY",
        "heatX",
        "heatY",
        "heatZ",
        "contourX",
        "contourY",
        "contourZ",
        "threeX",
        "threeY",
        "threeZ"
    ].forEach(id => populateSelect(id, false));

    ["dualLeft", "dualRight", "statCols"].forEach(id => populateSelect(id, true));
}


/* =========================================================
   数据预览表
========================================================= */

function updatePreviewTable() {
    const table = $("previewTable");
    if (!table) return;

    table.innerHTML = "";

    if (!columns.length || !dataRows.length) {
        table.innerHTML = "<tbody><tr><td>暂无数据</td></tr></tbody>";
        return;
    }

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");

    columns.forEach(col => {
        const th = document.createElement("th");
        th.textContent = col;
        headRow.appendChild(th);
    });

    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    dataRows.slice(0, 20).forEach(row => {
        const tr = document.createElement("tr");

        columns.forEach((_, i) => {
            const td = document.createElement("td");
            td.textContent = row[i] ?? "";
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
}


/* =========================================================
   曲线命名器
========================================================= */

function getCurrentTraceBaseNames() {
    if (!columns.length) return [];

    if (currentMode === "xy") {
        return columns.slice(1);
    }

    if (currentMode === "dual") {
        return [
            ...selectedValues("dualLeft"),
            ...selectedValues("dualRight")
        ];
    }

    if (currentMode === "fit") {
        const y = getValue("fitY");
        return y ? [y, `${y} 拟合`] : [];
    }

    if (currentMode === "stat") {
        return selectedValues("statCols");
    }

    if (currentMode === "heatmap") {
        const z = getValue("heatZ");
        return z ? [z] : [];
    }

    if (currentMode === "contour") {
        const z = getValue("contourZ");
        return z ? [z] : [];
    }

    if (currentMode === "plot3d") {
        const z = getValue("threeZ");
        return z ? [z] : [];
    }

    return [];
}

function updateTraceNameInputs() {
    const box = $("traceNameList");
    if (!box) return;

    const names = getCurrentTraceBaseNames();

    box.innerHTML = "";

    if (names.length === 0) {
        const empty = document.createElement("div");
        empty.className = "trace-name-empty";
        empty.textContent = "读取数据后，将根据当前图形模式自动生成曲线命名框。";
        box.appendChild(empty);
        return;
    }

    while (customTraceNames.length < names.length) {
        customTraceNames.push("");
    }

    customTraceNames = customTraceNames.slice(0, names.length);

    names.forEach((name, index) => {
        const item = document.createElement("div");
        item.className = "trace-name-item";

        const label = document.createElement("div");
        label.className = "trace-name-label";
        label.textContent = `曲线 ${index + 1}`;

        const input = document.createElement("input");
        input.className = "trace-name-input";
        input.type = "text";
        input.placeholder = name
            ? `请输入曲线 ${index + 1} 名称，原始列：${name}`
            : `请输入曲线 ${index + 1} 名称`;

        input.value = customTraceNames[index] || "";

        input.addEventListener("input", () => {
            customTraceNames[index] = input.value;
            drawPlot();
        });

        item.appendChild(label);
        item.appendChild(input);
        box.appendChild(item);
    });
}

function getDisplayTraceName(index, fallback) {
    const name = customTraceNames[index];

    if (name && name.trim()) {
        return sanitizeHtmlLabel(name.trim());
    }

    return fallback || `曲线 ${index + 1}`;
}


/* =========================================================
   颜色选择器
========================================================= */

function updateColorPickers() {
    const box = $("colorPickerList");
    if (!box) return;

    const names = getCurrentTraceBaseNames();

    box.innerHTML = "";

    if (!names.length) {
        const empty = document.createElement("div");
        empty.className = "color-picker-empty";
        empty.textContent = "读取数据后，将根据曲线数量自动生成颜色选择器。";
        box.appendChild(empty);
        return;
    }

    while (customColors.length < names.length) {
        customColors.push(randomColor(customColors.length));
    }

    customColors = customColors.slice(0, names.length);

    names.forEach((name, index) => {
        const item = document.createElement("div");
        item.className = "color-picker-item";

        const label = document.createElement("span");
        label.textContent = `曲线 ${index + 1}`;

        const input = document.createElement("input");
        input.type = "color";
        input.value = customColors[index];

        input.addEventListener("input", () => {
            customColors[index] = input.value;
            updateCustomColorsInput();
            drawPlot();
        });

        item.appendChild(label);
        item.appendChild(input);
        box.appendChild(item);
    });

    updateCustomColorsInput();
}

function updateCustomColorsInput() {
    const el = $("customColors");
    if (el) el.value = customColors.join(",");
}

function randomColor(i) {
    const palette = [
        "#253f5f",
        "#b24a3b",
        "#2f6f4e",
        "#7a4fa3",
        "#c47f1f",
        "#3b7ea1",
        "#8c564b",
        "#e377c2",
        "#17becf",
        "#bcbd22"
    ];

    return palette[i % palette.length];
}

function getTraceColor(index) {
    return customColors[index % customColors.length] || randomColor(index);
}


/* =========================================================
   网格样式
========================================================= */

function applyGridStyleToAxis(axis, gridStyle, theme) {
    if (!axis) return;

    let gridColor = "#d7d0c4";
    let minorGridColor = "#ebe5da";

    if (theme === "dark") {
        gridColor = "#374151";
        minorGridColor = "#1f2937";
    }

    if (gridStyle === "none") {
        axis.showgrid = false;
        axis.zeroline = false;
        axis.minor = {
            showgrid: false
        };
        return;
    }

    if (gridStyle === "major") {
        axis.showgrid = true;
        axis.gridcolor = gridColor;
        axis.gridwidth = 1;
        axis.griddash = "solid";
        axis.zeroline = false;
        axis.minor = {
            showgrid: false
        };
        return;
    }

    if (gridStyle === "majorMinor") {
        axis.showgrid = true;
        axis.gridcolor = gridColor;
        axis.gridwidth = 1;
        axis.griddash = "solid";
        axis.zeroline = false;
        axis.minor = {
            showgrid: true,
            gridcolor: minorGridColor,
            gridwidth: 0.5
        };
        return;
    }

    if (gridStyle === "dash") {
        axis.showgrid = true;
        axis.gridcolor = gridColor;
        axis.gridwidth = 1;
        axis.griddash = "dash";
        axis.zeroline = false;
        axis.minor = {
            showgrid: false
        };
        return;
    }

    if (gridStyle === "paperLight") {
        axis.showgrid = true;
        axis.gridcolor = theme === "dark" ? "#263241" : "#eeeeee";
        axis.gridwidth = 0.8;
        axis.griddash = "solid";
        axis.zeroline = false;
        axis.minor = {
            showgrid: false
        };
    }
}


/* =========================================================
   布局
========================================================= */

function getBaseLayout() {
    const theme = getValue("theme", "paper");
    const fontSize = getNumber("fontSize", 15);
    const gridStyle = getValue("gridStyle", "none");

    let paperBg = "#ffffff";
    let plotBg = "#ffffff";
    let fontColor = "#111827";
    let axisColor = "#111827";

    if (theme === "dark") {
        paperBg = "#111827";
        plotBg = "#111827";
        fontColor = "#f9fafb";
        axisColor = "#f9fafb";
    }

    if (theme === "presentation") {
        fontColor = "#0f172a";
        axisColor = "#0f172a";
    }

    const layout = {
        title: {
            text: sanitizeHtmlLabel(getValue("plotTitle", "Scientific Plot")),
            font: {
                size: fontSize + 8,
                color: fontColor
            },
            x: 0.5
        },

        paper_bgcolor: paperBg,
        plot_bgcolor: plotBg,

        font: {
            family: "Times New Roman, SimSun, serif",
            size: fontSize,
            color: fontColor
        },

        margin: {
            l: 85,
            r: 85,
            t: 90,
            b: 80
        },

        xaxis: {
            title: {
                text: sanitizeHtmlLabel(getValue("xTitle", "X"))
            },
            type: getValue("xScale", "linear"),
            showline: true,
            linewidth: 1.5,
            linecolor: axisColor,
            mirror: true,
            zeroline: false
        },

        yaxis: {
            title: {
                text: sanitizeHtmlLabel(getValue("yTitle", "Y"))
            },
            type: getValue("yScale", "linear"),
            showline: true,
            linewidth: 1.5,
            linecolor: axisColor,
            mirror: true,
            zeroline: false
        },

        legend: {
            x: 1.02,
            y: 1,
            bgcolor: "rgba(255,255,255,0)"
        },

        hovermode: "closest"
    };

    applyGridStyleToAxis(layout.xaxis, gridStyle, theme);
    applyGridStyleToAxis(layout.yaxis, gridStyle, theme);

    return layout;
}

function getPlotConfig() {
    return {
        responsive: true,
        displaylogo: false,
        scrollZoom: true,
        modeBarButtonsToRemove: ["lasso2d", "select2d"]
    };
}


/* =========================================================
   trace 构建：二维图
========================================================= */

function buildXYTraces() {
    if (columns.length < 2) return [];

    const xName = columns[0];
    const yNames = columns.slice(1);
    const x = getColumnData(xName);

    const typeValue = getValue("xyType", "lines");
    const lineWidth = getNumber("lineWidth", 2);
    const markerSize = getNumber("markerSize", 7);

    return yNames.map((yName, index) => {
        const y = getColumnData(yName);
        const color = getTraceColor(index);

        if (typeValue === "bar") {
            return {
                type: "bar",
                x,
                y,
                name: getDisplayTraceName(index, yName),
                marker: {
                    color
                }
            };
        }

        return {
            type: "scatter",
            mode: typeValue,
            x,
            y,
            name: getDisplayTraceName(index, yName),
            line: {
                width: lineWidth,
                color
            },
            marker: {
                size: markerSize,
                color
            }
        };
    });
}


/* =========================================================
   trace 构建：双 Y 轴
========================================================= */

function buildDualTraces() {
    const xName = getValue("dualX") || columns[0];
    const leftNames = selectedValues("dualLeft");
    const rightNames = selectedValues("dualRight");

    const lineWidth = getNumber("lineWidth", 2);
    const markerSize = getNumber("markerSize", 7);

    const traces = [];
    let index = 0;

    leftNames.forEach(yName => {
        const pair = getPairedColumnData(xName, yName);
        const color = getTraceColor(index);

        traces.push({
            type: "scatter",
            mode: "lines+markers",
            x: pair.x,
            y: pair.y,
            name: getDisplayTraceName(index, yName),
            yaxis: "y",
            line: {
                width: lineWidth,
                color
            },
            marker: {
                size: markerSize,
                color
            }
        });

        index++;
    });

    rightNames.forEach(yName => {
        const pair = getPairedColumnData(xName, yName);
        const color = getTraceColor(index);

        traces.push({
            type: "scatter",
            mode: "lines+markers",
            x: pair.x,
            y: pair.y,
            name: getDisplayTraceName(index, yName),
            yaxis: "y2",
            line: {
                width: lineWidth,
                dash: "dash",
                color
            },
            marker: {
                size: markerSize,
                color
            }
        });

        index++;
    });

    return traces;
}


/* =========================================================
   多项式拟合
========================================================= */

function polynomialFit(x, y, degree) {
    const n = degree + 1;

    const A = Array.from({ length: n }, () => Array(n).fill(0));
    const B = Array(n).fill(0);

    for (let row = 0; row < n; row++) {
        for (let col = 0; col < n; col++) {
            A[row][col] = x.reduce((sum, xi) => sum + Math.pow(xi, row + col), 0);
        }

        B[row] = x.reduce((sum, xi, i) => sum + y[i] * Math.pow(xi, row), 0);
    }

    return solveLinearSystem(A, B);
}

function solveLinearSystem(A, B) {
    const n = B.length;

    for (let i = 0; i < n; i++) {
        let maxRow = i;

        for (let k = i + 1; k < n; k++) {
            if (Math.abs(A[k][i]) > Math.abs(A[maxRow][i])) {
                maxRow = k;
            }
        }

        [A[i], A[maxRow]] = [A[maxRow], A[i]];
        [B[i], B[maxRow]] = [B[maxRow], B[i]];

        const pivot = A[i][i];

        if (Math.abs(pivot) < 1e-12) continue;

        for (let j = i; j < n; j++) {
            A[i][j] /= pivot;
        }

        B[i] /= pivot;

        for (let k = 0; k < n; k++) {
            if (k === i) continue;

            const factor = A[k][i];

            for (let j = i; j < n; j++) {
                A[k][j] -= factor * A[i][j];
            }

            B[k] -= factor * B[i];
        }
    }

    return B;
}

function evalPolynomial(coeffs, x) {
    return coeffs.reduce((sum, c, i) => sum + c * Math.pow(x, i), 0);
}

function buildFitTraces() {
    const xName = getValue("fitX") || columns[0];
    const yName = getValue("fitY") || columns[1];
    const degree = getNumber("fitDegree", 1);

    const pair = getPairedColumnData(xName, yName);
    const x = pair.x;
    const y = pair.y;

    if (x.length < degree + 1) return [];

    const coeffs = polynomialFit(x, y, degree);

    const minX = Math.min(...x);
    const maxX = Math.max(...x);

    const fitX = [];
    const fitY = [];

    for (let i = 0; i < 200; i++) {
        const xv = minX + (maxX - minX) * i / 199;
        fitX.push(xv);
        fitY.push(evalPolynomial(coeffs, xv));
    }

    return [
        {
            type: "scatter",
            mode: "markers",
            x,
            y,
            name: getDisplayTraceName(0, yName),
            marker: {
                size: getNumber("markerSize", 7),
                color: getTraceColor(0)
            }
        },
        {
            type: "scatter",
            mode: "lines",
            x: fitX,
            y: fitY,
            name: getDisplayTraceName(1, `${degree} 阶拟合`),
            line: {
                width: getNumber("lineWidth", 2.5),
                color: getTraceColor(1)
            }
        }
    ];
}


/* =========================================================
   统计图
========================================================= */

function buildStatTraces() {
    const statType = getValue("statType", "histogram");
    const names = selectedValues("statCols");

    return names.map((name, index) => {
        const values = getColumnData(name);
        const color = getTraceColor(index);

        if (statType === "histogram") {
            return {
                type: "histogram",
                x: values,
                nbinsx: getNumber("bins", 30),
                name: getDisplayTraceName(index, name),
                marker: {
                    color
                },
                opacity: 0.75
            };
        }

        if (statType === "box") {
            return {
                type: "box",
                y: values,
                name: getDisplayTraceName(index, name),
                marker: {
                    color
                },
                boxpoints: "outliers"
            };
        }

        return {
            type: "violin",
            y: values,
            name: getDisplayTraceName(index, name),
            line: {
                color
            },
            meanline: {
                visible: true
            },
            box: {
                visible: true
            }
        };
    });
}


/* =========================================================
   热力图 / 等高线辅助
========================================================= */

function buildGridData(xName, yName, zName, size) {
    const xi = columnIndex(xName);
    const yi = columnIndex(yName);
    const zi = columnIndex(zName);

    const points = [];

    dataRows.forEach(row => {
        const x = Number(row[xi]);
        const y = Number(row[yi]);
        const z = Number(row[zi]);

        if (Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(z)) {
            points.push({ x, y, z });
        }
    });

    if (!points.length) {
        return {
            xGrid: [],
            yGrid: [],
            zGrid: []
        };
    }

    const minX = Math.min(...points.map(p => p.x));
    const maxX = Math.max(...points.map(p => p.x));
    const minY = Math.min(...points.map(p => p.y));
    const maxY = Math.max(...points.map(p => p.y));

    const xGrid = [];
    const yGrid = [];
    const zGrid = [];

    for (let i = 0; i < size; i++) {
        xGrid.push(minX + (maxX - minX) * i / (size - 1));
        yGrid.push(minY + (maxY - minY) * i / (size - 1));
    }

    for (let j = 0; j < size; j++) {
        const row = [];

        for (let i = 0; i < size; i++) {
            const gx = xGrid[i];
            const gy = yGrid[j];

            let best = points[0];
            let bestDist = Infinity;

            points.forEach(p => {
                const d = Math.pow(p.x - gx, 2) + Math.pow(p.y - gy, 2);
                if (d < bestDist) {
                    bestDist = d;
                    best = p;
                }
            });

            row.push(best.z);
        }

        zGrid.push(row);
    }

    return {
        xGrid,
        yGrid,
        zGrid
    };
}

function buildHeatmapTraces() {
    const xName = getValue("heatX") || columns[0];
    const yName = getValue("heatY") || columns[1];
    const zName = getValue("heatZ") || columns[2];

    const grid = buildGridData(xName, yName, zName, getNumber("gridSize", 45));

    return [
        {
            type: "heatmap",
            x: grid.xGrid,
            y: grid.yGrid,
            z: grid.zGrid,
            colorscale: "Viridis",
            colorbar: {
                title: sanitizeHtmlLabel(getValue("zTitle", zName))
            },
            name: getDisplayTraceName(0, zName)
        }
    ];
}

function buildContourTraces() {
    const xName = getValue("contourX") || columns[0];
    const yName = getValue("contourY") || columns[1];
    const zName = getValue("contourZ") || columns[2];

    const grid = buildGridData(xName, yName, zName, getNumber("contourGridSize", 45));

    return [
        {
            type: "contour",
            x: grid.xGrid,
            y: grid.yGrid,
            z: grid.zGrid,
            colorscale: "Viridis",
            contours: {
                coloring: "heatmap",
                showlines: true
            },
            line: {
                width: getNumber("contourLineWidth", 1.2)
            },
            colorbar: {
                title: sanitizeHtmlLabel(getValue("zTitle", zName))
            },
            name: getDisplayTraceName(0, zName)
        }
    ];
}


/* =========================================================
   三维图
========================================================= */

function build3DTraces() {
    const xName = getValue("threeX") || columns[0];
    const yName = getValue("threeY") || columns[1];
    const zName = getValue("threeZ") || columns[2];
    const type = getValue("threeType", "scatter3d");

    const xi = columnIndex(xName);
    const yi = columnIndex(yName);
    const zi = columnIndex(zName);

    const x = [];
    const y = [];
    const z = [];

    dataRows.forEach(row => {
        const xv = Number(row[xi]);
        const yv = Number(row[yi]);
        const zv = Number(row[zi]);

        if (Number.isFinite(xv) && Number.isFinite(yv) && Number.isFinite(zv)) {
            x.push(xv);
            y.push(yv);
            z.push(zv);
        }
    });

    if (type === "surface") {
        const grid = buildGridData(xName, yName, zName, 45);

        return [
            {
                type: "surface",
                x: grid.xGrid,
                y: grid.yGrid,
                z: grid.zGrid,
                colorscale: "Viridis",
                name: getDisplayTraceName(0, zName),
                colorbar: {
                    title: sanitizeHtmlLabel(getValue("zTitle", zName))
                }
            }
        ];
    }

    return [
        {
            type: "scatter3d",
            mode: "markers",
            x,
            y,
            z,
            name: getDisplayTraceName(0, zName),
            marker: {
                size: getNumber("markerSize", 5),
                color: z,
                colorscale: "Viridis",
                colorbar: {
                    title: sanitizeHtmlLabel(getValue("zTitle", zName))
                }
            }
        }
    ];
}


/* =========================================================
   绘图主函数
========================================================= */

function drawPlot() {
    const plot = $("plot");
    if (!plot) return;

    if (!columns.length || !dataRows.length) {
        Plotly.newPlot(plot, [], {
            title: {
                text: "请先读取数据"
            }
        }, getPlotConfig());

        return;
    }

    let traces = [];
    const layout = getBaseLayout();

    if (currentMode === "xy") {
        traces = buildXYTraces();
    }

    if (currentMode === "dual") {
        traces = buildDualTraces();

        layout.yaxis2 = {
            title: {
                text: sanitizeHtmlLabel(getValue("zTitle", "Right Y"))
            },
            overlaying: "y",
            side: "right",
            showline: true,
            linewidth: 1.5,
            linecolor: layout.font.color,
            zeroline: false
        };

        applyGridStyleToAxis(
            layout.yaxis2,
            getValue("gridStyle", "none"),
            getValue("theme", "paper")
        );
    }

    if (currentMode === "fit") {
        traces = buildFitTraces();
    }

    if (currentMode === "stat") {
        traces = buildStatTraces();
    }

    if (currentMode === "heatmap") {
        traces = buildHeatmapTraces();
    }

    if (currentMode === "contour") {
        traces = buildContourTraces();
    }

    if (currentMode === "plot3d") {
        traces = build3DTraces();

        const gridStyle = getValue("gridStyle", "none");
        const show3DGrid = gridStyle !== "none";
        const theme = getValue("theme", "paper");

        const gridColor = theme === "dark"
            ? "#374151"
            : gridStyle === "paperLight"
                ? "#eeeeee"
                : "#d7d0c4";

        layout.scene = {
            xaxis: {
                title: {
                    text: sanitizeHtmlLabel(getValue("xTitle", "X"))
                },
                showgrid: show3DGrid,
                gridcolor: gridColor
            },
            yaxis: {
                title: {
                    text: sanitizeHtmlLabel(getValue("yTitle", "Y"))
                },
                showgrid: show3DGrid,
                gridcolor: gridColor
            },
            zaxis: {
                title: {
                    text: sanitizeHtmlLabel(getValue("zTitle", "Z"))
                },
                showgrid: show3DGrid,
                gridcolor: gridColor
            }
        };
    }

    Plotly.newPlot(plot, traces, layout, getPlotConfig());
}


/* =========================================================
   导出
========================================================= */

function downloadPlot(format) {
    const plot = $("plot");
    if (!plot) return;

    Plotly.downloadImage(plot, {
        format,
        width: getNumber("exportWidth", 1600),
        height: getNumber("exportHeight", 1100),
        filename: "scientific_plot"
    });
}

function exportHtml() {
    const plot = $("plot");
    if (!plot || !plot.data || !plot.layout) return;

    const html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Scientific Plot Export</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"><\/script>
</head>
<body>
<div id="plot" style="width:100%;height:95vh;"></div>
<script>
const data = ${JSON.stringify(plot.data)};
const layout = ${JSON.stringify(plot.layout)};
Plotly.newPlot("plot", data, layout, {responsive:true, displaylogo:false});
<\/script>
</body>
</html>
`;

    const blob = new Blob([html], {
        type: "text/html;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.download = "scientific_plot.html";
    a.click();

    URL.revokeObjectURL(url);
}


/* =========================================================
   清空
========================================================= */

function clearAll() {
    columns = [];
    dataRows = [];
    customTraceNames = [];

    $("manualData").value = "";
    $("dataFile").value = "";

    updatePreviewTable();
    updateTraceNameInputs();
    updateColorPickers();

    setStatus("已清空数据。");

    Plotly.purge("plot");
}


/* =========================================================
   事件绑定
========================================================= */

function bindEvents() {
    const parseManualBtn = $("parseManualBtn");

    if (parseManualBtn) {
        parseManualBtn.addEventListener("click", () => {
            try {
                const text = getValue("manualData");

                if (!text.trim()) {
                    setStatus("请先粘贴数据。");
                    return;
                }

                parseTextData(text);
            } catch (err) {
                setStatus("解析失败：" + err.message);
            }
        });
    }

    const dataFile = $("dataFile");

    if (dataFile) {
        dataFile.addEventListener("change", async event => {
            const file = event.target.files[0];
            if (!file) return;

            try {
                const name = file.name.toLowerCase();

                if (name.endsWith(".xls") || name.endsWith(".xlsx")) {
                    const buffer = await file.arrayBuffer();
                    parseExcelData(buffer);
                } else {
                    const text = await file.text();
                    parseTextData(text);
                }
            } catch (err) {
                setStatus("文件读取失败：" + err.message);
            }
        });
    }

    document.querySelectorAll(".tab").forEach(tab => {
        tab.addEventListener("click", () => {
            document.querySelectorAll(".tab").forEach(t => {
                t.classList.remove("active");
            });

            document.querySelectorAll(".mode-panel").forEach(panel => {
                panel.classList.remove("active");
            });

            tab.classList.add("active");

            currentMode = tab.dataset.mode;

            const panel = $("mode-" + currentMode);
            if (panel) panel.classList.add("active");

            customTraceNames = [];
            updateTraceNameInputs();
            updateColorPickers();
            drawPlot();
        });
    });

    const drawBtn = $("drawBtn");
    if (drawBtn) drawBtn.addEventListener("click", drawPlot);

    const downloadPngBtn = $("downloadPngBtn");
    if (downloadPngBtn) {
        downloadPngBtn.addEventListener("click", () => downloadPlot("png"));
    }

    const downloadSvgBtn = $("downloadSvgBtn");
    if (downloadSvgBtn) {
        downloadSvgBtn.addEventListener("click", () => downloadPlot("svg"));
    }

    const exportHtmlBtn = $("exportHtmlBtn");
    if (exportHtmlBtn) {
        exportHtmlBtn.addEventListener("click", exportHtml);
    }

    const clearBtn = $("clearBtn");
    if (clearBtn) clearBtn.addEventListener("click", clearAll);

    const redrawIds = [
        "xyType",
        "dualX",
        "dualLeft",
        "dualRight",
        "fitX",
        "fitY",
        "fitDegree",
        "statCols",
        "statType",
        "bins",
        "heatX",
        "heatY",
        "heatZ",
        "gridSize",
        "contourX",
        "contourY",
        "contourZ",
        "contourGridSize",
        "contourLineWidth",
        "threeX",
        "threeY",
        "threeZ",
        "threeType",
        "plotTitle",
        "xTitle",
        "yTitle",
        "zTitle",
        "xScale",
        "yScale",
        "gridStyle",
        "lineWidth",
        "markerSize",
        "fontSize",
        "theme",
        "exportWidth",
        "exportHeight"
    ];

    redrawIds.forEach(id => {
        const el = $(id);
        if (!el) return;

        el.addEventListener("change", () => {
            updateTraceNameInputs();
            updateColorPickers();
            drawPlot();
        });

        el.addEventListener("input", () => {
            if (
                id === "plotTitle" ||
                id === "xTitle" ||
                id === "yTitle" ||
                id === "zTitle" ||
                id === "lineWidth" ||
                id === "markerSize" ||
                id === "fontSize"
            ) {
                drawPlot();
            }
        });
    });
}


/* =========================================================
   初始化
========================================================= */

document.addEventListener("DOMContentLoaded", () => {
    bindEvents();
    updatePreviewTable();
    updateTraceNameInputs();
    updateColorPickers();

    Plotly.newPlot(
        "plot",
        [],
        {
            title: {
                text: "请上传或粘贴数据"
            },
            paper_bgcolor: "#ffffff",
            plot_bgcolor: "#ffffff",
            font: {
                family: "Times New Roman, SimSun, serif"
            }
        },
        getPlotConfig()
    );
});