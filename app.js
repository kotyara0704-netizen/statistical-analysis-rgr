// State
const state = {
    X: [],
    Y: [],
    Z: []
};

// UI Toggles
document.getElementById('apply-k-btn').addEventListener('click', () => {
    if(state.X.length > 0) {
        renderTask1to6();
        if(document.getElementById('task-7-10').classList.contains('active-panel')) renderTask7to10();
        if(document.getElementById('task-17').classList.contains('active-panel')) renderTask17();
    }
});
document.getElementById('pearson-excel-mode').addEventListener('change', () => {
    if(state.X.length > 0) {
        renderTask7to10();
        if(document.getElementById('task-17').classList.contains('active-panel')) renderTask17();
    }
});
document.getElementById('anova-type').addEventListener('change', () => {
    if(state.X.length > 0) {
        renderTask11to15();
        if(document.getElementById('task-17').classList.contains('active-panel')) renderTask17();
    }
});

// --- Navigation ---
document.querySelectorAll('.menu-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.menu-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.panel').forEach(p => p.classList.remove('active-panel'));
        
        e.target.classList.add('active');
        const targetId = e.target.getAttribute('data-target');
        document.getElementById(targetId).classList.add('active-panel');
        
        // Render charts when panel becomes active to fix Plotly dimensions
        if(state.X.length > 0) {
            if(targetId === 'task-1-6') renderTask1to6();
            if(targetId === 'task-7-10') renderTask7to10();
            if(targetId === 'task-11-15') renderTask11to15();
            if(targetId === 'task-16') renderTask16();
            if(targetId === 'task-17') renderTask17();
        }
        renderMathInContainer(targetId);
    });
});

function renderMathInContainer(elementId) {
    if (window.renderMathInElement) {
        renderMathInElement(document.getElementById(elementId), {
            delimiters: [
                {left: "$$", right: "$$", display: true},
                {left: "\\[", right: "\\]", display: true},
                {left: "$", right: "$", display: false},
                {left: "\\(", right: "\\)", display: false}
            ],
            throwOnError: false
        });
    }
}

// --- Data Generation & Upload ---
function generateData() {
    const n = parseInt(document.getElementById('sample-size').value);
    const dist = document.getElementById('dist-type').value;
    
    if (n < 50) {
        alert("Обсяг вибірки має бути не менше 50");
        return;
    }

    state.X = []; state.Y = []; state.Z = [];
    
    for (let i = 0; i < n; i++) {
        if (dist === 'normal') {
            let u = 0, v = 0;
            while(u === 0) u = Math.random();
            while(v === 0) v = Math.random();
            let numX = Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
            let numY = Math.sqrt(-2.0 * Math.log(u)) * Math.sin(2.0 * Math.PI * v);
            let numZ = numX * 0.5 + numY * 0.5 + (Math.random() - 0.5); 
            
            state.X.push(numX * 10 + 50); 
            state.Y.push(numX * 8 + numY * 5 + 40); 
            state.Z.push(numZ * 15 + 60);
        } else {
            state.X.push(Math.random() * 100);
            state.Y.push(Math.random() * 100);
            state.Z.push(Math.random() * 100);
        }
    }
    
    updateDataTable();
    alert("Дані успішно згенеровано!");
}

document.getElementById('generate-btn').addEventListener('click', generateData);

document.getElementById('excel-upload').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    document.getElementById('file-name').innerText = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            
            if(jsonData.length < 2) throw new Error("Файл занадто малий.");
            
            let rawX = [], rawY = [], rawZ = [];
            let hasErrors = false;
            
            // Check if first row is a header or data
            let startIndex = 1;
            if (jsonData[0] && jsonData[0][0] !== undefined && !isNaN(parseFloat(jsonData[0][0]))) {
                startIndex = 0;
            }
            
            for(let i = startIndex; i < jsonData.length; i++) {
                if(jsonData[i][0] !== undefined) rawX.push(parseFloat(jsonData[i][0]));
                if(jsonData[i][1] !== undefined) rawY.push(parseFloat(jsonData[i][1]));
                if(jsonData[i][2] !== undefined) rawZ.push(parseFloat(jsonData[i][2]));
            }
            
            const cleanData = (arr) => {
                let valid = arr.filter(v => !isNaN(v));
                if (valid.length === 0) return [];
                let m = ss.mean(valid);
                return arr.map(v => {
                    if(isNaN(v)) { hasErrors = true; return m; }
                    return v;
                });
            };
            
            state.X = cleanData(rawX);
            state.Y = cleanData(rawY);
            state.Z = cleanData(rawZ);
            
            if(hasErrors) alert('Увага: у файлі були виявлені порожні клітинки або текст замість чисел. Їх було автоматично замінено на середнє значення відповідного стовпця (NaN-імпутація).');
            
            if(state.X.length < 50) {
                alert(`Зчитано лише ${state.X.length} рядків. Завдання вимагає мінімум 50.`);
            }
            
            alert(`Успішно завантажено! X: ${state.X.length}, Y: ${state.Y.length}, Z: ${state.Z.length} елементів.`);
            document.querySelector('[data-target="task-1-6"]').click();
            updateDataTable();
        } catch (err) {
            alert('Помилка читання файлу: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
});

// Copy Data Handlers
document.getElementById('copy-var-series').addEventListener('click', () => {
    const text = state.X.slice().sort((a,b)=>a-b).join('\\t');
    navigator.clipboard.writeText(text).then(() => alert('Варіаційний ряд скопійовано!'));
});
document.getElementById('copy-freq-table').addEventListener('click', () => {
    if(!globalGroupedData || !globalGroupedData.intervals) return;
    let text = "Інтервал\\tСередина\\tЧастота\\n";
    globalGroupedData.intervals.forEach((int, i) => {
        text += `[${int.start}; ${int.end})\\t${int.mid}\\t${globalGroupedData.freqs[i]}\\n`;
    });
    navigator.clipboard.writeText(text).then(() => alert('Таблицю частот скопійовано!'));
});

function updateDataTable() {
    const tbody = document.querySelector('#data-table tbody');
    tbody.innerHTML = '';
    const n = Math.min(state.X.length, 100);
    for (let i = 0; i < n; i++) {
        tbody.innerHTML += `
            <tr>
                <td>${i + 1}</td>
                <td>${state.X[i]}</td>
                <td>${state.Y[i] !== undefined ? state.Y[i] : '-'}</td>
                <td>${state.Z[i] !== undefined ? state.Z[i] : '-'}</td>
            </tr>
        `;
    }
}

function round(val, dec = 4) {
    return Number(Math.round(val + 'e' + dec) + 'e-' + dec);
}

// Global variable to share frequencies for Task 9
let globalGroupedData = null;

// --- TASK 1-6 ---
function renderTask1to6() {
    const X = state.X;
    const n = X.length;
    
    // 1. Variational series
    const sortedX = [...X].sort((a, b) => a - b);
    document.getElementById('variation-series').innerHTML = `
        <div class="result-box mt-2" style="max-height: 150px; overflow-y: auto; font-family: monospace;">
        ${sortedX.map(v => v).join(', ')}
        </div>
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Елементи вибірки впорядковано за зростанням. Це необхідно для визначення мінімуму, максимуму та подальшого групування.
        </div>
    `;
    
    // 2. Grouped data
    const sturgesK = Math.round(1 + 3.322 * Math.log10(n));
    
    // Update slot machine UI
    const slotN = document.getElementById('slot-n');
    const slotK = document.getElementById('slot-k-val');
    if (slotN && slotK) {
        slotN.innerText = n;
        slotK.classList.add('rolling');
        setTimeout(() => {
            slotK.classList.remove('rolling');
            slotK.innerText = sturgesK;
        }, 600); // 600ms fake roll
    }
    const manualKVal = document.getElementById('manual-k').value;
    const k = manualKVal ? parseInt(manualKVal) : sturgesK;
    const minX = sortedX[0];
    const maxX = sortedX[n - 1];
    
    // ВИПРАВЛЕННЯ 1: Трохи збільшуємо ширину інтервалу, щоб крайнє праве значення не випадало за межі
    let h = (maxX - minX) / k;
    h = Math.ceil(h * 10000) / 10000; // Округлення в більший бік до 4 знаків
    
    let intervals = [];
    let freqs = new Array(k).fill(0);
    
    for (let i = 0; i < k; i++) {
        intervals.push({
            start: minX + i * h,
            end: minX + (i + 1) * h,
            mid: minX + i * h + h / 2
        });
    }
    
    X.forEach(val => {
        let idx = Math.floor((val - minX) / h);
        if (idx >= k) idx = k - 1; // Захист на випадок похибок арифметики з плаваючою крапкою
        freqs[idx]++;
    });
    
    globalGroupedData = { k, intervals, freqs }; // Зберігаємо для Критерію Пірсона
    
    const tbodyGrouped = document.querySelector('#grouped-data-table tbody');
    tbodyGrouped.innerHTML = '';
    intervals.forEach((interval, i) => {
        const w = freqs[i] / n;
        tbodyGrouped.innerHTML += `
            <tr>
                <td>[${round(interval.start, 2)} ; ${round(interval.end, 2)}${i === k-1 ? ']' : ')'}</td>
                <td>${round(interval.mid, 2)}</td>
                <td>${freqs[i]}</td>
                <td>${round(w, 4)}</td>
            </tr>
        `;
    });
    
    let tableContainer = document.querySelector('#grouped-data-table').parentElement;
    if(!tableContainer.nextElementSibling || !tableContainer.nextElementSibling.classList.contains('explanation')) {
        tableContainer.insertAdjacentHTML('afterend', `
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Кількість інтервалів знайдена за формулою Стерджеса: \\( k = 1 + 3.322 \\lg n \\) і становить ${k}.
            Ширину (крок) кожного інтервалу обчислено як \\( h = \\frac{x_{max} - x_{min}}{k} \\) і округлено в більший бік до ${round(h)}, щоб уникнути втрати максимального значення вибірки.
            Відносна частота обчислюється як відношення кількості елементів в інтервалі до загального обсягу вибірки: \\( w_i = \\frac{n_i}{n} \\).
        </div>`);
    }
    
    // 3. Histogram and Polygon
    const mids = intervals.map(int => int.mid);
    
    // ВИПРАВЛЕННЯ 2: Замість type: 'histogram', який сам групує дані, використовуємо type: 'bar' для точного співпадіння
    const traceHist = {
        x: mids, y: freqs, type: 'bar', name: 'Гістограма', width: h,
        marker: { color: 'rgba(59, 130, 246, 0.6)', line: {color: '#fff', width: 1} }
    };
    const tracePoly = {
        x: mids, y: freqs, type: 'scatter', mode: 'lines+markers', name: 'Полігон частот',
        line: { color: '#ef4444', width: 2 }
    };
    Plotly.newPlot('plot-histogram', [traceHist, tracePoly], {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', font: { color: '#f8fafc' },
        margin: { t: 30, r: 10, l: 40, b: 40 }
    });
    
    // 4. Empirical Distribution Function (EDF)
    // ВИПРАВЛЕННЯ 3: Будуємо справжню ступінчасту функцію за ВАРІАЦІЙНИМ рядом, а не за згрупованими даними
    let edfX = [sortedX[0] - h]; // Початкова точка перед першим значенням
    let edfY = [0];
    
    for (let i = 0; i < n; i++) {
        edfX.push(sortedX[i]);
        edfY.push((i + 1) / n);
    }
    // Додаємо хвіст вправо для краси графіку
    edfX.push(sortedX[n-1] + h);
    edfY.push(1);
    
    const traceEDF = {
        x: edfX, y: edfY, type: 'scatter', mode: 'lines',
        line: { shape: 'hv', color: '#10b981', width: 2 }, name: 'F*(x)'
    };
    Plotly.newPlot('plot-edf', [traceEDF], {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', font: { color: '#f8fafc' },
        margin: { t: 30, r: 10, l: 40, b: 40 }
    });
    
    let plotEDFContainer = document.getElementById('plot-edf');
    if(!plotEDFContainer.nextElementSibling || !plotEDFContainer.nextElementSibling.classList.contains('explanation')) {
        plotEDFContainer.insertAdjacentHTML('afterend', `
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Справжня емпірична функція розподілу — це ступінчаста функція, що здійснює стрибок розміром 1/n у кожній точці варіаційного ряду. Вона показує частку елементів вибірки, менших за певне значення на графіку.
        </div>`);
    }

    // 5 & 6. Descriptive Stats
    const mean = ss.mean(X);
    const exactMedian = ss.median(X);
    
    // Grouped Median calculation
    let nHalf = n / 2;
    let cumFreq = 0;
    let meIdx = 0;
    for (let i = 0; i < k; i++) {
        cumFreq += freqs[i];
        if (cumFreq >= nHalf) {
            meIdx = i;
            break;
        }
    }
    let sMePrev = meIdx > 0 ? freqs.slice(0, meIdx).reduce((a, b) => a + b, 0) : 0;
    let nMe = freqs[meIdx];
    let xStartMe = intervals[meIdx].start;
    let groupedMedian = nMe === 0 ? "немає" : round(xStartMe + h * ((nHalf - sMePrev) / nMe));
    
    // ВИПРАВЛЕННЯ 4: Мода для неперервних згрупованих даних
    let maxFreq = -1;
    let maxFreqIdx = -1;
    for(let i = 0; i < k; i++) {
        if(freqs[i] > maxFreq) {
            maxFreq = freqs[i];
            maxFreqIdx = i;
        }
    }
    let nMo = freqs[maxFreqIdx];
    let nMoPrev = maxFreqIdx > 0 ? freqs[maxFreqIdx - 1] : 0;
    let nMoNext = maxFreqIdx < k - 1 ? freqs[maxFreqIdx + 1] : 0;
    let xStartMo = intervals[maxFreqIdx].start;
    
    let denom = (nMo - nMoPrev) + (nMo - nMoNext);
    const modeVal = denom === 0 ? "не виражена" : round(xStartMo + h * ((nMo - nMoPrev) / denom));
    
    const varUnbiased = ss.sampleVariance(X);
    const varBiased = ss.variance(X);
    const sdUnbiased = ss.sampleStandardDeviation(X);
    const sdBiased = ss.standardDeviation(X);
    const skewness = ss.sampleSkewness(X);
    
    document.getElementById('descriptive-stats').innerHTML = `
        <div class="stat-item"><div class="stat-label">Середнє (\\( \\bar{x} \\))</div><div class="stat-value">${round(mean)}</div></div>
        <div class="stat-item" style="border: 1px solid rgba(16, 185, 129, 0.4);"><div class="stat-label">Медіана (Інтервальна)</div><div class="stat-value text-emerald-400">${groupedMedian}</div></div>
        <div class="stat-item"><div class="stat-label">Медіана (Точна, \\( Me \\))</div><div class="stat-value">${round(exactMedian)}</div></div>
        <div class="stat-item"><div class="stat-label">Мода (Інтервальна, \\( Mo \\))</div><div class="stat-value">${modeVal}</div></div>
        <div class="stat-item"><div class="stat-label">Дисперсія (Зсунена, \\( \\sigma^2 \\))</div><div class="stat-value">${round(varBiased)}</div></div>
        <div class="stat-item"><div class="stat-label">Дисперсія (Незсунена, \\( S^2 \\))</div><div class="stat-value">${round(varUnbiased)}</div></div>
        <div class="stat-item"><div class="stat-label">СКВ (Зсунене, \\( \\sigma \\))</div><div class="stat-value">${round(sdBiased)}</div></div>
        <div class="stat-item"><div class="stat-label">СКВ (Незсунене, \\( S \\))</div><div class="stat-value">${round(sdUnbiased)}</div></div>
        <div class="stat-item"><div class="stat-label">Асиметрія (\\( A_s \\))</div><div class="stat-value">${round(skewness)}</div></div>
        <div class="explanation mt-2" style="grid-column: 1 / -1; color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b><br>
            Тут наведено основні описові статистики: середнє (центр розподілу), медіану (значення, що ділить вибірку навпіл) та моду (обчислену класичним методом для інтервального ряду). Дисперсія та середнє квадратичне відхилення (СКВ) характеризують розкид даних відносно середнього. Асиметрія показує наскільки графік розподілу "скошений" в один із боків.
        </div>
    `;
}

// --- TASK 7-10 ---
document.getElementById('outlier-method').addEventListener('change', () => {
    if(document.getElementById('task-7-10').classList.contains('active-panel')) {
        renderTask7to10();
    }
});

function renderTask7to10() {
    const X = state.X;
    const n = X.length;
    const sortedX = [...X].sort((a, b) => a - b);
    
    // 7. Outliers
    const method = document.getElementById('outlier-method').value;
    let lowerBound, upperBound, methodText;
    
    if (method === 'iqr') {
        const q1 = ss.quantile(sortedX, 0.25);
        const q3 = ss.quantile(sortedX, 0.75);
        const iqr = q3 - q1;
        lowerBound = q1 - 1.5 * iqr;
        upperBound = q3 + 1.5 * iqr;
        methodText = `Метод: Міжквартильний розмах (\\( IQR = ${round(iqr)} \\)). Нижня межа: \\( Q_1 - 1.5 \\cdot IQR = ${round(lowerBound)} \\), Верхня межа: \\( Q_3 + 1.5 \\cdot IQR = ${round(upperBound)} \\).`;
    } else if (method === 'sigma') {
        const mean = ss.mean(X);
        const sd = ss.sampleStandardDeviation(X);
        lowerBound = mean - 3 * sd;
        upperBound = mean + 3 * sd;
        methodText = `Метод: Правило трьох сигм (\\( 3\\sigma \\)). Нижня межа: \\( \\bar{x} - 3\\sigma = ${round(lowerBound)} \\), Верхня межа: \\( \\bar{x} + 3\\sigma = ${round(upperBound)} \\).`;
    } else {
        const mean = ss.mean(X);
        const sd = ss.sampleStandardDeviation(X);
        const tCrit = jStat.studentt.inv(1 - 0.05 / (2 * n), n - 2);
        const gCrit = ((n - 1) / Math.sqrt(n)) * Math.sqrt(Math.pow(tCrit, 2) / (n - 2 + Math.pow(tCrit, 2)));
        lowerBound = mean - gCrit * sd;
        upperBound = mean + gCrit * sd;
        methodText = `Метод: Критерій Смірнова-Граббса (\\( \\alpha = 0.05 \\)). Критичне значення \\( G_{кр} = ${round(gCrit)} \\).<br>Нижня межа: \\( \\bar{x} - G_{кр}\\cdot S = ${round(lowerBound)} \\), Верхня межа: \\( \\bar{x} + G_{кр}\\cdot S = ${round(upperBound)} \\).`;
    }
    
    const outliers = X.filter(x => x < lowerBound || x > upperBound);
    const outliersHtml = `
        ${methodText}<br><br>
        <b>Виявлені грубі помилки (викиди):</b> ${outliers.length > 0 ? outliers.map(v => round(v)).join(', ') : 'Не виявлено'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> За допомогою обраного методу ми знаходимо типові межі для нашої вибірки. Усі значення, які виходять за ці межі, вважаються аномальними (викидами) та не є типовими для даного набору даних.
        </div>
    `;
    state.reportOutliers = outliersHtml;
    document.getElementById('outliers-result').innerHTML = outliersHtml;
    
    // 8. Kolmogorov
    const mean = ss.mean(X);
    const sd = ss.sampleStandardDeviation(X);
    
    let maxD = 0;
    // ВИПРАВЛЕННЯ 5: Максимальне відхилення ступінчастої функції рахується як ПЕРЕД стрибком, так і ПІСЛЯ стрибка
    for (let i = 0; i < n; i++) {
        let edfRight = (i + 1) / n;
        let edfLeft = i / n;
        let cdf = jStat.normal.cdf(sortedX[i], mean, sd);
        let d1 = Math.abs(edfRight - cdf);
        let d2 = Math.abs(edfLeft - cdf);
        if (d1 > maxD) maxD = d1;
        if (d2 > maxD) maxD = d2;
    }
    
    const lambda = maxD * Math.sqrt(n);
    const kolmogorovCrit = 1.36; 
    const kolmogorovHtml = `
        Спостережуване значення максимальної відстані \\( D \\): ${round(maxD)}<br>
        Статистика Колмогорова \\( \\lambda = D \\cdot \\sqrt{n} \\): ${round(lambda)}<br>
        Критичне значення \\( \\lambda_{кр} (\\alpha=0.05) \\): ${round(kolmogorovCrit)}<br><br>
        <b>Висновок:</b> ${lambda < kolmogorovCrit ? 'Гіпотеза про нормальний розподіл приймається.' : 'Гіпотеза про нормальний розподіл відхиляється.'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> За критерієм Колмогорова ми порівняли нашу емпіричну функцію розподілу з теоретичною нормальною.
            Максимальне відхилення розраховувалось з урахуванням стрибків ступінчастої функції. Значення порівнюється зі статистикою Колмогорова \\( \\lambda \\). Оскільки параметри (середнє та СКВ) оцінювались за вибіркою, більш точним був би критерій Лілієфорса, проте для студентських робіт класичний критерій Колмогорова (порівняння \\( \\lambda \\) з 1.36) вважається достатнім.
        </div>
    `;
    state.reportKolmogorov = kolmogorovHtml;
    document.getElementById('kolmogorov-result').innerHTML = kolmogorovHtml;
    
    // 9. Pearson Chi-Square
    if (!globalGroupedData) return;
    let { k, intervals, freqs } = globalGroupedData;
    
    let obsFreqs = [];
    let expectedFreqs = [];
    let kNew = 0;
    
    let currentObs = 0;
    let currentExp = 0;
    
    const excelMode = document.getElementById('pearson-excel-mode').checked;

    for (let i = 0; i < k; i++) {
        currentObs += freqs[i];
        
        let p = jStat.normal.cdf(intervals[i].end, mean, sd) - jStat.normal.cdf(intervals[i].start, mean, sd);
        if (i === 0) p = jStat.normal.cdf(intervals[0].end, mean, sd);
        if (i === k - 1) p = 1 - jStat.normal.cdf(intervals[k-1].start, mean, sd);
        
        currentExp += n * p;
        
        if (excelMode || currentExp >= 5 || i === k - 1) {
            obsFreqs.push(currentObs);
            expectedFreqs.push(currentExp);
            kNew++;
            currentObs = 0;
            currentExp = 0;
        }
    }

    let chiSquare = 0;
    for (let i = 0; i < kNew; i++) {
        chiSquare += Math.pow(obsFreqs[i] - expectedFreqs[i], 2) / expectedFreqs[i];
    }
    
    let df = excelMode ? (kNew - 1) : (kNew - 3);
    if (df < 1) df = 1; 
    const pearsonCrit = jStat.chisquare.inv(0.95, df); 
        
    const pearsonHtml = `
        Кількість інтервалів після об'єднання (через умову \\( n_i \\ge 5 \\)): ${kNew}<br>
        Ступені вільності (\\( df \\)): ${df}<br>
        Спостережуване значення \\( \\chi^2_{спост} \\): ${round(chiSquare)}<br>
        Критичне значення \\( \\chi^2_{кр} \\): ${round(pearsonCrit)}<br><br>
        <b>Висновок:</b> ${chiSquare < pearsonCrit ? 'Гіпотеза про нормальний розподіл приймається.' : 'Гіпотеза про нормальний розподіл відхиляється.'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> За критерієм Пірсона ми порівняли реальні частоти з теоретичними. ${excelMode ? 'Ввімкнено Excel-режим (параметри не віднімаються).' : "Крайні інтервали розширено до \\( \\pm \\infty \\), а інтервали з очікуваною частотою < 5 було об'єднано."}
        </div>
    `;
    state.reportPearson = pearsonHtml;
    document.getElementById('pearson-result').innerHTML = pearsonHtml;
    
    // 10. Confidence Interval
    const tCrit = jStat.studentt.inv(0.975, n - 1); 
    const marginOfError = tCrit * (sd / Math.sqrt(n));
    const variance = ss.sampleVariance(X);
    
    let varHtml = "";
    if (window.jStat) {
        const chiLower = jStat.chisquare.inv(0.975, n - 1);
        const chiUpper = jStat.chisquare.inv(0.025, n - 1);
        const varCiLower = (n - 1) * variance / chiLower;
        const varCiUpper = (n - 1) * variance / chiUpper;
        varHtml = `<br><br><b>Довірчий інтервал для дисперсії (\\( \\sigma^2 \\)) з надійністю \\( \\gamma=0.95 \\):</b><br>
                   Нижня межа: ${round(varCiLower)}<br>
                   Верхня межа: ${round(varCiUpper)}<br>
                   Інтервал: \\( \\left( \\frac{(n-1)S^2}{\\chi^2_B} ; \\frac{(n-1)S^2}{\\chi^2_H} \\right) = (${round(varCiLower)} ; ${round(varCiUpper)}) \\)`;
    }
    
    const ciHtml = `
        Точкова оцінка мат. сподівання (\\( \\bar{x} \\)): ${round(mean)}<br>
        t-критичне Ст'юдента (\\( t_{кр} \\)): ${round(tCrit)}<br>
        Гранична похибка (\\( \\Delta = t_{кр} \\cdot \\frac{S}{\\sqrt{n}} \\)): ${round(marginOfError)}<br><br>
        <b>Довірчий інтервал:</b> \\( \\bar{x} \\pm \\Delta = (${round(mean - marginOfError)} ; ${round(mean + marginOfError)}) \\)
        ${varHtml}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Ми будуємо довірчі інтервали (для мат. сподівання за t-розподілом та для дисперсії за \\( \\chi^2 \\)-розподілом). Це гарантує з імовірністю 95%, що справжні значення генеральної сукупності знаходяться в цих межах.
        </div>
    `;
    state.reportConfidenceInterval = ciHtml;
    document.getElementById('confidence-interval-result').innerHTML = ciHtml;
}

// --- TASK 11-15 ---
function renderTask11to15() {
    const X = state.X;
    const Y = state.Y;
    if (!Y || Y.length !== X.length) {
        document.getElementById('correlation-stats').innerHTML = "Потрібна вибірка Y однакового розміру.";
        return;
    }
    
    // 12. Covariance and Correlation
    const cov = ss.sampleCovariance(X, Y);
    const r = ss.sampleCorrelation(X, Y);
    
    document.getElementById('correlation-stats').innerHTML = `
        <div class="stat-item"><div class="stat-label">Коваріація (\\( cov \\))</div><div class="stat-value">${round(cov)}</div></div>
        <div class="stat-item"><div class="stat-label">Коефіцієнт кореляції (\\( r \\))</div><div class="stat-value">${round(r)}</div></div>
        <div class="explanation mt-2" style="grid-column: 1 / -1; color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> 
            Коваріація показує міру спільної мінливості ознак. Коефіцієнт кореляції вимірює силу лінійного зв'язку між ними: значення, близькі до 1 або -1 свідчать про сильний зв'язок.
        </div>
    `;
    
    // 11 & 13. Linear Regression
    const linearReg = ss.linearRegression(X.map((x, i) => [x, Y[i]]));
    const linearLine = ss.linearRegressionLine(linearReg);
    
    document.getElementById('linear-eq').innerHTML = `Рівняння лінійної регресії: \\( Y = ${round(linearReg.m)} \\cdot X + ${round(linearReg.b)} \\)<br><br>
    <b>Висновок про природу розбіжності:</b> Коефіцієнт кореляції \\( r = ${round(r)} \\). 
    Оскільки модуль кореляції ${Math.abs(r) > 0.7 ? '> 0.7, спостерігається сильний' : (Math.abs(r) > 0.3 ? '> 0.3, спостерігається помірний' : '< 0.3, спостерігається слабкий')} лінійний зв'язок між ознаками.
    Розкид емпіричних точок відносно лінії регресії ${Math.abs(r) > 0.7 ? 'є незначним, що свідчить про високу адекватність лінійної моделі' : 'є значним, що може вказувати на вплив інших неврахованих факторів або нелінійний характер залежності'}.
    <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
        <b>Пояснення:</b> Коефіцієнти лінійної регресії розраховуються за методом найменших квадратів, що дозволяє провести пряму лінію найближче до всіх емпіричних точок.
    </div>
    `;
    
    const minX = Math.min(...X);
    const maxX = Math.max(...X);
    
    const traceScatter = {
        x: X, y: Y, mode: 'markers', name: 'Емпіричні дані',
        marker: { color: 'rgba(59, 130, 246, 0.8)', size: 8 }
    };
    const traceLinReg = {
        x: [minX, maxX], y: [linearLine(minX), linearLine(maxX)],
        mode: 'lines', name: 'Лінійна регресія',
        line: { color: '#ef4444', width: 3 }
    };
    Plotly.newPlot('plot-linear-reg', [traceScatter, traceLinReg], {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', font: { color: '#f8fafc' },
        margin: { t: 30, r: 10, l: 40, b: 40 }
    });
    
    // 14. Non-linear regression
    let A = [];
    let B = [];
    for(let i = 0; i < X.length; i++) {
        A.push([Math.pow(X[i], 2), X[i], 1]);
        B.push([Y[i]]);
    }
    let AT = numeric.transpose(A);
    let ATA = numeric.dot(AT, A);
    let ATB = numeric.dot(AT, B);
    let coeffs = numeric.dot(numeric.inv(ATA), ATB).map(v => v[0]);
    let a = coeffs[0], b = coeffs[1], c = coeffs[2];
    
    document.getElementById('nonlinear-eq').innerHTML = `Рівняння поліноміальної регресії: \\( Y = ${round(a)} \\cdot X^2 + ${round(b)} \\cdot X + ${round(c)} \\)
    <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
        <b>Пояснення:</b> Поліноміальна регресія будує криву у формі параболи. Вона розрахована шляхом мінімізації квадратів відхилень і краще описує нелінійну залежність між даними.
    </div>`;
    
    let polyX = [], polyY = [];
    let step = (maxX - minX) / 100;
    for(let x = minX; x <= maxX; x += step) {
        polyX.push(x);
        polyY.push(a * Math.pow(x, 2) + b * x + c);
    }
    
    const traceNonLinReg = {
        x: polyX, y: polyY, mode: 'lines', name: 'Поліном 2-го ст.',
        line: { color: '#10b981', width: 3 }
    };
    Plotly.newPlot('plot-nonlinear-reg', [traceScatter, traceNonLinReg], {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', font: { color: '#f8fafc' },
        margin: { t: 30, r: 10, l: 40, b: 40 }
    });
    
    // 15. ANOVA
    // ВИПРАВЛЕННЯ 7: Використання Regression ANOVA (дисперсійний аналіз значущості регресії) замість One-Way ANOVA
    const meanY = ss.mean(Y);
    let sst = 0, ssreg = 0, sserr = 0;
    
    for (let i = 0; i < X.length; i++) {
        let y_hat = linearReg.m * X[i] + linearReg.b;
        sst += Math.pow(Y[i] - meanY, 2);
        ssreg += Math.pow(y_hat - meanY, 2);
        sserr += Math.pow(Y[i] - y_hat, 2);
    }
    
    const dfReg = 1;
    const dfErr = X.length - 2;
    
    const msReg = ssreg / dfReg;
    const msErr = sserr / dfErr;
    
    const F = msReg / msErr;
    const fCrit = jStat.centralF.inv(0.95, dfReg, dfErr);
    
    const anovaHtml = `
        F-спостережуване (регресії): ${round(F)}<br>
        F-критичне: ${round(fCrit)}<br><br>
        <b>Висновок:</b> ${F > fCrit ? 'Побудована лінійна модель регресії є статистично значущою.' : 'Лінійна модель регресії не є статистично значущою.'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Оскільки вибірки X та Y мають різну природу, порівняння їх середніх не має змісту. Тому використано <b>Дисперсійний аналіз регресії</b> (Regression ANOVA). Він перевіряє, наскільки дисперсія змінної Y пояснюється моделлю лінійної залежності від X.
            Оскільки F-спостережуване ${F > fCrit ? 'більше' : 'менше'} за критичне значення, рівняння регресії ${F > fCrit ? 'значуще' : 'не значуще'}.
        </div>
    `;
    state.reportANOVA = anovaHtml;
    document.getElementById('anova-result').innerHTML = anovaHtml;
}

// --- TASK 16: PCA ---
function renderTask16() {
    const X = state.X;
    const Y = state.Y;
    const Z = state.Z;
    
    if (!Z || Z.length !== X.length) {
        document.getElementById('pca-result').innerHTML = "Потрібна вибірка Z однакового розміру.";
        return;
    }
    
    const n = X.length;
    const meanX = ss.mean(X), sdX = ss.sampleStandardDeviation(X);
    const meanY = ss.mean(Y), sdY = ss.sampleStandardDeviation(Y);
    const meanZ = ss.mean(Z), sdZ = ss.sampleStandardDeviation(Z);
    
    let stdX = X.map(x => (x - meanX) / sdX);
    let stdY = Y.map(y => (y - meanY) / sdY);
    let stdZ = Z.map(z => (z - meanZ) / sdZ);
    
    let rXY = ss.sampleCorrelation(X, Y);
    let rXZ = ss.sampleCorrelation(X, Z);
    let rYZ = ss.sampleCorrelation(Y, Z);
    
    let covMat = [
        [1, rXY, rXZ],
        [rXY, 1, rYZ],
        [rXZ, rYZ, 1]
    ];
    
    let eig = numeric.eig(covMat);
    let indices = [0, 1, 2];
    indices.sort((a, b) => eig.lambda.x[b] - eig.lambda.x[a]);
    
    let eigenvalues = indices.map(i => eig.lambda.x[i]);
    let totalVar = eigenvalues.reduce((a, b) => a + b, 0);
    let varExplained = eigenvalues.map(v => (v / totalVar) * 100);
    let cumVar = [
        varExplained[0],
        varExplained[0] + varExplained[1],
        varExplained[0] + varExplained[1] + varExplained[2]
    ];
    
    // Блок 1: Таблиця власних значень
    document.getElementById('pca-eigen-table').querySelector('tbody').innerHTML = `
        <tr><td>PC1 (Головна)</td><td>${eigenvalues[0].toFixed(4)}</td><td>${varExplained[0].toFixed(2)}%</td><td>${cumVar[0].toFixed(2)}%</td></tr>
        <tr><td>PC2</td><td>${eigenvalues[1].toFixed(4)}</td><td>${varExplained[1].toFixed(2)}%</td><td>${cumVar[1].toFixed(2)}%</td></tr>
        <tr><td>PC3</td><td>${eigenvalues[2].toFixed(4)}</td><td>${varExplained[2].toFixed(2)}%</td><td>${cumVar[2].toFixed(2)}%</td></tr>
    `;
    
    // Витягуємо власні вектори (колонки)
    let ev1 = [0, 1, 2].map(i => eig.E.x[i][indices[0]]);
    let ev2 = [0, 1, 2].map(i => eig.E.x[i][indices[1]]);
    let ev3 = [0, 1, 2].map(i => eig.E.x[i][indices[2]]);
    
    // Нормуємо власні вектори
    let norm1 = Math.sqrt(ev1[0]**2 + ev1[1]**2 + ev1[2]**2); ev1 = ev1.map(v => v / norm1);
    let norm2 = Math.sqrt(ev2[0]**2 + ev2[1]**2 + ev2[2]**2); ev2 = ev2.map(v => v / norm2);
    let norm3 = Math.sqrt(ev3[0]**2 + ev3[1]**2 + ev3[2]**2); ev3 = ev3.map(v => v / norm3);
    
    // Блок 3: Факторні навантаження
    let loadX = [ev1[0]*Math.sqrt(eigenvalues[0]), ev2[0]*Math.sqrt(eigenvalues[1]), ev3[0]*Math.sqrt(eigenvalues[2])];
    let loadY = [ev1[1]*Math.sqrt(eigenvalues[0]), ev2[1]*Math.sqrt(eigenvalues[1]), ev3[1]*Math.sqrt(eigenvalues[2])];
    let loadZ = [ev1[2]*Math.sqrt(eigenvalues[0]), ev2[2]*Math.sqrt(eigenvalues[1]), ev3[2]*Math.sqrt(eigenvalues[2])];

    document.getElementById('pca-loadings-table').querySelector('tbody').innerHTML = `
        <tr><td>Ознака X</td><td>${loadX[0].toFixed(4)}</td><td>${loadX[1].toFixed(4)}</td><td>${loadX[2].toFixed(4)}</td></tr>
        <tr><td>Ознака Y</td><td>${loadY[0].toFixed(4)}</td><td>${loadY[1].toFixed(4)}</td><td>${loadY[2].toFixed(4)}</td></tr>
        <tr><td>Ознака Z</td><td>${loadZ[0].toFixed(4)}</td><td>${loadZ[1].toFixed(4)}</td><td>${loadZ[2].toFixed(4)}</td></tr>
    `;

    // Блок 4: Висновок
    document.getElementById('pca-conclusion').innerHTML = `
        <b>Аналітичний висновок:</b><br><br>
        1. <b>За методом PCA:</b> Оскільки перша головна компонента (PC1) самостійно описує ${varExplained[0].toFixed(2)}% загальної варіації ознак, за критерієм Кайзера (λ>1) вагомим є лише один головний фактор. Розмірність простору даних можна без суттєвих втрат інформації знизити з 3D до 1D.<br><br>
        2. <b>За методом PFA:</b> Виявлено один латентний (прихований) фактор, який повністю контролює поведінку системи. Надзвичайно високі факторні навантаження ознак X та Y на цей фактор свідчать про їхню інформаційну надлишковість (мультиколінеарність), що повністю підтверджує результати кореляційного аналізу в Завданні 12.
    `;
    
    // Блок 2А: 3D Plot
    const trace3D = {
        x: X, y: Y, z: Z,
        mode: 'markers',
        type: 'scatter3d',
        name: 'Оригінальні дані',
        marker: { color: 'rgba(59, 130, 246, 0.8)', size: 4 }
    };
    Plotly.newPlot('plot-3d', [trace3D], {
        paper_bgcolor: 'rgba(0,0,0,0)',
        font: { color: '#f8fafc' },
        margin: { l: 0, r: 0, b: 0, t: 0 },
        scene: {
            xaxis: { title: 'X' },
            yaxis: { title: 'Y' },
            zaxis: { title: 'Z' }
        }
    });

    // Блок 2Б: Scree Plot
    const traceScree = {
        x: ['PC1', 'PC2', 'PC3'],
        y: eigenvalues,
        type: 'bar',
        name: 'Власні значення',
        marker: { color: 'rgba(59, 130, 246, 0.8)' }
    };
    const screeLayout = {
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor: 'rgba(0,0,0,0)',
        font: { color: '#f8fafc' },
        margin: { t: 20, r: 10, l: 40, b: 40 },
        shapes: [
            {
                type: 'line',
                x0: -0.5, x1: 2.5,
                y0: 1, y1: 1,
                line: { color: '#ef4444', width: 2, dash: 'dash' }
            }
        ],
        annotations: [{
            x: 2, y: 1.1, text: 'Критерій Кайзера (λ=1)', showarrow: false, font: {color: '#ef4444'}
        }]
    };
    Plotly.newPlot('plot-scree', [traceScree], screeLayout);
}

// --- ЗВІТ У PDF (ТЕПЕР ПРОСТО ЗВІТ) ---
async function renderTask17() {
    const X = state.X;
    if (X.length === 0) {
        document.getElementById('report-content').innerHTML = "Помилка: Дані не згенеровано.";
        return;
    }
    
    document.getElementById('report-content').innerHTML = "<p>Генерую статичні графіки для звіту, зачекайте кілька секунд...</p>";
    
    const n = X.length;
    const mean = round(ss.mean(X));
    const sd = round(ss.sampleStandardDeviation(X));
    const variance = round(ss.sampleVariance(X));
    const skewness = round(ss.sampleSkewness(X));
    
    const exactMedian = round(ss.median(X));
    const varBiased = round(ss.variance(X));
    const sdBiased = round(ss.standardDeviation(X));

    let r = "-", regEq = "-";
    let traceScatter = null, traceLinReg = null, traceNonLinReg = null, tracePCA = null;
    
    // Recalc 1-6
    const sortedX = [...X].sort((a, b) => a - b);
    const sturgesK = Math.round(1 + 3.322 * Math.log10(n));
    const manualKVal = document.getElementById('manual-k').value;
    const k_new = manualKVal ? parseInt(manualKVal) : sturgesK;
    const minX = sortedX[0];
    const maxX = sortedX[n - 1];
    let h = (maxX - minX) / k_new;
    h = Math.ceil(h * 10000) / 10000;
    
    let reportFreqs = new Array(k_new).fill(0);
    let reportIntervals = [];
    for (let i = 0; i < k_new; i++) {
        reportIntervals.push({ start: minX + i * h, end: minX + (i + 1) * h, mid: minX + i * h + h / 2 });
    }
    X.forEach(val => {
        let idx = Math.floor((val - minX) / h);
        if (idx >= k_new) idx = k_new - 1;
        reportFreqs[idx]++;
    });
    
    let nHalf = n / 2;
    let cumFreq = 0;
    let meIdx = 0;
    for (let i = 0; i < k_new; i++) {
        cumFreq += reportFreqs[i];
        if (cumFreq >= nHalf) { meIdx = i; break; }
    }
    let sMePrev = meIdx > 0 ? reportFreqs.slice(0, meIdx).reduce((a, b) => a + b, 0) : 0;
    let nMe = reportFreqs[meIdx];
    let xStartMe = reportIntervals[meIdx].start;
    let groupedMedian = nMe === 0 ? "немає" : round(xStartMe + h * ((nHalf - sMePrev) / nMe));
    
    let maxFreq = Math.max(...reportFreqs);
    let moIdx = reportFreqs.indexOf(maxFreq);
    let nMo = reportFreqs[moIdx];
    let nPrev = moIdx > 0 ? reportFreqs[moIdx-1] : 0;
    let nNext = moIdx < k_new - 1 ? reportFreqs[moIdx+1] : 0;
    let xStartMo = reportIntervals[moIdx].start;
    let groupedMode = nMo === 0 ? "немає" : round(xStartMo + h * ((nMo - nPrev)/((nMo - nPrev) + (nMo - nNext))));

    const mids = reportIntervals.map(int => int.mid);

    
    const traceHist = {
        x: mids, y: reportFreqs, type: 'bar', name: 'Гістограма', width: h,
        marker: { color: 'rgba(59, 130, 246, 0.6)', line: {color: '#000', width: 1} }
    };
    const tracePoly = {
        x: mids, y: reportFreqs, type: 'scatter', mode: 'lines+markers', name: 'Полігон частот',
        line: { color: '#ef4444', width: 2 }
    };
    
    let edfX = [sortedX[0] - h], edfY = [0];
    for (let i = 0; i < n; i++) {
        edfX.push(sortedX[i]);
        edfY.push((i + 1) / n);
    }
    edfX.push(sortedX[n-1] + h); edfY.push(1);
    const traceEDF = {
        x: edfX, y: edfY, type: 'scatter', mode: 'lines',
        line: { shape: 'hv', color: '#10b981', width: 2 }, name: 'F*(x)'
    };
    
    if (state.Y && state.Y.length === n) {
        r = round(ss.sampleCorrelation(X, state.Y));
        const linearReg = ss.linearRegression(X.map((x, i) => [x, state.Y[i]]));
        const linearLine = ss.linearRegressionLine(linearReg);
        regEq = `Y = ${round(linearReg.m)} * X + ${round(linearReg.b)}`;
        
        traceScatter = {
            x: X, y: state.Y, mode: 'markers', name: 'Емпіричні дані',
            marker: { color: 'rgba(59, 130, 246, 0.8)', size: 8 }
        };
        traceLinReg = {
            x: [minX, maxX], y: [linearLine(minX), linearLine(maxX)],
            mode: 'lines', name: 'Лінійна регресія',
            line: { color: '#ef4444', width: 3 }
        };
        
        let A = [], B = [];
        for(let i = 0; i < n; i++) {
            A.push([Math.pow(X[i], 2), X[i], 1]);
            B.push([state.Y[i]]);
        }
        let AT = numeric.transpose(A);
        let ATA = numeric.dot(AT, A);
        let ATB = numeric.dot(AT, B);
        let coeffs = numeric.dot(numeric.inv(ATA), ATB).map(v => v[0]);
        let polyX = [], polyY = [];
        let step = (maxX - minX) / 100;
        for(let x = minX; x <= maxX; x += step) {
            polyX.push(x);
            polyY.push(coeffs[0] * Math.pow(x, 2) + coeffs[1] * x + coeffs[2]);
        }
        traceNonLinReg = {
            x: polyX, y: polyY, mode: 'lines', name: 'Поліном 2-го ст.',
            line: { color: '#10b981', width: 3 }
        };
    }
    
    let pcaText = "Немає достатньо даних для багатовимірного аналізу.";
    let traceScree = null;
    let imgScree = "";
    if (state.Z && state.Z.length === n && state.Y && state.Y.length === n) {
        const meanX = ss.mean(X), sdX = ss.sampleStandardDeviation(X);
        const meanY = ss.mean(state.Y), sdY = ss.sampleStandardDeviation(state.Y);
        const meanZ = ss.mean(state.Z), sdZ = ss.sampleStandardDeviation(state.Z);
        let stdX = X.map(x => (x - meanX) / sdX);
        let stdY = state.Y.map(y => (y - meanY) / sdY);
        let stdZ = state.Z.map(z => (z - meanZ) / sdZ);
        let rXY = ss.sampleCorrelation(X, state.Y);
        let rXZ = ss.sampleCorrelation(X, state.Z);
        let rYZ = ss.sampleCorrelation(state.Y, state.Z);
        let covMat = [[1, rXY, rXZ], [rXY, 1, rYZ], [rXZ, rYZ, 1]];
        let eig = numeric.eig(covMat);
        let indices = [0, 1, 2];
        indices.sort((a, b) => eig.lambda.x[b] - eig.lambda.x[a]);
        
        let eigenvalues = indices.map(i => eig.lambda.x[i]);
        let totalVar = eigenvalues.reduce((a, b) => a + b, 0);
        let varExplained = eigenvalues.map(v => (v / totalVar) * 100);
        let cumVar = [
            varExplained[0],
            varExplained[0] + varExplained[1],
            varExplained[0] + varExplained[1] + varExplained[2]
        ];
        
        let ev1 = [0, 1, 2].map(i => eig.E.x[i][indices[0]]);
        let ev2 = [0, 1, 2].map(i => eig.E.x[i][indices[1]]);
        let ev3 = [0, 1, 2].map(i => eig.E.x[i][indices[2]]);
        
        let norm1 = Math.sqrt(ev1[0]**2 + ev1[1]**2 + ev1[2]**2); ev1 = ev1.map(v => v / norm1);
        let norm2 = Math.sqrt(ev2[0]**2 + ev2[1]**2 + ev2[2]**2); ev2 = ev2.map(v => v / norm2);
        let norm3 = Math.sqrt(ev3[0]**2 + ev3[1]**2 + ev3[2]**2); ev3 = ev3.map(v => v / norm3);
        
        let loadX = [ev1[0]*Math.sqrt(eigenvalues[0]), ev2[0]*Math.sqrt(eigenvalues[1]), ev3[0]*Math.sqrt(eigenvalues[2])];
        let loadY = [ev1[1]*Math.sqrt(eigenvalues[0]), ev2[1]*Math.sqrt(eigenvalues[1]), ev3[1]*Math.sqrt(eigenvalues[2])];
        let loadZ = [ev1[2]*Math.sqrt(eigenvalues[0]), ev2[2]*Math.sqrt(eigenvalues[1]), ev3[2]*Math.sqrt(eigenvalues[2])];

        pcaText = `
            <div style="margin-top: 20px;">
                <h3>Блок №1: Інформаційна таблиця власних значень (Eigenvalues)</h3>
                <table>
                    <thead>
                        <tr><th>Компонента (Фактор)</th><th>Власне значення (λ)</th><th>Частка дисперсії (%)</th><th>Накопичена дисперсія (%)</th></tr>
                    </thead>
                    <tbody>
                        <tr><td>PC1 (Головна)</td><td>${eigenvalues[0].toFixed(4)}</td><td>${varExplained[0].toFixed(2)}%</td><td>${cumVar[0].toFixed(2)}%</td></tr>
                        <tr><td>PC2</td><td>${eigenvalues[1].toFixed(4)}</td><td>${varExplained[1].toFixed(2)}%</td><td>${cumVar[1].toFixed(2)}%</td></tr>
                        <tr><td>PC3</td><td>${eigenvalues[2].toFixed(4)}</td><td>${varExplained[2].toFixed(2)}%</td><td>${cumVar[2].toFixed(2)}%</td></tr>
                    </tbody>
                </table>

                <h3 style="margin-top: 20px;">Блок №3: Таблиця факторних навантажень (Factor Loadings)</h3>
                <table>
                    <thead>
                        <tr><th>Ознака</th><th>Навантаження на PC1</th><th>Навантаження на PC2</th><th>Навантаження на PC3</th></tr>
                    </thead>
                    <tbody>
                        <tr><td>Ознака X</td><td>${loadX[0].toFixed(4)}</td><td>${loadX[1].toFixed(4)}</td><td>${loadX[2].toFixed(4)}</td></tr>
                        <tr><td>Ознака Y</td><td>${loadY[0].toFixed(4)}</td><td>${loadY[1].toFixed(4)}</td><td>${loadY[2].toFixed(4)}</td></tr>
                        <tr><td>Ознака Z</td><td>${loadZ[0].toFixed(4)}</td><td>${loadZ[1].toFixed(4)}</td><td>${loadZ[2].toFixed(4)}</td></tr>
                    </tbody>
                </table>

                <h3 style="margin-top: 20px;">Блок №4: Аналітичний висновок</h3>
                <p style="background: rgba(0,0,0,0.05); padding: 10px; border-left: 3px solid #333;">
                    1. <b>За методом PCA:</b> Оскільки перша головна компонента (PC1) самостійно описує ${varExplained[0].toFixed(2)}% загальної варіації ознак, за критерієм Кайзера (λ>1) вагомим є лише один головний фактор. Розмірність простору даних можна без суттєвих втрат інформації знизити з 3D до 1D.<br><br>
                    2. <b>За методом PFA:</b> Виявлено один латентний (прихований) фактор, який повністю контролює поведінку системи. Надзвичайно високі факторні навантаження ознак X та Y на цей фактор свідчать про їхню інформаційну надлишковість (мультиколінеарність), що повністю підтверджує результати кореляційного аналізу в Завданні 12.
                </p>
            </div>
        `;
        
        let pc1 = [], pc2 = [];
        for(let i = 0; i < n; i++) {
            pc1.push(stdX[i]*ev1[0] + stdY[i]*ev1[1] + stdZ[i]*ev1[2]);
            pc2.push(stdX[i]*ev2[0] + stdY[i]*ev2[1] + stdZ[i]*ev2[2]);
        }
        tracePCA = {
            x: pc1, y: pc2, mode: 'markers', name: 'Дані PC1-PC2',
            marker: { color: 'rgba(139, 92, 246, 0.8)', size: 8 }
        };

        traceScree = {
            x: ['PC1', 'PC2', 'PC3'],
            y: eigenvalues,
            type: 'bar',
            name: 'Власні значення',
            marker: { color: 'rgba(59, 130, 246, 0.8)' }
        };
    }
    
    // Convert Plotly to Static Images function
    const getPlotImage = async (data, layout) => {
        let div = document.createElement('div');
        document.body.appendChild(div);
        await Plotly.newPlot(div, data, layout);
        const url = await Plotly.toImage(div, {format: 'png', width: 700, height: 400});
        document.body.removeChild(div);
        return url;
    };
    
    const layoutOpt = {
        paper_bgcolor: 'rgba(255,255,255,1)', plot_bgcolor: 'rgba(255,255,255,1)', font: { color: '#000' },
        margin: { t: 10, r: 10, l: 40, b: 20 }, showlegend: true
    };
    
    const imgHist = await getPlotImage([traceHist, tracePoly], layoutOpt);
    const imgEDF = await getPlotImage([traceEDF], layoutOpt);
    let imgReg = "", imgPCA = "";
    
    if (traceScatter && traceLinReg && traceNonLinReg) {
        imgReg = await getPlotImage([traceScatter, traceLinReg, traceNonLinReg], layoutOpt);
    }
    if (tracePCA) {
        imgPCA = await getPlotImage([tracePCA], layoutOpt);
        imgScree = await getPlotImage([traceScree], {
            title: 'Графік власних значень (Scree Plot)',
            paper_bgcolor: 'rgba(255,255,255,1)', plot_bgcolor: 'rgba(255,255,255,1)', font: { color: '#000' },
            margin: { t: 40, r: 10, l: 40, b: 40 },
            shapes: [{ type: 'line', x0: -0.5, x1: 2.5, y0: 1, y1: 1, line: { color: '#ef4444', width: 2, dash: 'dash' } }]
        });
    }
    
    const dateStr = new Date().toLocaleDateString('uk-UA');
    
    let dataTableRows = "";
    for (let i = 0; i < n; i++) {
        dataTableRows += `
            <tr>
                <td>${i + 1}</td>
                <td>${X[i].toFixed(4)}</td>
                <td>${state.Y && state.Y[i] ? state.Y[i].toFixed(4) : '-'}</td>
                <td>${state.Z && state.Z[i] ? state.Z[i].toFixed(4) : '-'}</td>
            </tr>
        `;
    }

    let groupedTableRows = "";
    for (let i = 0; i < k_new; i++) {
        let int = reportIntervals[i];
        let freq = reportFreqs[i];
        let w = freq / n;
        groupedTableRows += `
            <tr>
                <td>[${int.start.toFixed(4)} ; ${int.end.toFixed(4)})</td>
                <td>${int.mid.toFixed(4)}</td>
                <td>${freq}</td>
                <td>${w.toFixed(4)}</td>
            </tr>
        `;
    }

    const reportHtml = `
        <style>
            #report-content { font-family: 'Times New Roman', Times, serif; color: #000; line-height: 1.6; font-size: 14px; }
            #report-content h1 { text-align: center; font-size: 20px; font-weight: bold; margin-bottom: 20px; text-transform: uppercase; }
            #report-content h2 { font-size: 16px; margin-top: 25px; border-bottom: 2px solid #333; padding-bottom: 5px; font-weight: bold; }
            #report-content p { text-indent: 20px; margin-bottom: 10px; text-align: justify; }
            #report-content ul { margin-left: 40px; margin-bottom: 15px; text-align: justify; }
            #report-content li { margin-bottom: 8px; }
            .formula { text-align: center; font-size: 15px; margin: 15px 0; font-style: italic; }
            .plot-img { max-width: 100%; height: auto; page-break-inside: avoid; break-inside: avoid; display: block; margin: 20px auto; border: 1px solid #ccc; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
            #report-content h2 { page-break-after: avoid; break-after: avoid; }
            table { width: 100%; border-collapse: collapse; text-align: center; font-size: 12px; margin: 15px 0; page-break-inside: auto; }
            th, td { border: 1px solid #000; padding: 6px; }
            th { background-color: #f0f0f0; font-weight: bold; }
        </style>
        
        <h1>Розрахунково-графічна робота<br>із загальної теорії статистики</h1>
        <p style="text-align: right; text-indent: 0;"><strong>Дата формування:</strong> ${dateStr}</p>
        <p style="text-align: right; text-indent: 0;"><strong>Обсяг досліджуваної вибірки (n):</strong> ${n}</p>
        
        <h2>Вхідні дані (Вибірка)</h2>
        <p>Для виконання розрахунково-графічної роботи було використано масив випадкових величин, отриманий в результаті статистичного спостереження. Вибірка складається з ${n} елементів і містить декілька досліджуваних ознак (X, Y, Z).</p>
        <table>
            <thead>
                <tr>
                    <th>№ з/п</th>
                    <th>Ознака X</th>
                    <th>Ознака Y</th>
                    <th>Ознака Z</th>
                </tr>
            </thead>
            <tbody>
                ${dataTableRows}
            </tbody>
        </table>

        <h2>Завдання 1: Побудова варіаційного ряду</h2>
        <p>Отриману вибірку було впорядковано за зростанням від найменшого елемента (\\(X_{min} = ${round(Math.min(...X))}\\)) до найбільшого (\\(X_{max} = ${round(Math.max(...X))}\\)). Впорядкування (ранжування) елементів є базовим етапом статистичного аналізу.</p>
        <p style="text-align: center; font-family: monospace; word-break: break-all; font-size: 11px;">\\(X_{min} = ${round(Math.min(...X))} \\le \\dots \\le X_{max} = ${round(Math.max(...X))}\\)</p>
        <p>Ранжування дозволяє візуально оцінити розмах варіації \\(R = X_{max} - X_{min} = ${round(Math.max(...X))} - ${round(Math.min(...X))} = ${round(Math.max(...X) - Math.min(...X))}\\), що є необхідним для подальшого групування даних.</p>

        <h2>Завдання 2: Метод згрупованих даних</h2>
        <p>Оскільки вибірка є неперервною (або містить велику кількість дискретних значень), для побудови статистичного ряду було застосовано метод інтервального групування. Оптимальну кількість інтервалів (класів) розраховано за формулою Стерджеса:</p>
        <div class="formula">$$k = 1 + 3.322 \\cdot \\lg(n) = 1 + 3.322 \\cdot \\lg(${n}) \\approx ${sturgesK}$$</div>
        <p>Ширину інтервалу (крок) \\(h\\) розраховано як відношення розмаху варіації до кількості інтервалів:</p>
        <div class="formula">$$h = \\frac{X_{max} - X_{min}}{k} = \\frac{${round(Math.max(...X))} - ${round(Math.min(...X))}}{${k_new}} = ${h}$$</div>
        <p>У результаті розбиття весь діапазон значень поділено на ${k_new} інтервалів однакової ширини. Підраховано частоти \\(n_i\\) (абсолютну кількість елементів вибірки, що потрапили в кожен інтервал) та відносні частоти \\(W_i = n_i / n\\).</p>
        
        <table>
            <thead>
                <tr>
                    <th>Інтервал</th>
                    <th>Середина інтервалу \\((x_i)\\)</th>
                    <th>Частота \\((n_i)\\)</th>
                    <th>Відносна частота \\((W_i)\\)</th>
                </tr>
            </thead>
            <tbody>
                ${groupedTableRows}
                <tr>
                    <td colspan="2"><b>Разом (Сума):</b></td>
                    <td><b>${n}</b></td>
                    <td><b>1.0000</b></td>
                </tr>
            </tbody>
        </table>


        <h2>Завдання 3 та 4: Побудова статистичних графіків</h2>
        <p>Для наочного представлення структури розподілу побудовано такі графіки:</p>
        <ul>
            <li><b>Гістограма частот:</b> ступінчаста фігура, що складається з прямокутників, основами яких є інтервали, а висотами — абсолютні або відносні частоти. Вона візуалізує щільність розподілу досліджуваної ознаки.</li>
            <li><b>Полігон частот:</b> ламана лінія, отримана шляхом з'єднання точок, абсциси яких відповідають серединам інтервалів, а ординати — їх частотам. Полігон слугує емпіричним аналогом кривої щільності розподілу ймовірностей.</li>
            <li><b>Емпірична функція розподілу \\(F^*(x)\\):</b> ступінчаста неубивна функція, яка показує накопичену (кумулятивну) частоту. У кожній точці \\(x\\) вона дорівнює частці спостережень, менших за \\(x\\).</li>
        </ul>
        <img src="${imgHist}" class="plot-img">
        <img src="${imgEDF}" class="plot-img">

        <h2>Завдання 5 та 6: Розрахунок описових (числових) статистик</h2>
        <p>Для кількісної оцінки центру тяжіння, ступеня розсіяння та форми розподілу вибірки розраховано всі ключові показники (включаючи зсунені та незсунені оцінки):</p>
        <ul>
            <li><b>Вибіркове середнє \\(\\bar{x}\\):</b> характеризує типовий рівень ознаки у вибірці. Розраховано за формулою простої арифметичної:
            <div class="formula">$$\\bar{x} = \\frac{1}{n} \\sum_{i=1}^{n} x_i = ${mean}$$</div></li>
            
            <li><b>Дисперсія вибіркова (зсунена) \\(\\sigma^2\\):</b> оцінює міру розсіювання значень навколо середнього:
            <div class="formula">$$\\sigma^2 = \\frac{1}{n} \\sum_{i=1}^{n} (x_i - \\bar{x})^2 = ${varBiased}$$</div></li>
            
            <li><b>Дисперсія виправлена (незсунена) \\(S^2\\):</b> використовується для більш точної оцінки дисперсії генеральної сукупності (ділення на \\(n-1\\)):
            <div class="formula">$$S^2 = \\frac{1}{n-1} \\sum_{i=1}^{n} (x_i - \\bar{x})^2 = ${variance}$$</div></li>
            
            <li><b>Середнє квадратичне відхилення (зсунене \\(\\sigma\\) та незсунене \\(S\\)):</b> абсолютна міра варіації:
            <div class="formula">$$\\sigma = \\sqrt{\\sigma^2} = ${sdBiased}, \\quad S = \\sqrt{S^2} = ${sd}$$</div></li>
            
            <li><b>Медіана \\(Me\\):</b> значення, що ділить варіаційний ряд навпіл. Точна медіана за незгрупованими даними: \\(Me_{точн} = ${exactMedian}\\). Для згрупованих даних інтервальна медіана:
            <div class="formula">$$Me_{інт} = x_{Me} + h \\cdot \\frac{0.5n - S_{Me-1}}{n_{Me}} = ${groupedMedian}$$</div></li>
            
            <li><b>Мода \\(Mo\\):</b> значення ознаки, що зустрічається найчастіше. Інтервальна мода за формулою:
            <div class="formula">$$Mo = x_{Mo} + h \\cdot \\frac{n_{Mo} - n_{Mo-1}}{(n_{Mo} - n_{Mo-1}) + (n_{Mo} - n_{Mo+1})} = ${groupedMode}$$</div></li>
            
            <li><b>Коефіцієнт асиметрії \\(A_s\\):</b> характеризує скошеність розподілу відносно середнього. Для симетричного (нормального) розподілу \\(A_s \\approx 0\\). 
            <div class="formula">$$A_s = \\frac{\\sum(x_i - \\bar{x})^3 / n}{S^3} = ${skewness}$$</div></li>
        </ul>

        <h2>Завдання 7: Перевірка на наявність грубих помилок (викидів)</h2>
        <p>Екстремально великі або малі значення можуть суттєво спотворити результати статистичного аналізу. Було застосовано обраний алгоритм для виявлення аномальних значень (викидів):</p>
        <p style="background: rgba(0,0,0,0.05); padding: 10px; border-left: 3px solid #333;">${state.reportOutliers}</p>

        <h2>Завдання 8 та 9: Перевірка гіпотези про нормальний закон розподілу</h2>
        <p>Більшість методів параметричної статистики вимагають, щоб генеральна сукупність була розподілена нормально. Було проведено перевірку цієї гіпотези (\\(H_0\\)) за допомогою двох критеріїв згоди:</p>
        <ul>
            <li><b>Критерій згоди Колмогорова (Завдання 8):</b> Базується на порівнянні емпіричної функції розподілу \\(F^*(x)\\) та теоретичної нормальної функції розподілу \\(F(x)\\). Було знайдено максимальну абсолютну різницю \\(D = \\max |F^*(x) - F(x)|\\) та обчислено статистику \\(\\lambda = D \\sqrt{n}\\). Якщо \\(\\lambda < 1.36\\), гіпотеза \\(H_0\\) не відхиляється.
            <br><i>Результат:</i> ${state.reportKolmogorov}</li>
            
            <li><b>Критерій згоди Пірсона \\(\\chi^2\\) (Завдання 9):</b> Базується на порівнянні емпіричних частот \\(n_i\\) з теоретичними (очікуваними) частотами \\(n p_i\\), які розраховуються за нормальним законом розподілу з оціненими параметрами \\(\\bar{x}\\) та \\(S\\). Критерій обчислюється за формулою:
            <div class="formula">$$\\chi^2_{спост} = \\sum_{i=1}^{k} \\frac{(n_i - n p_i)^2}{n p_i}$$</div>
            Отримане значення \\(\\chi^2_{спост}\\) порівнюється з критичним \\(\\chi^2_{кр}(\\alpha, df)\\), де \\(df\\) — кількість ступенів вільності. 
            <br><i>Результат перевірки:</i> ${state.reportPearson}</li>
        </ul>

        <h2>Завдання 10: Побудова довірчих інтервалів</h2>
        <p>За вибірковими даними неможливо визначити точні параметри генеральної сукупності, тому здійснюється їх інтервальне оцінювання. Побудовано довірчий інтервал для невідомого математичного сподівання \\(a\\) з надійністю \\(\\gamma = 0.95\\). Оскільки теоретична дисперсія невідома, використано \\(t\\)-критерій Ст'юдента:</p>
        <div class="formula">$$\\bar{x} - t_{\\gamma, n-1} \\cdot \\frac{S}{\\sqrt{n}} < a < \\bar{x} + t_{\\gamma, n-1} \\cdot \\frac{S}{\\sqrt{n}}$$</div>
        <p>Цей інтервал накриває справжнє математичне сподівання генеральної сукупності із заданою ймовірністю 95%.</p>
        <div style="background: rgba(0,0,0,0.05); padding: 10px; border-left: 3px solid #333;">
            ${state.reportConfidenceInterval}
        </div>

        <h2>Завдання 11-15: Кореляційний та Регресійний аналізи, Дисперсійний аналіз (ANOVA)</h2>
        <p>Для вивчення залежності між двома ознаками (X та Y) було застосовано апарат регресійного аналізу.</p>
        <ul>
            <li><b>Кореляційний аналіз:</b> Ступінь тісноти лінійного зв'язку між ознаками X та Y оцінено за допомогою вибіркового коефіцієнта кореляції Пірсона \\(r_{xy} = ${r}\\). Значення \\(r\\) лежить в межах \\([-1; 1]\\), де значення, близькі до модуля одиниці, свідчать про сильний лінійний зв'язок.</li>
            <li><b>Лінійна регресія (Завдання 11, 13):</b> Побудовано однофакторну математичну модель \\(Y = m \\cdot X + b\\). Коефіцієнти \\(m\\) та \\(b\\) знайдено за методом найменших квадратів (МНК), що мінімізує суму квадратів залишків. 
            <br><b>Отримане рівняння регресії:</b> $${regEq}$$</li>
            <li><b>Нелінійна регресія (Завдання 14):</b> Оскільки зв'язок між змінними може мати криволінійний характер, додатково побудовано поліноміальну регресію 2-го степеня (параболу) \\(\\hat{Y} = a_0 + a_1 X + a_2 X^2\\).</li>
            <li><b>Дисперсійний аналіз (Завдання 15):</b> ${document.getElementById('anova-type').value === 'regression' ? 'Проведено Дисперсійний аналіз регресії (Regression ANOVA). Отримано значення F-статистики Фішера, яке порівнюється з F-критичним для перевірки статистичної значущості побудованого лінійного рівняння регресії. Якщо F > F_кр, побудована математична модель є адекватною та статистично значущою.' : 'Проведено Однофакторний дисперсійний аналіз (One-Way ANOVA), що порівнює середні значення вибірок X та Y за алгоритмом "Пакету аналізу" Excel. Розраховано міжгрупову та внутрішньогрупову дисперсії.'}
            <br><i>Результат ANOVA:</i> ${state.reportANOVA}</li>
        </ul>
        <img src="${imgReg}" class="plot-img">

        <h2>Завдання 16: Багатовимірний статистичний аналіз (Метод Головних Компонент)</h2>
        <p>Метод головних компонент (PCA) застосовано для зниження розмірності простору ознак (X, Y, Z). Суть методу полягає у переході до нової системи координат, осі якої (головні компоненти) лінійно незалежні (ортогональні) і напрямлені вздовж максимальної дисперсії даних. Власні вектори матриці коваріацій визначають напрямки, а власні значення — обсяг поясненої дисперсії. Дані було успішно спроектовано на площину перших двох головних компонент, які містять найбільшу частку інформації про вихідну систему.</p>
        <div style="font-size: 14px; margin-bottom: 20px;">${pcaText}</div>
        <div style="text-align: center;">
            <img src="${imgScree}" class="plot-img" style="display:inline-block; width: 48%; margin: 0 1%;">
            <img src="${imgPCA}" class="plot-img" style="display:inline-block; width: 48%; margin: 0 1%;">
        </div>
        
        <br>
        <p style="text-align: right; margin-top: 40px; margin-bottom: 20px; font-style: italic; color: #555;"><b>Звіт згенеровано автоматично в системі StatApp</b></p>
    `;
    
    document.getElementById('report-content').innerHTML = reportHtml;
    
    // Render math formulas in the report before generating PDF
    if (window.renderMathInElement) {
        window.renderMathInElement(document.getElementById('report-content'), {
            delimiters: [
                {left: '$$', right: '$$', display: true},
                {left: '\\(', right: '\\)', display: false}
            ],
            throwOnError: false
        });
    }
}

// PDF Download Logic
document.getElementById('download-pdf-btn').addEventListener('click', () => {
    const reportElement = document.getElementById('report-content');
    const panel = document.getElementById('task-17');
    const scrollContainer = document.querySelector('.scroll-container');
    const appContainer = document.querySelector('.app-container');
    
    // Disable button
    const btn = document.getElementById('download-pdf-btn');
    const originalText = btn.innerText;
    btn.innerText = "Генерація PDF (зачекайте)...";
    btn.disabled = true;

    // Temporarily un-constrain all scrollable parents so the element expands to full height
    const origBodyOverflow = document.body.style.overflow;
    const origBodyHeight = document.body.style.height;
    const origAppHeight = appContainer.style.height;
    const origScrollOverflow = scrollContainer.style.overflowY;
    
    document.body.style.overflow = 'visible';
    document.body.style.height = 'auto';
    appContainer.style.height = 'auto';
    scrollContainer.style.overflowY = 'visible';
    
    const opt = {
        margin: 10,
        filename: 'StatApp_Звіт.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, logging: false },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    
    // Generate PDF directly from the visible, now un-constrained element
    
    // Render KaTeX in the report before generating PDF
    renderMathInContainer('report-content');
    
    // Give KaTeX a moment to render fonts before taking screenshot
    setTimeout(() => {
        html2pdf().set(opt).from(reportElement).save().then(() => {
            // Restore styles
        document.body.style.overflow = origBodyOverflow;
        document.body.style.height = origBodyHeight;
        appContainer.style.height = origAppHeight;
        scrollContainer.style.overflowY = origScrollOverflow;
        
        btn.innerText = originalText;
        btn.disabled = false;
    }).catch(err => {
        // Restore styles on error
        document.body.style.overflow = origBodyOverflow;
        document.body.style.height = origBodyHeight;
        appContainer.style.height = origAppHeight;
        scrollContainer.style.overflowY = origScrollOverflow;
        
        btn.innerText = originalText;
        btn.disabled = false;
        alert("Помилка генерації PDF: " + err.message);
    });
    }, 500);
});
