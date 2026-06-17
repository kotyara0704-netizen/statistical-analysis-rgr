// State
const state = {
    X: [],
    Y: [],
    Z: []
};

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
        }
    });
});

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
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
        
        state.X = []; state.Y = []; state.Z = [];
        for (let i = 1; i < jsonData.length; i++) {
            if(jsonData[i][0] !== undefined) state.X.push(Number(jsonData[i][0]));
            if(jsonData[i][1] !== undefined) state.Y.push(Number(jsonData[i][1]));
            if(jsonData[i][2] !== undefined) state.Z.push(Number(jsonData[i][2]));
        }
        
        if (state.X.length < 50) {
            alert("Увага: Обсяг вибірки менше 50!");
        }
        updateDataTable();
        alert("Дані успішно завантажено!");
    };
    reader.readAsArrayBuffer(file);
});

function updateDataTable() {
    const tbody = document.querySelector('#data-table tbody');
    tbody.innerHTML = '';
    const n = Math.min(state.X.length, 100);
    for (let i = 0; i < n; i++) {
        tbody.innerHTML += `
            <tr>
                <td>${i + 1}</td>
                <td>${state.X[i].toFixed(4)}</td>
                <td>${state.Y[i] ? state.Y[i].toFixed(4) : '-'}</td>
                <td>${state.Z[i] ? state.Z[i].toFixed(4) : '-'}</td>
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
        ${sortedX.map(v => round(v, 2)).join(', ')}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Елементи вибірки впорядковано за зростанням. Це необхідно для визначення мінімуму, максимуму та подальшого групування.
        </div>
    `;
    
    // 2. Grouped data
    const k = Math.floor(1 + 3.322 * Math.log10(n));
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
            <b>Пояснення:</b> Кількість інтервалів знайдена за формулою Стерджеса і становить ${k}.
            Ширину (крок) кожного інтервалу округлено в більший бік до ${round(h)}, щоб уникнути втрати максимального значення вибірки.
            Відносна частота обчислюється як відношення кількості елементів в інтервалі до загального обсягу вибірки.
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
    const median = ss.median(X);
    
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
        <div class="stat-item"><div class="stat-label">Середнє (Незсунена оцінка мат.сподівання)</div><div class="stat-value">${round(mean)}</div></div>
        <div class="stat-item"><div class="stat-label">Медіана (Me)</div><div class="stat-value">${round(median)}</div></div>
        <div class="stat-item"><div class="stat-label">Мода (Mo) для інтервального ряду</div><div class="stat-value">${modeVal}</div></div>
        <div class="stat-item"><div class="stat-label">Дисперсія (Зсунена, n)</div><div class="stat-value">${round(varBiased)}</div></div>
        <div class="stat-item"><div class="stat-label">Дисперсія (Незсунена, n-1)</div><div class="stat-value">${round(varUnbiased)}</div></div>
        <div class="stat-item"><div class="stat-label">СКВ (Зсунене)</div><div class="stat-value">${round(sdBiased)}</div></div>
        <div class="stat-item"><div class="stat-label">СКВ (Незсунене)</div><div class="stat-value">${round(sdUnbiased)}</div></div>
        <div class="stat-item"><div class="stat-label">Асиметрія</div><div class="stat-value">${round(skewness)}</div></div>
        <div class="explanation mt-2" style="grid-column: 1 / -1; color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b><br>
            Тут наведено основні описові статистики: середнє (центр розподілу), медіану (значення, що ділить вибірку навпіл) та моду (обчислену класичним методом для інтервального ряду). Дисперсія та середнє квадратичне відхилення (СКВ) характеризують розкид даних відносно середнього. Асиметрія показує наскільки графік розподілу "скошений" в один із боків.
        </div>
    `;
}

// --- TASK 7-10 ---
function renderTask7to10() {
    const X = state.X;
    const n = X.length;
    const sortedX = [...X].sort((a, b) => a - b);
    
    // 7. Outliers
    const q1 = ss.quantile(sortedX, 0.25);
    const q3 = ss.quantile(sortedX, 0.75);
    const iqr = q3 - q1;
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    
    const outliers = X.filter(x => x < lowerBound || x > upperBound);
    document.getElementById('outliers-result').innerHTML = `
        Q1 (25-й перцентиль): ${round(q1)}, Q3 (75-й перцентиль): ${round(q3)}, IQR: ${round(iqr)}<br>
        Нижня межа: ${round(lowerBound)}, Верхня межа: ${round(upperBound)}<br><br>
        <b>Виявлені грубі помилки (викиди):</b> ${outliers.length > 0 ? outliers.map(v => round(v)).join(', ') : 'Не виявлено'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> За допомогою міжквартильного розмаху ми знаходимо типові межі для нашої вибірки. Усі значення, які виходять за ці межі, вважаються аномальними (викидами) та не є типовими для даного набору даних.
        </div>
    `;
    
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
    document.getElementById('kolmogorov-result').innerHTML = `
        Спостережуване значення максимальної відстані D: ${round(maxD)}<br>
        Статистика Колмогорова (λ = D * √n): ${round(lambda)}<br>
        Критичне значення λ (α=0.05): ${round(kolmogorovCrit)}<br><br>
        <b>Висновок:</b> ${lambda < kolmogorovCrit ? 'Гіпотеза про нормальний розподіл приймається.' : 'Гіпотеза про нормальний розподіл відхиляється.'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> За критерієм Колмогорова ми порівняли нашу емпіричну функцію розподілу з теоретичною нормальною.
            Максимальне відхилення розраховувалось з урахуванням стрибків ступінчастої функції. Значення порівнюється зі статистикою Колмогорова $\\lambda$. Оскільки параметри (середнє та СКВ) оцінювались за вибіркою, більш точним був би критерій Лілієфорса, проте для студентських робіт класичний критерій Колмогорова (порівняння λ з 1.36) вважається достатнім.
        </div>
    `;
    
    // 9. Pearson
    // ВИПРАВЛЕННЯ 6: 
    // 1. Інтервали розширюються до -/+ нескінченності.
    // 2. Інтервали з очікуваною частотою < 5 об'єднуються.
    if (!globalGroupedData) return;
    let { k, intervals, freqs } = globalGroupedData;
    
    let expectedFreqs = [];
    for (let i = 0; i < k; i++) {
        let start = (i === 0) ? -Infinity : intervals[i].start;
        let end = (i === k - 1) ? Infinity : intervals[i].end;
        let prob = jStat.normal.cdf(end, mean, sd) - jStat.normal.cdf(start, mean, sd);
        expectedFreqs.push(n * prob);
    }
    
    let mergedObserved = [];
    let mergedExpected = [];
    let currentObs = 0;
    let currentExp = 0;
    
    for (let i = 0; i < k; i++) {
        currentObs += freqs[i];
        currentExp += expectedFreqs[i];
        
        if (currentExp >= 5 || i === k - 1) {
            mergedObserved.push(currentObs);
            mergedExpected.push(currentExp);
            currentObs = 0;
            currentExp = 0;
        }
    }
    
    // Якщо останній об'єднаний інтервал має < 5 очікуваних значень, приєднуємо його до попереднього
    if (mergedExpected.length > 1 && mergedExpected[mergedExpected.length - 1] < 5) {
        mergedExpected[mergedExpected.length - 2] += mergedExpected.pop();
        mergedObserved[mergedObserved.length - 2] += mergedObserved.pop();
    }
    
    let chiSquare = 0;
    let kNew = mergedExpected.length;
    
    if (kNew < 4) {
        document.getElementById('pearson-result').innerHTML = `
            Кількість інтервалів після об'єднання (через умову n_i ≥ 5): ${kNew}<br><br>
            <b style="color:#ef4444">Попередження:</b> Критерій Хі-квадрат неможливо застосувати коректно, оскільки k_new < 4. 
            Недостатньо ступенів вільності для перевірки нормального розподілу (df = ${kNew} - 3 ≤ 0). 
            Рекомендується збільшити обсяг вибірки n.
        `;
    } else {
        for (let i = 0; i < kNew; i++) {
            chiSquare += Math.pow(mergedObserved[i] - mergedExpected[i], 2) / mergedExpected[i];
        }
        
        const df = kNew - 3; 
        const pearsonCrit = jStat.chisquare.inv(0.95, df); 
        
        document.getElementById('pearson-result').innerHTML = `
            Кількість інтервалів після об'єднання (через умову n_i ≥ 5): ${kNew}<br>
            Ступені вільності (df): ${df}<br>
            Спостережуване значення хі-квадрат: ${round(chiSquare)}<br>
            Критичне значення хі-квадрат: ${round(pearsonCrit)}<br><br>
            <b>Висновок:</b> ${chiSquare < pearsonCrit ? 'Гіпотеза про нормальний розподіл приймається.' : 'Гіпотеза про нормальний розподіл відхиляється.'}
            <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
                <b>Пояснення:</b> За критерієм Пірсона ми порівняли реальні частоти з теоретичними. Згідно з методологією, крайні інтервали було розширено до нескінченності, а інтервали з очікуваною частотою менше 5 було об'єднано з сусідніми для коректної роботи критерію.
                Оскільки спостережуване значення ${chiSquare < pearsonCrit ? 'менше' : 'більше'} за критичне, ми робимо такий висновок.
            </div>
        `;
    }
    
    // 10. Confidence Interval
    const tCrit = jStat.studentt.inv(0.975, n - 1); 
    const marginOfError = tCrit * (sd / Math.sqrt(n));
    
    document.getElementById('confidence-interval-result').innerHTML = `
        Точкова оцінка: ${round(mean)}<br>
        t-критичне (Ст'юдента): ${round(tCrit)}<br>
        Гранична похибка: ${round(marginOfError)}<br><br>
        <b>Довірчий інтервал:</b> (${round(mean - marginOfError)} ; ${round(mean + marginOfError)})
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> За допомогою розподілу Ст'юдента ми побудували довірчий інтервал. Це діапазон, у якому ми гарантуємо з імовірністю 95%, що знаходиться справжнє математичне сподівання всієї генеральної сукупності.
        </div>
    `;
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
        <div class="stat-item"><div class="stat-label">Коваріація</div><div class="stat-value">${round(cov)}</div></div>
        <div class="stat-item"><div class="stat-label">Коефіцієнт кореляції</div><div class="stat-value">${round(r)}</div></div>
        <div class="explanation mt-2" style="grid-column: 1 / -1; color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> 
            Коваріація показує міру спільної мінливості ознак. Коефіцієнт кореляції вимірює силу лінійного зв'язку між ними: значення, близькі до 1 або -1 свідчать про сильний зв'язок.
        </div>
    `;
    
    // 11 & 13. Linear Regression
    const linearReg = ss.linearRegression(X.map((x, i) => [x, Y[i]]));
    const linearLine = ss.linearRegressionLine(linearReg);
    
    document.getElementById('linear-eq').innerHTML = `Рівняння: Y = ${round(linearReg.m)} * X + ${round(linearReg.b)}<br><br>
    <b>Висновок про природу розбіжності:</b> Коефіцієнт кореляції r = ${round(r)}. 
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
    
    document.getElementById('nonlinear-eq').innerHTML = `Рівняння: Y = ${round(a)} * X^2 + ${round(b)} * X + ${round(c)}
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
    
    document.getElementById('anova-result').innerHTML = `
        F-спостережуване (регресії): ${round(F)}<br>
        F-критичне: ${round(fCrit)}<br><br>
        <b>Висновок:</b> ${F > fCrit ? 'Побудована лінійна модель регресії є статистично значущою.' : 'Лінійна модель регресії не є статистично значущою.'}
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Оскільки вибірки X та Y мають різну природу, порівняння їх середніх не має змісту. Тому використано <b>Дисперсійний аналіз регресії</b> (Regression ANOVA). Він перевіряє, наскільки дисперсія змінної Y пояснюється моделлю лінійної залежності від X.
            Оскільки F-спостережуване ${F > fCrit ? 'більше' : 'менше'} за критичне значення, рівняння регресії ${F > fCrit ? 'значуще' : 'не значуще'}.
        </div>
    `;
}

// --- TASK 16: PCA / PFA ---
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
    
    document.getElementById('pca-result').innerHTML = `
        <b>Власні значення (Eigenvalues):</b> λ1=${round(eigenvalues[0])}, λ2=${round(eigenvalues[1])}, λ3=${round(eigenvalues[2])}<br>
        <b>Пояснена дисперсія:</b> PC1=${round(varExplained[0])}%, PC2=${round(varExplained[1])}%, PC3=${round(varExplained[2])}%<br><br>
        <b>Висновки (PCA/PFA):</b> Головна компонента 1 (PC1) пояснює ${round(varExplained[0])}% загальної дисперсії. 
        Якщо перші дві компоненти в сумі пояснюють понад 70-80% дисперсії, розмірність даних (X, Y, Z) можна сміливо зменшити до двовимірної без суттєвої втрати інформації.
        <div class="explanation mt-2" style="color:var(--text-secondary);font-size:0.9rem;">
            <b>Пояснення:</b> Метод головних компонент (PCA) шукає нові незалежні осі (компоненти), вздовж яких дані мають найбільший розкид (дисперсію). Чим більший відсоток дисперсії пояснює компонента, тим вона інформативніша. Це дозволяє зменшити розмірність даних від тривимірної до двовимірної без значної втрати інформації.
        </div>
    `;
    
    // ВИПРАВЛЕННЯ 4: Витягуємо власні вектори коректно. Вони йдуть по стовпцях, тому перебираємо оригінальні координати 0, 1, 2
    let ev1 = [0, 1, 2].map(i => eig.E.x[i][indices[0]]);
    let ev2 = [0, 1, 2].map(i => eig.E.x[i][indices[1]]);
    
    // Нормуємо власні вектори
    let norm1 = Math.sqrt(ev1[0]**2 + ev1[1]**2 + ev1[2]**2);
    ev1 = ev1.map(v => v / norm1);
    let norm2 = Math.sqrt(ev2[0]**2 + ev2[1]**2 + ev2[2]**2);
    ev2 = ev2.map(v => v / norm2);
    
    let pc1 = [], pc2 = [];
    for(let i = 0; i < n; i++) {
        pc1.push(stdX[i]*ev1[0] + stdY[i]*ev1[1] + stdZ[i]*ev1[2]);
        pc2.push(stdX[i]*ev2[0] + stdY[i]*ev2[1] + stdZ[i]*ev2[2]);
    }
    
    const tracePCA = {
        x: pc1, y: pc2, mode: 'markers', name: 'Дані в просторі PC1-PC2',
        marker: { color: 'rgba(139, 92, 246, 0.8)', size: 8 }
    };
    
    Plotly.newPlot('plot-pca', [tracePCA], {
        title: 'Метод Головних Компонент (PC1 vs PC2)',
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', font: { color: '#f8fafc' },
        xaxis: { title: 'PC1' }, yaxis: { title: 'PC2' },
        margin: { t: 40, r: 10, l: 40, b: 40 }
    });
}
