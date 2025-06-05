console.log("app.js: Script iniciado, esperando DOMContentLoaded...");

document.addEventListener('DOMContentLoaded', () => {
    console.log("app.js: DOMContentLoaded evento disparado. Inicializando aplicación...");

    // ---- INICIO DE MODIFICACIÓN ----
    const MAX_ROWS_TO_PROCESS = 5000; // Límite de filas a procesar. Puedes ajustar este número.
    // ---- FIN DE MODIFICACIÓN ----
    let currentFileName = '';
    let availableHeaders = [];
    let parsedCsvData = [];
    let lastChartConfig = null;
    let currentSmartSuggestion = null;
    let inferredColumnTypes = {};
    let currentSortColumnKey = null;
    let currentSortDirection = 'asc';
    let currentTableData = [];

    // --- Referencias a Elementos del DOM ---
    // ---- INICIO DE MODIFICACIÓN ----
    const dataProcessingInfoElement = document.getElementById('dataProcessingInfo'); // Nueva referencia
    // ---- FIN DE MODIFICACIÓN ----
    const dataUploadSection = document.getElementById('dataUploadSection');
    const selectFileButton = document.getElementById('selectFileButton');
    const fileInput = document.getElementById('fileInput');
    const fileLoadedStatus = document.getElementById('fileLoadedStatus');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const changeFileButton = document.getElementById('changeFileButton');
    const analysisDefinitionSection = document.getElementById('analysisDefinitionSection');
    const smartSuggestionBox = document.getElementById('smartSuggestionBox');
    const smartSuggestionText = document.getElementById('smartSuggestionText');
    const useSuggestionButton = document.getElementById('useSuggestionButton');
    const manualConfigButton = document.getElementById('manualConfigButton');
    const manualAnalysisForm = document.getElementById('manualAnalysisForm');
    const analysisTypeSelect = document.getElementById('analysisType');
    const columnSelectors = document.querySelectorAll('.column-selector');
    const trendFields = document.getElementById('trendFields');
    const comparisonFields = document.getElementById('comparisonFields');
    const distributionFields = document.getElementById('distributionFields');
    const generateAnalysisButton = document.getElementById('generateAnalysisButton');
    const dashboardResultsSection = document.getElementById('dashboardResultsSection');
    const dashboardFileName = document.getElementById('dashboardFileName');
    const chartContainer = document.getElementById('chartContainer'); // Importante
    const descriptiveSummaryElement = document.querySelector('#descriptiveSummary p'); // Ya lo tenías
    const dataTableContainer = document.getElementById('dataTableContainer');
    const downloadChartButton = document.getElementById('downloadChartButton');
    const toggleDataLabelsCheckbox = document.getElementById('toggleDataLabelsCheckbox');
    const addFilterButton = document.getElementById('addFilterButton');
    const activeFiltersContainer = document.getElementById('activeFiltersContainer');
    const downloadDataButton = document.getElementById('downloadDataButton');
    const performNewAnalysisButton = document.getElementById('performNewAnalysisButton');

    // Nuevas referencias para la funcionalidad de IA
    const getAISummaryButton = document.getElementById('getAISummaryButton');
    const aiSummaryResultContainer = document.getElementById('aiSummaryResultContainer');
    const aiSummaryTextElement = document.getElementById('aiSummaryText');

    // --- INICIO DE DEFINICIÓN DE FUNCIONES ---
    // (Las funciones que tenías marcadas con "/* ... (código completo de la función) ... */"
    // se asume que están completas y correctas según tu última versión)

    const filterConditions = [
        { value: 'equals', text: 'Es igual a (=)' }, { value: 'not_equals', text: 'No es igual a (≠)' },
        { value: 'contains', text: 'Contiene' }, { value: 'not_contains', text: 'No contiene' },
        { value: 'greater_than', text: 'Mayor que (>)' }, { value: 'less_than', text: 'Menor que (<)' },
        { value: 'greater_than_or_equal', text: 'Mayor o igual que (≥)' }, { value: 'less_than_or_equal', text: 'Menor o igual que (≤)' },
        { value: 'between', text: 'Está entre' },
        { value: 'is_empty', text: 'Está vacío' }, { value: 'is_not_empty', text: 'No está vacío' }
    ];

    function parseCSVText(csvText) { /* ... (tu código completo para parseCSVText) ... */
        if (!csvText) return { headers: [], data: [] };
        const lines = csvText.trim().split(/\r\n|\n/);
        if (lines.length === 0) return { headers: [], data: [] };
        const firstLine = lines[0];
        let delimiter = ';';
        const commaCount = (firstLine.match(/,/g) || []).length;
        const semicolonCount = (firstLine.match(/;/g) || []).length;
        if (commaCount > 0 && commaCount > semicolonCount) {
            delimiter = ',';
        }
        const rawHeaders = firstLine.split(delimiter).map(header => String(header || '').trim().replace(/^"|"$/g, ''));
        const data = [];
        if (lines.length > 1) {
            for (let i = 1; i < lines.length; i++) {
                if (lines[i].trim() === '') { continue; }
                const values = lines[i].split(delimiter);
                const rowObject = {};
                for (let j = 0; j < rawHeaders.length; j++) {
                    const headerKey = rawHeaders[j];
                    rowObject[headerKey] = values[j] ? String(values[j]).trim().replace(/^"|"$/g, '') : '';
                }
                data.push(rowObject);
            }
        }
        return { headers: rawHeaders, data: data };
    }

    function parseExcelData(arrayBuffer) { /* ... (tu código completo para parseExcelData) ... */
        try {
            if (typeof XLSX === 'undefined') {
                console.error("SheetJS (XLSX) no está definido. Asegúrate de que la librería esté cargada.");
                alert('Error: La librería para procesar archivos Excel no está cargada.');
                return { headers: [], data: [] };
            }
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const aoaData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'yyyy-mm-dd' });
            if (!aoaData || aoaData.length === 0) return { headers: [], data: [] };
            const rawHeaders = aoaData[0].map(header => String(header || '').trim());
            const data = [];
            if (aoaData.length > 1) {
                for (let i = 1; i < aoaData.length; i++) {
                    const rowArray = aoaData[i];
                    const rowObject = {};
                    for (let j = 0; j < rawHeaders.length; j++) {
                        let cellValue = rowArray[j];
                        if (cellValue instanceof Date) {
                            const year = cellValue.getFullYear();
                            const month = ('0' + (cellValue.getMonth() + 1)).slice(-2);
                            const day = ('0' + cellValue.getDate()).slice(-2);
                            rowObject[rawHeaders[j]] = `${year}-${month}-${day}`;
                        } else {
                            rowObject[rawHeaders[j]] = cellValue !== null && cellValue !== undefined ? String(cellValue).trim() : '';
                        }
                    }
                    data.push(rowObject);
                }
            }
            return { headers: rawHeaders, data: data };
        } catch (error) {
            console.error("Error en parseExcelData:", error);
            throw error;
        }
    }

    function inferColumnTypes(dataSample, headers) { /* ... (tu código completo para inferColumnTypes) ... */
        const types = {};
        const SAMPLES_TO_CHECK = Math.min(dataSample.length, 50);
        if (SAMPLES_TO_CHECK === 0) {
            if (headers && headers.length > 0) {
                headers.forEach(header => { types[header] = 'text'; });
            }
            return types;
        }
        if (headers && headers.length > 0) {
            headers.forEach(header => {
                let numCount = 0; let dateCount = 0; let emptyCount = 0;
                const uniqueValues = new Set();
                for (let i = 0; i < SAMPLES_TO_CHECK; i++) {
                    if(!dataSample[i]) { continue; }
                    const value = dataSample[i][header];
                    if (value === null || value === undefined || String(value).trim() === '') { emptyCount++; continue; }
                    uniqueValues.add(String(value));
                    if (!isNaN(parseFloat(value)) && isFinite(value)) numCount++;
                    if (value instanceof Date && !isNaN(value.getTime())) {
                        dateCount++;
                    } else if (typeof value === 'string' && (value.match(/^\d{4}-\d{2}-\d{2}$/) || value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/))) {
                         if (!isNaN(new Date(value).getTime())) dateCount++;
                    }
                }
                const nonEmptySamples = SAMPLES_TO_CHECK - emptyCount;
                if (nonEmptySamples === 0) { types[header] = 'text_empty'; }
                else if (dateCount / nonEmptySamples > 0.7) types[header] = 'date';
                else if (numCount / nonEmptySamples > 0.7) types[header] = 'numeric';
                else if (uniqueValues.size / nonEmptySamples < 0.5 && uniqueValues.size <= 15) types[header] = 'categorical';
                else types[header] = 'text';
            });
        }
        return types;
    }

    function populateColumnSelectors(headers) { /* ... (tu código completo para populateColumnSelectors) ... */
        if (!headers || headers.length === 0 || !columnSelectors) return;
        columnSelectors.forEach(select => {
            if (select) {
                const firstOption = select.options[0];
                select.innerHTML = '';
                if (firstOption) select.appendChild(firstOption);
                headers.forEach(header => {
                    if (header) {
                        const option = document.createElement('option');
                        option.value = header;
                        option.textContent = header;
                        select.appendChild(option);
                    }
                });
            }
        });
    }

    function resetColumnSelectors() { /* ... (tu código completo para resetColumnSelectors) ... */
        if (!columnSelectors) return;
        columnSelectors.forEach(select => {
            if (select && select.options[0]) {
                const firstOption = select.options[0];
                select.innerHTML = '';
                select.appendChild(firstOption);
                firstOption.selected = true;
            }
        });
    }

    function updateSmartSuggestion(headers, columnTypes) { /* ... (tu código completo para updateSmartSuggestion) ... */
        currentSmartSuggestion = null;
        if (!smartSuggestionText) return;
        if (!headers || headers.length === 0 || !columnTypes || Object.keys(columnTypes).length === 0) {
            smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> (No hay datos o encabezados para generar una sugerencia.)`;
            return;
        }
        let potentialDateCols = headers.filter(h => columnTypes[h] === 'date');
        let potentialNumericCols = headers.filter(h => columnTypes[h] === 'numeric');
        let potentialCategoricalCols = headers.filter(h => columnTypes[h] === 'categorical');
        let suggestionMade = false;
        if (potentialDateCols.length > 0 && potentialNumericCols.length > 0) {
            const dateCol = potentialDateCols[0];
            const numCol = potentialNumericCols.find(nc => nc !== dateCol) || potentialNumericCols[0];
            currentSmartSuggestion = { type: 'trend', xCol: dateCol, yCol: numCol };
            smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> Podrías analizar la <strong>'Tendencia de '${numCol}'</strong> usando <strong>'${dateCol}'</strong>. <br><small><em>(Considera explorar agrupaciones por mes o año para más detalle)</em></small>. ¿Usar esta sugerencia?`;
            suggestionMade = true;
        }
        else if (potentialCategoricalCols.length > 0 && potentialNumericCols.length > 0) {
            const catCol = potentialCategoricalCols[0];
            const numCol = potentialNumericCols.find(nc => nc !== catCol) || potentialNumericCols[0];
            currentSmartSuggestion = { type: 'comparison', catCol: catCol, valCol: numCol };
            smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> Intenta una <strong>Comparación de los totales/promedios de '${numCol}'</strong> para cada <strong>'${catCol}'</strong>. ¿Usar esta sugerencia?`;
            suggestionMade = true;
        }
        else if (potentialNumericCols.length > 0) {
            const numCol = potentialNumericCols[0];
            currentSmartSuggestion = { type: 'distribution', dataCol: numCol };
            smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> Explora la <strong>Distribución de '${numCol}'</strong> para ver cómo se agrupan sus valores. ¿Usar esta sugerencia?`;
            suggestionMade = true;
        }
        if (!suggestionMade) {
            if (headers.length >= 2) {
                currentSmartSuggestion = { type: 'trend', xCol: headers[0], yCol: headers[1] };
                smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> Podrías analizar la <strong>'Tendencia de ${headers[1]}'</strong> usando <strong>'${headers[0]}'</strong>. ¿Usar esta sugerencia?`;
            } else if (headers.length === 1) {
                 currentSmartSuggestion = { type: 'distribution', dataCol: headers[0] };
                 smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> Podrías ver la <strong>'Distribución de ${headers[0]}'</strong>. ¿Usar esta sugerencia?`;
            } else {
                smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> (No se pudo generar una sugerencia clara.)`;
            }
        }
    }

    function generateDescriptiveSummary(type, xVals, yVals, xColName, yColName, catColName, valColName, distColName, dataSetForSummary) { /* ... (tu código completo para generateDescriptiveSummary - considera aplicar las mejoras de formato aquí) ... */
        let summary = "No se pudo generar un resumen detallado.";
        if (type === 'trend') {
            if (!yVals || yVals.length === 0) return "No hay datos Y para el resumen de tendencia.";
            let numericY = yVals.map(v => parseFloat(v)).filter(v => !isNaN(v));
            if (numericY.length === 0) return `La columna '${yColName}' no contiene datos numéricos para el resumen.`;
            let maxY = -Infinity, minY = Infinity;
            let maxXassociated = null, minXassociated = null;
            for (let i = 0; i < yVals.length; i++) {
                const currentY = parseFloat(yVals[i]);
                if (!isNaN(currentY)) {
                    if (currentY > maxY) { maxY = currentY; maxXassociated = xVals[i]; }
                    if (currentY < minY) { minY = currentY; minXassociated = xVals[i]; }
                }
            }
            summary = `Análisis de tendencia para <strong>'${yColName}'</strong> por <strong>'${xColName}'</strong>:<br>`;
            summary += `&bull; Valor más alto observado: <strong>${maxY.toLocaleString()}</strong>`;
            if (maxXassociated !== null) summary += ` (en '${xColName}' = ${maxXassociated})`;
            summary += `<br>&bull; Valor más bajo observado: <strong>${minY.toLocaleString()}</strong>`;
            if (minXassociated !== null) summary += ` (en '${xColName}' = ${minXassociated})`;
            summary += ".";
            if (numericY.length > 1) {
                const firstVal = numericY[0]; const lastVal = numericY[numericY.length - 1];
                summary += "<br><em>";
                if (lastVal > firstVal) summary += "En general, se observa una tendencia al alza.";
                else if (lastVal < firstVal) summary += "En general, se observa una tendencia a la baja.";
                else summary += "En general, no se observa una clara tendencia ascendente o descendente en los extremos.";
                summary += "</em>";
            }
            return summary;
        } else if (type === 'comparison') {
            if (!yVals || yVals.length === 0) return "No hay datos de valor para el resumen de comparación.";
            let numericY = yVals.map(v => parseFloat(v)).filter(v => !isNaN(v));
            if (numericY.length === 0) return `La columna '${valColName}' no contiene datos numéricos para el resumen.`;
            let maxY = -Infinity, minY = Infinity;
            let maxCatAssociated = null, minCatAssociated = null;
            for (let i = 0; i < numericY.length; i++) {
                const currentY = numericY[i];
                if (currentY > maxY) { maxY = currentY; maxCatAssociated = xVals[i]; }
                if (currentY < minY) { minY = currentY; minCatAssociated = xVals[i]; }
            }
            summary = `Comparación de <strong>'${valColName}'</strong> por <strong>'${catColName}'</strong>:<br>`;
            summary += `&bull; Categoría con valor más alto (<strong>${maxCatAssociated}</strong>): <strong>${maxY.toLocaleString()}</strong><br>`;
            summary += `&bull; Categoría con valor más bajo (<strong>${minCatAssociated}</strong>): <strong>${minY.toLocaleString()}</strong>.`;
            return summary;
        } else if (type === 'distribution') {
            if (!xVals || xVals.length === 0) return "No hay datos para el resumen de distribución.";
            let numericX = xVals.map(v => parseFloat(v)).filter(v => !isNaN(v));
            if (numericX.length === 0) return `La columna '${distColName}' no contiene datos numéricos para el resumen.`;
            let sum = numericX.reduce((a, b) => a + b, 0);
            let mean = sum / numericX.length;
            let minVal = Math.min(...numericX);
            let maxVal = Math.max(...numericX);
            summary = `Distribución de <strong>'${distColName}'</strong>:<br>`;
            summary += `&bull; Los valores varían entre <strong>${minVal.toLocaleString()}</strong> y <strong>${maxVal.toLocaleString()}</strong>.<br>`;
            summary += `&bull; El promedio es <strong>${mean.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</strong>.`;
            return summary;
        }
        return "Resumen descriptivo no disponible para este tipo de análisis.";
    }

    function resetFileSelection() { /* ... (tu código completo para resetFileSelection) ... */
        if(fileLoadedStatus) fileLoadedStatus.style.display = 'none';
        currentFileName = '';
        if(fileNameDisplay) fileNameDisplay.textContent = '';
        if(fileInput) fileInput.value = null;
        updateSmartSuggestion([], {});
        // ---- INICIO DE MODIFICACIÓN ----
        if (dataProcessingInfoElement) dataProcessingInfoElement.textContent = '';
        // ---- FIN DE MODIFICACIÓN ----
    }

    function displayDynamicFields() { /* ... (tu código completo para displayDynamicFields) ... */
        if(trendFields) trendFields.style.display = 'none';
        if(comparisonFields) comparisonFields.style.display = 'none';
        if(distributionFields) distributionFields.style.display = 'none';

        if (!analysisTypeSelect) return;
        const selectedType = analysisTypeSelect.value;

        const checkboxLabel = toggleDataLabelsCheckbox ? toggleDataLabelsCheckbox.parentElement : null;
        if (checkboxLabel) {
            if (selectedType === 'distribution') {
                if(toggleDataLabelsCheckbox) toggleDataLabelsCheckbox.checked = false;
                checkboxLabel.style.display = 'none';
            } else {
                checkboxLabel.style.display = 'inline-flex';
            }
        }

        if (selectedType === 'trend' && trendFields) trendFields.style.display = 'block';
        else if (selectedType === 'comparison' && comparisonFields) comparisonFields.style.display = 'block';
        else if (selectedType === 'distribution' && distributionFields) distributionFields.style.display = 'block';
    }

    function updateFilterValueElement(filterRow, selectedColumnName, selectedCondition) { /* ... (tu código completo para updateFilterValueElement) ... */
        const valueContainer = filterRow.querySelector('.filter-value-container');
        if (!valueContainer) return;
        valueContainer.innerHTML = '';
        const valueNeedsSingleInput = !(selectedCondition === 'is_empty' || selectedCondition === 'is_not_empty' || selectedCondition === 'between');
        const valueNeedsTwoInputs = selectedCondition === 'between';
        const columnType = selectedColumnName ? inferredColumnTypes[selectedColumnName] : 'text';
        let inputType = 'text';

        if (valueNeedsTwoInputs || valueNeedsSingleInput) {
            if (columnType === 'numeric') inputType = 'number';
            else if (columnType === 'date') inputType = 'date';
        }

        if (valueNeedsTwoInputs && selectedColumnName) {
            const valueInputMin = document.createElement('input');
            valueInputMin.type = inputType;
            valueInputMin.className = 'filter-value-min';
            valueInputMin.placeholder = 'Valor mínimo';
            valueInputMin.setAttribute('aria-label', 'Valor mínimo del rango para el filtro');
            if (inputType === 'number') valueInputMin.step = 'any';
            valueContainer.appendChild(valueInputMin);

            const valueInputMax = document.createElement('input');
            valueInputMax.type = inputType;
            valueInputMax.className = 'filter-value-max';
            valueInputMax.placeholder = 'Valor máximo';
            valueInputMax.setAttribute('aria-label', 'Valor máximo del rango para el filtro');
            if (inputType === 'number') valueInputMax.step = 'any';
            valueContainer.appendChild(valueInputMax);

        } else if ((selectedCondition === 'equals' || selectedCondition === 'not_equals') && selectedColumnName && valueNeedsSingleInput) {
            const valueSelect = document.createElement('select');
            valueSelect.className = 'filter-value-select';
            valueSelect.setAttribute('aria-label', 'Seleccionar valor para el filtro');
            const defaultOpt = document.createElement('option');
            defaultOpt.value = ""; defaultOpt.textContent = "-- Selecciona valor --"; valueSelect.appendChild(defaultOpt);
            if (parsedCsvData.length > 0) {
                const uniqueValues = [...new Set(parsedCsvData.map(row => row[selectedColumnName]))]
                                         .filter(val => val !== null && val !== undefined && String(val).trim() !== '')
                                         .sort((a, b) => {
                                             if (columnType === 'numeric') return parseFloat(a) - parseFloat(b);
                                             return String(a).localeCompare(String(b), undefined, {numeric: true});
                                         });
                uniqueValues.forEach(val => {
                    const option = document.createElement('option'); option.value = val; option.textContent = val; valueSelect.appendChild(option);
                });
            }
            valueContainer.appendChild(valueSelect);
        } else {
            const valueInput = document.createElement('input');
            valueInput.type = inputType;
            valueInput.className = 'filter-value';
            valueInput.placeholder = 'Valor a filtrar';
            valueInput.setAttribute('aria-label', 'Valor para filtrar');
            valueInput.disabled = !valueNeedsSingleInput;
            if (inputType === 'number') valueInput.step = 'any';
            if (!valueNeedsSingleInput && !valueNeedsTwoInputs) valueInput.value = '';
            valueContainer.appendChild(valueInput);
        }
    }

    function collectFilterDefinitions() { /* ... (tu código completo para collectFilterDefinitions) ... */
        const definitions = [];
        if(!activeFiltersContainer) return definitions;
        const filterRows = activeFiltersContainer.querySelectorAll('.filter-row');
        filterRows.forEach(row => {
            const column = row.querySelector('.filter-column').value;
            const condition = row.querySelector('.filter-condition').value;
            const valueContainer = row.querySelector('.filter-value-container');
            let value;
            if (!column) return;
            if (condition === 'between') {
                const minInput = valueContainer.querySelector('.filter-value-min');
                const maxInput = valueContainer.querySelector('.filter-value-max');
                value = {
                    min: minInput ? minInput.value : '',
                    max: maxInput ? maxInput.value : ''
                };
            } else if (condition === 'is_empty' || condition === 'is_not_empty') {
                value = '';
            } else {
                const valueSelect = valueContainer.querySelector('.filter-value-select');
                const valueTextInput = valueContainer.querySelector('.filter-value');
                if (valueSelect) {
                    value = valueSelect.value;
                } else if (valueTextInput) {
                    value = valueTextInput.value;
                } else {
                    value = '';
                }
            }
            definitions.push({ column, condition, value });
        });
        return definitions;
    }

    function applyAllFilters(dataToFilter, filterDefinitions) { /* ... (tu código completo para applyAllFilters) ... */
        if (!filterDefinitions || filterDefinitions.length === 0) {
            return dataToFilter;
        }
        return dataToFilter.filter(row => {
            return filterDefinitions.every(filter => {
                const rowValueOriginal = row[filter.column];
                if (filter.condition === 'is_empty') {
                    return rowValueOriginal === null || rowValueOriginal === undefined || String(rowValueOriginal).trim() === '';
                }
                if (filter.condition === 'is_not_empty') {
                    return !(rowValueOriginal === null || rowValueOriginal === undefined || String(rowValueOriginal).trim() === '');
                }
                const rowValueString = String(rowValueOriginal || '').toLowerCase();
                const rowValueNumeric = parseFloat(rowValueOriginal);
                if (filter.condition === 'between') {
                    if (typeof filter.value !== 'object' || filter.value.min === undefined || filter.value.max === undefined) return true;
                    const filterMinString = String(filter.value.min || '').toLowerCase();
                    const filterMaxString = String(filter.value.max || '').toLowerCase();
                    const filterMinNumeric = parseFloat(filter.value.min);
                    const filterMaxNumeric = parseFloat(filter.value.max);
                    if (filter.value.min === '' && filter.value.max === '') return true;
                    let minCheck = filter.value.min === '' ? true :
                                   (!isNaN(rowValueNumeric) && !isNaN(filterMinNumeric) ? rowValueNumeric >= filterMinNumeric : rowValueString >= filterMinString);
                    let maxCheck = filter.value.max === '' ? true :
                                   (!isNaN(rowValueNumeric) && !isNaN(filterMaxNumeric) ? rowValueNumeric <= filterMaxNumeric : rowValueString <= filterMaxString);
                    return minCheck && maxCheck;
                }
                const filterValueString = String(filter.value || '').toLowerCase();
                const filterValueNumeric = parseFloat(filter.value);
                switch (filter.condition) {
                    case 'equals':
                        if (filter.value.trim() === '' && (rowValueOriginal === null || rowValueOriginal === undefined || String(rowValueOriginal).trim() === '')) return true;
                        if (!isNaN(rowValueNumeric) && !isNaN(filterValueNumeric) && String(filter.value).trim() !== '') {
                           return rowValueNumeric === filterValueNumeric;
                        }
                        return rowValueString === filterValueString;
                    case 'not_equals':
                        if (filter.value.trim() === '' && (rowValueOriginal === null || rowValueOriginal === undefined || String(rowValueOriginal).trim() === '')) return false;
                        if (!isNaN(rowValueNumeric) && !isNaN(filterValueNumeric) && String(filter.value).trim() !== '') {
                            return rowValueNumeric !== filterValueNumeric;
                        }
                        return rowValueString !== filterValueString;
                    case 'contains': return rowValueString.includes(filterValueString);
                    case 'not_contains': return !rowValueString.includes(filterValueString);
                    case 'greater_than': return !isNaN(rowValueNumeric) && !isNaN(filterValueNumeric) && rowValueNumeric > filterValueNumeric;
                    case 'less_than': return !isNaN(rowValueNumeric) && !isNaN(filterValueNumeric) && rowValueNumeric < filterValueNumeric;
                    case 'greater_than_or_equal': return !isNaN(rowValueNumeric) && !isNaN(filterValueNumeric) && rowValueNumeric >= filterValueNumeric;
                    case 'less_than_or_equal': return !isNaN(rowValueNumeric) && !isNaN(filterValueNumeric) && rowValueNumeric <= filterValueNumeric;
                    default: return true;
                }
            });
        });
    }

    function displayDataTable(headers, data) { /* ... (tu código completo para displayDataTable) ... */
        if (!dataTableContainer) return;
        dataTableContainer.innerHTML = '';
        currentTableData = [...data];
        if (!currentTableData || currentTableData.length === 0) {
            dataTableContainer.innerHTML = '<p style="text-align:center; padding: 20px; color: #777;">No hay datos para mostrar en la tabla.</p>';
            return;
        }
        if (!headers || headers.length === 0) {
            dataTableContainer.innerHTML = '<p style="text-align:center; padding: 20px; color: #777;">Faltan encabezados para mostrar la tabla.</p>';
            return;
        }
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            th.style.cursor = 'pointer';
            th.dataset.columnKey = headerText;
            if (headerText === currentSortColumnKey) {
                th.textContent += currentSortDirection === 'asc' ? ' ▲' : ' ▼';
            }
            th.addEventListener('click', () => {
                sortTableByColumn(headerText, headers, table.querySelector('tbody'));
            });
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);
        const tbody = document.createElement('tbody');
        populateTableBody(tbody, currentTableData, headers);
        table.appendChild(tbody);
        dataTableContainer.appendChild(table);
    }

    function populateTableBody(tbodyElement, dataToDisplay, headers) { /* ... (tu código completo para populateTableBody) ... */
        if (!tbodyElement) return;
        tbodyElement.innerHTML = '';
        dataToDisplay.forEach(rowData => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = rowData[header] !== null && rowData[header] !== undefined ? rowData[header] : '';
                tr.appendChild(td);
            });
            tbodyElement.appendChild(tr);
        });
    }

    function sortTableByColumn(headerKey, allHeaders, tbodyElement) { /* ... (tu código completo para sortTableByColumn) ... */
        if (!tbodyElement) {
            console.error("Error: tbodyElement no existe en sortTableByColumn.");
            return;
        }
        if (currentSortColumnKey === headerKey) {
            currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            currentSortColumnKey = headerKey;
            currentSortDirection = 'asc';
        }
        currentTableData.sort((a, b) => {
            let valA = a[currentSortColumnKey];
            let valB = b[currentSortColumnKey];
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);
            if (!isNaN(numA) && !isNaN(numB)) {
                valA = numA;
                valB = numB;
            } else {
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
            }
            if (valA < valB) { return currentSortDirection === 'asc' ? -1 : 1; }
            if (valA > valB) { return currentSortDirection === 'asc' ? 1 : -1; }
            return 0;
        });
        const tableElement = tbodyElement.closest('table');
        if (tableElement && tableElement.querySelector('thead tr')) {
            const headerCells = tableElement.querySelector('thead tr').cells;
            for (let th of headerCells) {
                if (th.dataset.columnKey) {
                    th.textContent = th.dataset.columnKey;
                    if (th.dataset.columnKey === currentSortColumnKey) {
                        th.textContent += currentSortDirection === 'asc' ? ' ▲' : ' ▼';
                    }
                }
            }
        }
        populateTableBody(tbodyElement, currentTableData, allHeaders);
    }

    function renderOrUpdateChart(config, dataForChart) { /* ... (tu código completo para renderOrUpdateChart, pero considera añadir las líneas para ocultar el resumen de IA como te mostré) ... */
        const {
            selectedAnalysisType,
            xAxisColumnName, yAxisColumnName,
            categoryColumnName, valueColumnName,
            dataColumnName
        } = config;
        const showDataLabels = toggleDataLabelsCheckbox ? toggleDataLabelsCheckbox.checked : false;

        if (chartContainer) chartContainer.innerHTML = '';
        if (descriptiveSummaryElement) descriptiveSummaryElement.innerHTML = `<span class="icon">📝</span> <strong>Resumen:</strong> Generando...`;
        
        // ---- INICIO DE MODIFICACIÓN ----
        // Ocultar el resumen de IA al generar un nuevo gráfico
        if (aiSummaryResultContainer) aiSummaryResultContainer.style.display = 'none';
        if (aiSummaryTextElement) aiSummaryTextElement.innerHTML = ''; // Limpiar texto anterior
        // ---- FIN DE MODIFICACIÓN ----

        if (!dataForChart || dataForChart.length === 0) {
            if(chartContainer) chartContainer.innerHTML = 'No hay datos para mostrar con los filtros y configuración actual.';
            if(descriptiveSummaryElement) descriptiveSummaryElement.innerHTML = `<span class="icon">📝</span> <strong>Resumen:</strong> No hay datos disponibles para el análisis.`;
            if(dataTableContainer) dataTableContainer.innerHTML = '<p style="text-align:center; padding: 20px; color: #777;">No hay datos para mostrar en la tabla.</p>';
            return;
        }

        let plotData = [];
        let layout = {
            title: {
                text: `Análisis de ${currentFileName}`,
                x: 0.5,
                xanchor: 'center',
                y: 0.97,
                yanchor: 'top',
                font: { size: 16 }
            },
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            autosize: false,
            width: chartContainer ? chartContainer.offsetWidth - 20 : 780,
            height: 450,
            margin: {
                l: 70,
                r: 40,
                b: 110,
                t: 75,
                pad: 5
            },
            xaxis: {
                automargin: true,
                tickangle: -45,
                tickfont: { size: 10 }
            },
            yaxis: {
                automargin: true,
                tickfont: { size: 10 }
            }
        };

        let traceMode = 'lines+markers';
        let textValues = null;
        let textPosition = 'top center';
        const textFont = { family: 'Arial, sans-serif', size: 10, color: 'grey' };
        let xDataMapped, yDataMapped;

        if (selectedAnalysisType === 'trend') {
            xDataMapped = dataForChart.map(row => row[xAxisColumnName]);
            yDataMapped = dataForChart.map(row => row[yAxisColumnName]);
            const yNumericForPlot = yDataMapped.map(v => parseFloat(v));
            if (showDataLabels) {
                traceMode = 'lines+markers+text';
                textValues = yNumericForPlot.map(v => v.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2}));
            }
            plotData = [{ x: xDataMapped, y: yNumericForPlot, type: 'scatter', mode: traceMode, text: textValues, textposition: textPosition, textfont: textFont, name: `${yAxisColumnName} vs ${xAxisColumnName}` }];
            layout.xaxis.title = xAxisColumnName;
            layout.yaxis.title = yAxisColumnName;
            layout.title.text = `Tendencia de ${yAxisColumnName} por ${xAxisColumnName}`;
        } else if (selectedAnalysisType === 'comparison') {
            const groupedData = {};
            dataForChart.forEach(row => {
                const category = row[categoryColumnName];
                const value = parseFloat(row[valueColumnName]);
                if (!isNaN(value)) {
                    groupedData[category] = (groupedData[category] || 0) + value;
                }
            });
            xDataMapped = Object.keys(groupedData);
            yDataMapped = xDataMapped.map(cat => groupedData[cat]);
            const yNumericForPlot = yDataMapped;
            textPosition = 'outside';
            if (showDataLabels) {
                textValues = yNumericForPlot.map(v => v.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2}));
            }
            plotData = [{ x: xDataMapped, y: yNumericForPlot, type: 'bar', text: textValues, textposition: textPosition, textfont: textFont, name: `${valueColumnName} por ${categoryColumnName}` }];
            layout.xaxis.title = categoryColumnName;
            layout.yaxis.title = `Suma de ${valueColumnName}`;
            layout.title.text = `Comparación de Suma de ${valueColumnName} por ${categoryColumnName}`;
        } else if (selectedAnalysisType === 'distribution') {
            xDataMapped = dataForChart.map(row => row[dataColumnName]);
            const xNumericForPlot = xDataMapped.map(v => parseFloat(v));
            plotData = [{ x: xNumericForPlot, type: 'histogram', name: `Distribución de ${dataColumnName}` }];
            layout.xaxis.title = dataColumnName;
            layout.xaxis.tickangle = 0;
            layout.yaxis.title = "Frecuencia";
            layout.title.text = `Distribución de ${dataColumnName}`;
        }

        Plotly.newPlot(chartContainer, plotData, layout, {responsive: true})
            .then(function(gd) {
                if (gd && gd.offsetParent !== null) {
                    Plotly.Plots.resize(gd);
                }
            })
            .catch(function(err) {
                console.error("Error al dibujar el gráfico con Plotly:", err);
                if(chartContainer) chartContainer.innerHTML = 'Error al generar el gráfico.';
            });

        let summaryOutput;
        if (selectedAnalysisType === 'trend') {
            summaryOutput = generateDescriptiveSummary(selectedAnalysisType, xDataMapped, yDataMapped, xAxisColumnName, yAxisColumnName, null,null,null, dataForChart);
        } else if (selectedAnalysisType === 'comparison') {
             summaryOutput = generateDescriptiveSummary(selectedAnalysisType, xDataMapped, yDataMapped, categoryColumnName, valueColumnName, categoryColumnName, valueColumnName, null, dataForChart);
        } else if (selectedAnalysisType === 'distribution') {
            const originalDistributionData = dataForChart.map(row => row[dataColumnName]);
            summaryOutput = generateDescriptiveSummary(selectedAnalysisType, originalDistributionData, null, dataColumnName, null, null, null, dataColumnName, dataForChart);
        }
        if(descriptiveSummaryElement) descriptiveSummaryElement.innerHTML = `<span class="icon">📝</span> <strong>Resumen:</strong> ${summaryOutput}`;

        if (availableHeaders.length > 0) {
            displayDataTable(availableHeaders, dataForChart);
        }
    }

    // --- Event Listeners para Botones (Existentes) ---
    if (selectFileButton && fileInput) {
        selectFileButton.addEventListener('click', () => {
            if (fileInput) {
                fileInput.click();
            } else {
                console.error("app.js: fileInput es NULL DENTRO del listener del botón 'Seleccionar Archivo'.");
            }
        });
        console.log("app.js: Event listener para selectFileButton AÑADIDO.");
    } else {
        console.error("app.js: No se pudo añadir listener a selectFileButton. Elemento(s) no encontrado(s).");
    }

    if(changeFileButton && fileInput) {
        changeFileButton.addEventListener('click', () => { resetToUploadView(); if(fileInput) fileInput.click(); });
    }

    if(fileInput) {
        fileInput.addEventListener('change', (event) => {
            const files = event.target.files;
            if (files.length > 0) {
                const selectedFile = files[0];
                currentFileName = selectedFile.name;
                if (fileNameDisplay) fileNameDisplay.textContent = currentFileName;
                if (fileLoadedStatus) fileLoadedStatus.style.display = 'block';

                resetApplicationStateBeforeNewFile(); // Limpia estados, incluyendo dataProcessingInfoElement

                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        let parsedResult;
                        if (selectedFile.name.toLowerCase().endsWith('.csv')) {
                            const fileContent = e.target.result;
                            parsedResult = parseCSVText(fileContent);
                        } else if (selectedFile.name.toLowerCase().endsWith('.xlsx') || selectedFile.name.toLowerCase().endsWith('.xls')) {
                            if (typeof XLSX === 'undefined') {
                                alert('La librería para leer archivos Excel (SheetJS) no se ha cargado correctamente.');
                                console.error("SheetJS (XLSX) no está definido.");
                                resetFileSelection(); // resetFileSelection también debería limpiar dataProcessingInfoElement si lo añadimos allí
                                return;
                            }
                            const arrayBuffer = e.target.result;
                            parsedResult = parseExcelData(arrayBuffer);
                        } else {
                            alert('Formato de archivo no soportado. Por favor, sube un archivo CSV o Excel.');
                            resetFileSelection();
                            return;
                        }
                        // ---- INICIO DE MODIFICACIÓN PARA TRUNCAMIENTO ----
                        availableHeaders = parsedResult.headers;
                        let originalRowCount = parsedResult.data.length;

                        // Limpiar mensaje previo (aunque resetApplicationStateBeforeNewFile ya debería hacerlo)
                        if (dataProcessingInfoElement) dataProcessingInfoElement.textContent = ''; 

                        if (originalRowCount > MAX_ROWS_TO_PROCESS) {
                            parsedCsvData = parsedResult.data.slice(0, MAX_ROWS_TO_PROCESS);
                            if (dataProcessingInfoElement) {
                                dataProcessingInfoElement.textContent = `Nota: Su archivo original contiene ${originalRowCount.toLocaleString()} filas. Para optimizar el rendimiento en esta demo, se han procesado las primeras ${MAX_ROWS_TO_PROCESS.toLocaleString()} filas.`;
                            }
                            console.log(`Datos truncados: ${originalRowCount} filas originales, procesadas ${parsedCsvData.length}.`);
                        } else {
                            parsedCsvData = parsedResult.data;
                            // Opcional: Mensaje si no se trunca y el elemento existe
                            // if (dataProcessingInfoElement && originalRowCount > 0) {
                            //    dataProcessingInfoElement.textContent = `Se procesaron ${originalRowCount.toLocaleString()} filas.`;
                            // }
                        }
                        // ---- FIN DE MODIFICACIÓN PARA TRUNCAMIENTO ----

                        if (availableHeaders.length > 0 && parsedCsvData.length > 0) { // Comprobar también parsedCsvData por si el archivo original estaba vacío
                            inferredColumnTypes = inferColumnTypes(parsedCsvData, availableHeaders);
                            populateColumnSelectors(availableHeaders);
                            updateSmartSuggestion(availableHeaders, inferredColumnTypes);
                        } else {
                            alert('No se pudieron leer encabezados o datos del archivo, o el archivo está vacío después del procesamiento. Asegúrate de que el formato sea correcto.');
                            updateSmartSuggestion([], {});
                             // Resetear a la vista de carga si no hay nada que procesar
                            resetToUploadView(); // Asegura que la UI refleje que no hay datos
                            return; // Detener la ejecución adicional si no hay datos válidos
                        }

                        if (dataUploadSection) dataUploadSection.style.display = 'none';
                        if (analysisDefinitionSection) analysisDefinitionSection.style.display = 'block';
                        if (smartSuggestionBox) smartSuggestionBox.style.display = 'block';
                        if (manualAnalysisForm) manualAnalysisForm.style.display = 'none';

                    } catch (error) {
                        console.error("Error procesando el archivo en reader.onload:", error);
                        alert(`Ocurrió un error al procesar el archivo: ${error.message}`);
                        resetFileSelection(); // resetFileSelection también debería limpiar dataProcessingInfoElement
                        if (dataProcessingInfoElement) dataProcessingInfoElement.textContent = ''; // Doble seguro
                    }
                };
                reader.onerror = () => {
                    console.error("app.js: FileReader onerror - Error al leer el archivo.");
                    alert('Error al leer el archivo.');
                    resetFileSelection();
                };

                if (selectedFile.name.toLowerCase().endsWith('.csv')) {
                    reader.readAsText(selectedFile);
                } else if (selectedFile.name.toLowerCase().endsWith('.xlsx') || selectedFile.name.toLowerCase().endsWith('.xls')) {
                    reader.readAsArrayBuffer(selectedFile);
                } else {
                    alert('Formato de archivo no soportado.');
                    resetFileSelection();
                }
            }
        });
    }

    if(useSuggestionButton) { useSuggestionButton.addEventListener('click', () => { /* ... (tu código completo) ... */
        if (currentSmartSuggestion && currentSmartSuggestion.type && analysisTypeSelect && manualAnalysisForm && smartSuggestionBox && typeof displayDynamicFields === 'function') {
            analysisTypeSelect.value = currentSmartSuggestion.type;
            displayDynamicFields();
            if (currentSmartSuggestion.type === 'trend' && currentSmartSuggestion.xCol && currentSmartSuggestion.yCol) {
                const xAxisTrendEl = document.getElementById('xAxisTrend');
                const yAxisTrendEl = document.getElementById('yAxisTrend');
                if(xAxisTrendEl) xAxisTrendEl.value = currentSmartSuggestion.xCol;
                if(yAxisTrendEl) yAxisTrendEl.value = currentSmartSuggestion.yCol;
            } else if (currentSmartSuggestion.type === 'comparison' && currentSmartSuggestion.catCol && currentSmartSuggestion.valCol) {
                const catAxisCompEl = document.getElementById('categoryAxisComparison');
                const valAxisCompEl = document.getElementById('valueAxisComparison');
                if(catAxisCompEl) catAxisCompEl.value = currentSmartSuggestion.catCol;
                if(valAxisCompEl) valAxisCompEl.value = currentSmartSuggestion.valCol;
            } else if (currentSmartSuggestion.type === 'distribution' && currentSmartSuggestion.dataCol) {
                const dataColDistEl = document.getElementById('dataColumnDistribution');
                if(dataColDistEl) dataColDistEl.value = currentSmartSuggestion.dataCol;
            }
            manualAnalysisForm.style.display = 'block';
            smartSuggestionBox.style.display = 'none';
        } else {
            alert("No hay una sugerencia activa para aplicar o la sugerencia es inválida.");
            if (manualAnalysisForm) manualAnalysisForm.style.display = 'block';
            if (smartSuggestionBox) smartSuggestionBox.style.display = 'none';
        }
    });}

    if(manualConfigButton) { manualConfigButton.addEventListener('click', () => { /* ... (tu código completo) ... */
        if (smartSuggestionBox) smartSuggestionBox.style.display = 'none';
        if (manualAnalysisForm) manualAnalysisForm.style.display = 'block';
        if (analysisTypeSelect) analysisTypeSelect.value = '';
        displayDynamicFields();
    });}

    if(analysisTypeSelect) { analysisTypeSelect.addEventListener('change', displayDynamicFields); }

    if(addFilterButton) { addFilterButton.addEventListener('click', () => { /* ... (tu código completo) ... */
        if (availableHeaders.length === 0) {
            alert("Carga un archivo primero para poder definir filtros basados en sus columnas.");
            return;
        }
        const filterRow = document.createElement('div');
        filterRow.className = 'filter-row';
        const columnSelectElement = document.createElement('select');
        columnSelectElement.className = 'filter-column';
        columnSelectElement.setAttribute('aria-label', 'Columna para el filtro');
        const defaultColOpt = document.createElement('option');
        defaultColOpt.value = ""; defaultColOpt.textContent = "-- Columna --"; columnSelectElement.appendChild(defaultColOpt);
        availableHeaders.forEach(header => {
            const option = document.createElement('option'); option.value = header; option.textContent = header; columnSelectElement.appendChild(option);
        });
        const conditionSelect = document.createElement('select');
        conditionSelect.className = 'filter-condition';
        conditionSelect.setAttribute('aria-label', 'Condición del filtro');
        filterConditions.forEach(cond => {
            const option = document.createElement('option'); option.value = cond.value; option.textContent = cond.text; conditionSelect.appendChild(option);
        });
        const valueContainer = document.createElement('span');
        valueContainer.className = 'filter-value-container';
        columnSelectElement.addEventListener('change', () => updateFilterValueElement(filterRow, columnSelectElement.value, conditionSelect.value));
        conditionSelect.addEventListener('change', () => updateFilterValueElement(filterRow, columnSelectElement.value, conditionSelect.value));
        const removeButton = document.createElement('button');
        removeButton.type = 'button'; removeButton.className = 'remove-filter-button'; removeButton.textContent = '✕';
        removeButton.setAttribute('aria-label', 'Quitar este filtro');
        removeButton.addEventListener('click', () => { filterRow.remove(); });
        filterRow.appendChild(columnSelectElement);
        filterRow.appendChild(conditionSelect);
        filterRow.appendChild(valueContainer);
        filterRow.appendChild(removeButton);
        if (activeFiltersContainer) activeFiltersContainer.appendChild(filterRow);
        updateFilterValueElement(filterRow, columnSelectElement.value, conditionSelect.value);
    });}

    if(generateAnalysisButton) { generateAnalysisButton.addEventListener('click', (event) => { /* ... (tu código completo) ... */
        event.preventDefault();
        if(!analysisTypeSelect) return;
        const selectedAnalysisType = analysisTypeSelect.value;
        if (!selectedAnalysisType) {
            alert("Por favor, selecciona un Tipo de Análisis.");
            return;
        }
        const filterDefinitions = collectFilterDefinitions();
        const filteredData = applyAllFilters(parsedCsvData, filterDefinitions); // parsedCsvData ya está (potencialmente) truncado

        if (filteredData.length === 0 && parsedCsvData.length > 0 && filterDefinitions.length > 0) {
            if(chartContainer) chartContainer.innerHTML = 'Ningún dato coincide con los filtros aplicados.';
            if(descriptiveSummaryElement) descriptiveSummaryElement.innerHTML = `<span class="icon">📝</span> <strong>Resumen:</strong> No hay datos para analizar después de aplicar los filtros.`;
            if(dataTableContainer) dataTableContainer.innerHTML = '<p style="text-align:center; padding: 20px; color: #777;">Ningún dato coincide con los filtros aplicados.</p>';
            if (aiSummaryResultContainer) aiSummaryResultContainer.style.display = 'none';
            if (analysisDefinitionSection) analysisDefinitionSection.style.display = 'none';
            if (dashboardResultsSection) dashboardResultsSection.style.display = 'block';
            if(dashboardFileName) dashboardFileName.textContent = currentFileName + " (filtrado)";
            if(chartContainer) Plotly.purge(chartContainer);
            lastChartConfig = null;
            return;
        } else if (filteredData.length === 0 && parsedCsvData.length === 0){ // Esto cubre el caso de archivo vacío o no parseable
            if(chartContainer) chartContainer.innerHTML = 'No hay datos cargados o procesados para analizar.';
            if(descriptiveSummaryElement) descriptiveSummaryElement.innerHTML = `<span class="icon">📝</span> <strong>Resumen:</strong> No hay datos cargados o procesados.`;
            if(dataTableContainer) dataTableContainer.innerHTML = '<p style="text-align:center; padding: 20px; color: #777;">No hay datos cargados o procesados.</p>';
            if (aiSummaryResultContainer) aiSummaryResultContainer.style.display = 'none';
            if (analysisDefinitionSection) analysisDefinitionSection.style.display = 'none';
            if (dashboardResultsSection) dashboardResultsSection.style.display = 'block';
            if(dashboardFileName) dashboardFileName.textContent = currentFileName;
            if(chartContainer) Plotly.purge(chartContainer);
            lastChartConfig = null;
            return;
        }

        let xAxisColumnName = "", yAxisColumnName = "", categoryColumnName = "", valueColumnName = "", dataColumnName = "";
        if (selectedAnalysisType === 'trend') {
            xAxisColumnName = document.getElementById('xAxisTrend').value;
            yAxisColumnName = document.getElementById('yAxisTrend').value;
            if (!xAxisColumnName || !yAxisColumnName) {
                if(chartContainer) chartContainer.textContent = 'Configuración incompleta.';
                if(dataTableContainer) dataTableContainer.innerHTML = '';
                alert("Por favor, selecciona las columnas para Eje X y Eje Y.");
                return;
            }
        } else if (selectedAnalysisType === 'comparison') {
            categoryColumnName = document.getElementById('categoryAxisComparison').value;
            valueColumnName = document.getElementById('valueAxisComparison').value;
            if (!categoryColumnName || !valueColumnName) {
                if(chartContainer) chartContainer.textContent = 'Configuración incompleta.';
                if(dataTableContainer) dataTableContainer.innerHTML = '';
                alert("Por favor, selecciona las columnas para Categoría y Valor.");
                return;
            }
        } else if (selectedAnalysisType === 'distribution') {
            dataColumnName = document.getElementById('dataColumnDistribution').value;
            if (!dataColumnName) {
                if(chartContainer) chartContainer.textContent = 'Configuración incompleta.';
                if(dataTableContainer) dataTableContainer.innerHTML = '';
                alert("Por favor, selecciona la columna de datos para la distribución.");
                return;
            }
        } else {
            if(chartContainer) chartContainer.textContent = 'Tipo de análisis no implementado.';
            if(dataTableContainer) dataTableContainer.innerHTML = '';
            return;
        }

        lastChartConfig = {
            selectedAnalysisType,
            xAxisColumnName, yAxisColumnName,
            categoryColumnName, valueColumnName,
            dataColumnName,
            filters: filterDefinitions
        };

        if (analysisDefinitionSection) analysisDefinitionSection.style.display = 'none';
        if (dashboardResultsSection) dashboardResultsSection.style.display = 'block';
        if(dashboardFileName) dashboardFileName.textContent = currentFileName + (filterDefinitions.length > 0 ? " (filtrado)" : "");

        renderOrUpdateChart(lastChartConfig, filteredData);
    });}

    if (toggleDataLabelsCheckbox) { /* ... (tu código completo) ... */
        toggleDataLabelsCheckbox.addEventListener('change', () => {
            if (lastChartConfig && dashboardResultsSection && dashboardResultsSection.style.display === 'block' && parsedCsvData.length > 0) {
                const dataToRender = applyAllFilters(parsedCsvData, lastChartConfig.filters || []);
                renderOrUpdateChart(lastChartConfig, dataToRender);
            }
        });
    }

    if(downloadChartButton) { /* ... (tu código completo) ... */
        downloadChartButton.addEventListener('click', () => {
            if (parsedCsvData.length > 0 && chartContainer && chartContainer.querySelector('.plot-container')) {
                let filename = 'grafico_SIR-Analytics';
                if (currentFileName) {
                    const nameParts = currentFileName.split('.');
                    if (nameParts.length > 1) nameParts.pop();
                    filename = `grafico_${nameParts.join('.')}`;
                }
                Plotly.downloadImage(chartContainer, { format: 'png', width: 1000, height: 700, filename: filename });
            } else {
                alert("No hay ningún gráfico generado para descargar o no hay datos cargados.");
            }
        });
    }

    if(downloadDataButton) { /* ... (tu código completo) ... */
        downloadDataButton.addEventListener('click', () => {
            const filterDefinitions = collectFilterDefinitions();
            const dataToDownload = applyAllFilters(parsedCsvData, filterDefinitions); // Usa parsedCsvData que ya puede estar truncado
            if (!dataToDownload || dataToDownload.length === 0) {
                alert("No hay datos (o datos filtrados) para descargar.");
                return;
            }
            let csvContent = availableHeaders.join(';') + '\r\n';
            dataToDownload.forEach(rowObject => {
                const rowValues = availableHeaders.map(header => {
                    let cellValue = rowObject[header] === null || rowObject[header] === undefined ? '' : String(rowObject[header]);
                    if (cellValue.includes(';')) { cellValue = `"${cellValue.replace(/"/g, '""')}"`; }
                    return cellValue;
                });
                csvContent += rowValues.join(';') + '\r\n';
            });
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            let filename = 'datos_SIR-Analytics.csv';
            if (currentFileName) {
                const nameParts = currentFileName.split('.');
                if (nameParts.length > 1) nameParts.pop();
                filename = `datos_${nameParts.join('.')}${filterDefinitions.length > 0 ? '_filtrado' : ''}.csv`;
            }
            const link = document.createElement("a");
            if (link.download !== undefined) {
                const url = URL.createObjectURL(blob);
                link.setAttribute("href", url); link.setAttribute("download", filename);
                link.style.visibility = 'hidden'; document.body.appendChild(link);
                link.click(); document.body.removeChild(link);
                URL.revokeObjectURL(url);
            } else { alert("La descarga directa no es soportada por tu navegador."); }
        });
    }

    if(performNewAnalysisButton) { /* ... (tu código completo) ... */
        performNewAnalysisButton.addEventListener('click', () => {
            resetToUploadView();
        });
    }

    // --- Event Listener para el Botón de Resumen con IA ---
    if (getAISummaryButton && aiSummaryResultContainer && aiSummaryTextElement && descriptiveSummaryElement) {
        getAISummaryButton.addEventListener('click', async () => {
            if (!lastChartConfig || !parsedCsvData || parsedCsvData.length === 0) {
                alert("Primero genera un gráfico para obtener un análisis con IA.");
                if(aiSummaryResultContainer) aiSummaryResultContainer.style.display = 'none';
                return;
            }

            if(aiSummaryResultContainer) aiSummaryResultContainer.style.display = 'block';
            if(aiSummaryTextElement) aiSummaryTextElement.innerHTML = '<em>Conectando con la IA y generando análisis... Por favor, espera.</em> ✨';
            if(getAISummaryButton) getAISummaryButton.disabled = true;

            const { selectedAnalysisType } = lastChartConfig;
            let xAxisLabel = "N/A";
            let yAxisLabel = "N/A";

            if (selectedAnalysisType === 'trend') {
                xAxisLabel = lastChartConfig.xAxisColumnName || "Eje X";
                yAxisLabel = lastChartConfig.yAxisColumnName || "Eje Y";
            } else if (selectedAnalysisType === 'comparison') {
                xAxisLabel = lastChartConfig.categoryColumnName || "Categoría";
                yAxisLabel = `Suma de ${lastChartConfig.valueColumnName || "Valor"}`;
            } else if (selectedAnalysisType === 'distribution') {
                xAxisLabel = lastChartConfig.dataColumnName || "Datos";
                yAxisLabel = "Frecuencia";
            }

            const keyInsightsFromBasicSummary = descriptiveSummaryElement.textContent.replace(/📝 Resumen: /g, '').trim();

            const dataForAI = {
                chartType: selectedAnalysisType,
                xAxisColumnName: xAxisLabel,
                yAxisColumnName: yAxisLabel,
                keyDataInsights: keyInsightsFromBasicSummary
            };

            try {
                const response = await fetch('/api/get-ai-summary', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(dataForAI),
                });

                if(getAISummaryButton) getAISummaryButton.disabled = false;

                if (!response.ok) {
                    const errorResult = await response.json().catch(() => ({ error: response.statusText })); // Fallback si el cuerpo no es JSON
                    console.error("Error desde la función serverless:", errorResult.error || response.statusText);
                    if(aiSummaryTextElement) aiSummaryTextElement.textContent = `Error al obtener el análisis de IA: ${errorResult.error || response.statusText}`;
                    return;
                }

                const result = await response.json();
                if(aiSummaryTextElement) aiSummaryTextElement.innerHTML = result.aiSummary.replace(/\n/g, '<br>');

            } catch (error) {
                if(getAISummaryButton) getAISummaryButton.disabled = false;
                console.error("Error al llamar a la función para el resumen de IA:", error);
                if(aiSummaryTextElement) aiSummaryTextElement.textContent = "Error de conexión al intentar obtener el análisis de IA. Verifica tu conexión o inténtalo más tarde.";
            }
        });
        console.log("app.js: Event listener para getAISummaryButton AÑADIDO.");
    } else {
        console.warn("app.js: No se pudieron encontrar todos los elementos para la funcionalidad de Resumen con IA (getAISummaryButton, aiSummaryResultContainer, aiSummaryTextElement, descriptiveSummaryElement).");
    }

    // --- Funciones de Reseteo ---
    function resetApplicationStateBeforeNewFile() {
        availableHeaders = []; 
        // parsedCsvData ya se limpia/redefine al cargar nuevo archivo, no es necesario aquí.
        resetColumnSelectors();
        if(analysisTypeSelect) analysisTypeSelect.value = '';
        displayDynamicFields();
        if(smartSuggestionText) smartSuggestionText.innerHTML = `<span class="icon">💡</span> <strong>Sugerencia:</strong> (Esperando datos del archivo para generar sugerencia...)`;
        lastChartConfig = null;
        if(toggleDataLabelsCheckbox) toggleDataLabelsCheckbox.checked = false;
        currentSmartSuggestion = null;
        if(activeFiltersContainer) activeFiltersContainer.innerHTML = '';
        inferredColumnTypes = {};
        if(dataTableContainer) dataTableContainer.innerHTML = '';
        currentTableData = [];
        currentSortColumnKey = null;
        
        // ---- INICIO DE MODIFICACIÓN ----
        if (aiSummaryResultContainer) aiSummaryResultContainer.style.display = 'none';
        if (aiSummaryTextElement) aiSummaryTextElement.innerHTML = '';
        if (dataProcessingInfoElement) dataProcessingInfoElement.textContent = ''; // Limpiar mensaje de procesamiento de datos
        // ---- FIN DE MODIFICACIÓN ----
    }

    function resetToUploadView() {
        currentFileName = '';
        if(fileInput) fileInput.value = null; // Esto es importante para permitir volver a seleccionar el mismo archivo
        resetApplicationStateBeforeNewFile(); // Llama a la función que ya limpia la mayoría de las cosas
        
        if(dashboardResultsSection) dashboardResultsSection.style.display = 'none';
        if(analysisDefinitionSection) analysisDefinitionSection.style.display = 'none';
        if(manualAnalysisForm) manualAnalysisForm.style.display = 'none';
        if(smartSuggestionBox) smartSuggestionBox.style.display = 'block'; // Mostrar sugerencias de nuevo
        if(dataUploadSection) dataUploadSection.style.display = 'block'; // Mostrar área de carga
        if(fileLoadedStatus) fileLoadedStatus.style.display = 'none'; // Ocultar estado de archivo cargado
        if(fileNameDisplay) fileNameDisplay.textContent = '';
        
        if(chartContainer) {
            Plotly.purge(chartContainer);
            chartContainer.textContent = 'Aquí se mostrará el gráfico interactivo (Plotly.js).';
        }
        if(descriptiveSummaryElement) descriptiveSummaryElement.innerHTML = `<span class="icon">📝</span> <strong>Resumen:</strong> (Este es un resumen descriptivo básico generado automáticamente.)`;
        // No es necesario limpiar aiSummaryResultContainer y aiSummaryTextElement de nuevo aquí,
        // ya que resetApplicationStateBeforeNewFile() lo hace.
        // dataProcessingInfoElement también es limpiado por resetApplicationStateBeforeNewFile()
    }

    console.log("app.js: Aplicación SIR - Analytics completamente inicializada y todos los listeners configurados (o intentados).");
});
