javascript: (async () => {
    // === 基本常數定義 ===
    const API_URLS = {
        test: 'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisb/api/caseQuery/query',
        prod: 'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisb/api/caseQuery/query'
    };

    const TOKEN_STORAGE_KEY = 'euisToken';
    const A17_TEXT_SETTINGS_STORAGE_KEY = 'kgilifeQueryTool_A17TextSettings_vFinal';
    const TOOL_MAIN_CONTAINER_ID = 'kgilifeQueryToolMainContainer_vFinal';
    const Z_INDEX = {
        OVERLAY: 2147483640,
        MAIN_UI: 2147483630,
        NOTIFICATION: 2147483647
    };

    const QUERYABLE_FIELD_DEFINITIONS = [{
            queryApiKey: 'receiptNumber',
            queryDisplayName: '送金單號碼',
            color: '#007bff'
        },
        {
            queryApiKey: 'applyNumber',
            queryDisplayName: '受理號碼',
            color: '#6f42c1'
        },
        {
            queryApiKey: 'policyNumber',
            queryDisplayName: '保單號碼',
            color: '#28a745'
        },
        {
            queryApiKey: 'approvalNumber',
            queryDisplayName: '確認書編號',
            color: '#fd7e14'
        },
        {
            queryApiKey: 'insuredId',
            queryDisplayName: '被保人ＩＤ',
            color: '#17a2b8'
        }
    ];

    const FIELD_DISPLAY_NAMES_MAP = {
        applyNumber: '受理號碼',
        policyNumber: '保單號碼',
        approvalNumber: '確認書編號',
        receiptNumber: '送金單',
        insuredId: '被保人ＩＤ',
        statusCombined: '狀態',
        mainStatus: '主狀態',
        subStatus: '次狀態',
        uwApproverUnit: '分公司',
        uwApprover: '核保員',
        approvalUser: '覆核',
        _queriedValue_: '查詢值',
        NO: '序號',
        _apiQueryStatus: '查詢結果'
    };

    const ALL_DISPLAY_FIELDS_API_KEYS_MAIN = [
        'applyNumber', 'policyNumber', 'approvalNumber', 'receiptNumber', 'insuredId',
        'statusCombined', 'uwApproverUnit', 'uwApprover', 'approvalUser'
    ];

    const UNIT_CODE_MAPPINGS = {
        H: '核保部',
        B: '北一',
        C: '台中',
        K: '高雄',
        N: '台南',
        P: '北二',
        T: '桃竹',
        G: '保作'
    };

    const A17_UNIT_BUTTONS_DEFS = [{
            id: 'H',
            label: 'H-總公司',
            color: '#007bff'
        },
        {
            id: 'B',
            label: 'B-北一',
            color: '#28a745'
        },
        {
            id: 'P',
            label: 'P-北二',
            color: '#ffc107'
        },
        {
            id: 'T',
            label: 'T-桃竹',
            color: '#17a2b8'
        },
        {
            id: 'C',
            label: 'C-台中',
            color: '#fd7e14'
        },
        {
            id: 'N',
            label: 'N-台南',
            color: '#6f42c1'
        },
        {
            id: 'K',
            label: 'K-高雄',
            color: '#e83e8c'
        },
        {
            id: 'UNDEF',
            label: '查無單位',
            color: '#546e7a'
        }
    ];

    const UNIT_MAP_FIELD_API_KEY = 'uwApproverUnit';
    const A17_DEFAULT_TEXT_CONTENT = "DEAR,\n\n依據【管理報表：A17 新契約異常帳務】所載內容，報表中列示之送金單號碼，涉及多項帳務異常情形，例如：溢繳、短收、取消件需退費、以及無相對應之案件等問題。\n\n本週我們已逐筆查詢該等異常帳務，結果顯示，這些送金單應對應至下表所列之新契約案件。為利後續帳務處理，敬請協助確認各案件之實際帳務狀況，並如有需調整或處理事項，請一併協助辦理，謝謝。";

    // === 全域變數 ===
    let CURRENT_API_URL = '';
    let apiAuthToken = null;
    let selectedQueryDefinitionGlobal = QUERYABLE_FIELD_DEFINITIONS[0];
    let isEditMode = false;
    let dragState = {
        dragging: false,
        startX: 0,
        startY: 0,
        initialX: 0,
        initialY: 0
    };
    let a17ButtonLongPressTimer = null;

    // === 狀態管理器 ===
    const StateManager = {
        originalQueryResults: [],
        baseA17MasterData: [],
        csvImport: {
            fileName: '',
            rawHeaders: [],
            rawData: [],
            selectedColForQueryName: null,
            selectedColsForA17Merge: [],
            isA17CsvPrepared: false,
        },
        a17Mode: {
            isActive: false,
            selectedUnits: new Set(),
            textSettings: {
                mainContent: A17_DEFAULT_TEXT_CONTENT,
                mainFontSize: 12,
                mainLineHeight: 1.5,
                mainFontColor: '#333333',
                dateFontSize: 8,
                dateLineHeight: 1.2,
                dateFontColor: '#555555',
                genDateOffset: -3,
                compDateOffset: 0,
            },
        },
        currentTable: {
            sortDirections: {},
            currentHeaders: [],
            mainUIElement: null,
            tableBodyElement: null,
            tableHeadElement: null,
            a17UnitButtonsContainer: null,
        },
        history: [],

        loadA17Settings() {
            const saved = localStorage.getItem(A17_TEXT_SETTINGS_STORAGE_KEY);
            if (saved) {
                try {
                    const parsed = JSON.parse(saved);
                    for (const key in this.a17Mode.textSettings) {
                        if (parsed.hasOwnProperty(key)) {
                            this.a17Mode.textSettings[key] = parsed[key];
                        }
                    }
                } catch (e) {
                    console.error("載入A17文本設定失敗:", e);
                }
            }
        },

        saveA17Settings() {
            try {
                localStorage.setItem(A17_TEXT_SETTINGS_STORAGE_KEY, JSON.stringify(this.a17Mode.textSettings));
            } catch (e) {
                console.error("儲存A17文本設定失敗:", e);
            }
        },

        pushSnapshot(description = '操作') {
            const snapshot = structuredClone({
                originalQueryResults: this.originalQueryResults,
                baseA17MasterData: this.baseA17MasterData,
                csvImport: this.csvImport,
                a17Mode: this.a17Mode
            });
            this.history.push({
                description,
                snapshot,
                timestamp: Date.now()
            });
            if (this.history.length > 10) {
                this.history.shift();
            }
        },

        undo() {
            if (this.history.length === 0) {
                displaySystemNotification("沒有更多操作可復原", true);
                return;
            }
            const lastState = this.history.pop();
            Object.assign(this, lastState.snapshot);
            populateTableRows(this.a17Mode.isActive ? this.baseA17MasterData : this.originalQueryResults);
            displaySystemNotification(`已復原：${lastState.description}`, false);
        }
    };

    // === 工具函數 ===
    function escapeHtml(unsafe) {
        if (typeof unsafe !== 'string') {
            return unsafe === null || unsafe === undefined ? '' : String(unsafe);
        }
        return unsafe.replace(/[&<>"'`]/g, m => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;',
            '`': '&#x60;'
        })[m] || m);
    }

    function displaySystemNotification(message, isError = false, duration = 3000) {
        const id = TOOL_MAIN_CONTAINER_ID + '_Notification';
        document.getElementById(id)?.remove();
        const n = document.createElement('div');
        n.id = id;
        n.style.cssText = `position:fixed;top:20px;right:20px;background-color:${isError ? '#dc3545' : '#28a745'};color:white;padding:10px 15px;border-radius:5px;z-index:${Z_INDEX.NOTIFICATION};font-size:14px;font-family:'Microsoft JhengHei',Arial,sans-serif;box-shadow:0 2px 10px rgba(0,0,0,0.2);transform:translateX(calc(100% + 25px));transition:transform 0.3s ease-in-out;display:flex;align-items:center;cursor:pointer;`;
        const i = document.createElement('span');
        i.style.marginRight = '8px';
        i.style.fontSize = '16px';
        i.innerHTML = isError ? '&#x26A0;' : '&#x2714;';
        n.appendChild(i);
        n.appendChild(document.createTextNode(message));
        const closeBtn = document.createElement('span');
        closeBtn.innerHTML = '&times;';
        closeBtn.style.cssText = 'margin-left:10px;font-size:18px;cursor:pointer;';
        closeBtn.onclick = () => n.remove();
        n.appendChild(closeBtn);
        document.body.appendChild(n);
        setTimeout(() => n.style.transform = 'translateX(0)', 50);
        setTimeout(() => {
            n.style.transform = 'translateX(calc(100% + 25px))';
            setTimeout(() => n.remove(), 300);
        }, duration);
    }

    function createDialogBase(idSuffix, contentHtml, minWidth = '350px', maxWidth = '600px', customStyles = '') {
        const id = TOOL_MAIN_CONTAINER_ID + idSuffix;
        document.getElementById(id + '_overlay')?.remove();
        const overlay = document.createElement('div');
        overlay.id = id + '_overlay';
        overlay.style.cssText = `position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:${Z_INDEX.OVERLAY};display:flex;align-items:center;justify-content:center;font-family:'Microsoft JhengHei',Arial,sans-serif;backdrop-filter:blur(2px);`;
        const dialog = document.createElement('div');
        dialog.id = id + '_dialog';
        dialog.style.cssText = `background:#fff;padding:20px 25px;border-radius:8px;box-shadow:0 5px 20px rgba(0,0,0,0.25);min-width:${minWidth};max-width:${maxWidth};width:auto;animation:qtDialogAppear 0.2s ease-out;${customStyles}`;
        dialog.innerHTML = contentHtml;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        const styleEl = document.createElement('style');
        styleEl.textContent = `
        @keyframes qtDialogAppear{from{opacity:0;transform:scale(0.95)}to{opacity:1;transform:scale(1)}}
        .qt-dialog-btn{border:none;padding:8px 15px;border-radius:4px;font-size:13px;cursor:pointer;transition:opacity 0.2s ease,transform 0.1s ease;font-weight:500;margin-left:8px;}
        .qt-dialog-btn:hover{opacity:0.85;}
        .qt-dialog-btn:active{transform:scale(0.98);}
        .qt-dialog-btn-blue{background:#007bff;color:white;}
        .qt-dialog-btn-grey{background:#6c757d;color:white;}
        .qt-dialog-btn-red{background:#dc3545;color:white;}
        .qt-dialog-btn-orange{background:#fd7e14;color:white;}
        .qt-dialog-btn-green{background:#28a745;color:white;}
        .qt-dialog-title{margin:0 0 15px 0;color:#333;font-size:18px;text-align:center;font-weight:600;}
        .qt-input,.qt-textarea,.qt-select{width:calc(100% - 18px);padding:9px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-bottom:15px;color:#333;box-sizing:border-box;}
        .qt-textarea{min-height:70px;resize:vertical;}
        .qt-dialog-flex-end{display:flex;justify-content:flex-end;margin-top:15px;}
        .qt-dialog-flex-between{display:flex;justify-content:space-between;align-items:center;margin-top:15px;}
        .qt-retry-btn{background:#17a2b8;color:white;border:none;padding:4px 8px;border-radius:3px;font-size:11px;cursor:pointer;margin:2px;}
        .qt-retry-edit-btn{background:#fd7e14;color:white;border:none;padding:4px 8px;border-radius:3px;font-size:11px;cursor:pointer;margin:2px;}
        .qt-retry-btn:hover,.qt-retry-edit-btn:hover{opacity:0.8;}
        .qt-querytype-btn{border:2px solid transparent;padding:8px 10px;border-radius:4px;font-size:13px;font-weight:500;cursor:pointer;margin:4px;transition:all 0.2s ease;}
        .qt-querytype-btn:hover{opacity:0.9;}
        .qt-unit-filter-btn{border:none;padding:6px 12px;border-radius:4px;font-size:12px;cursor:pointer;margin:3px;transition:all 0.2s ease;font-weight:500;}
        .qt-unit-filter-btn:hover{opacity:0.8;}
        .qt-unit-filter-btn.active{box-shadow:0 0 8px rgba(0,0,0,0.3);}
        .qt-table{width:100%;border-collapse:collapse;margin-top:10px;}
        .qt-table th,.qt-table td{border:1px solid #ddd;padding:8px;text-align:left;font-size:13px;}
        .qt-table th{background-color:#f8f9fa;font-weight:600;cursor:pointer;user-select:none;}
        .qt-table th:hover{background-color:#e9ecef;}
        .qt-table tbody tr:nth-child(even){background-color:#f8f9fa;}
        .qt-table tbody tr:hover{background-color:#e3f2fd;}
        .qt-editable-cell{cursor:pointer;border:1px dashed transparent;}
        .qt-editable-cell:hover{border-color:#007bff;background-color:#e3f2fd;}
        .qt-delete-btn{background:#dc3545;color:white;border:none;padding:4px 6px;border-radius:3px;font-size:10px;cursor:pointer;}
        .qt-delete-btn:hover{opacity:0.8;}
        `;
        dialog.appendChild(styleEl);

        // ESC鍵關閉
        const escListener = (e) => {
            if (e.key === 'Escape') {
                overlay.remove();
                document.removeEventListener('keydown', escListener);
            }
        };
        document.addEventListener('keydown', escListener);

        return {
            overlay,
            dialog
        };
    }

    function extractName(strVal) {
        if (!strVal || typeof strVal !== 'string') return '';
        const matchResult = strVal.match(/^[\u4e00-\u9fa5\uff0a*\u00b7\uff0e]+/);
        return matchResult ? matchResult[0] : strVal.split(' ')[0];
    }

    function getFirstLetter(unitString) {
        if (!unitString || typeof unitString !== 'string') return 'UNDEF';
        for (let i = 0; i < unitString.length; i++) {
            const char = unitString.charAt(i).toUpperCase();
            if (/[A-Z]/.test(char)) return char;
        }
        return 'UNDEF';
    }

    function formatDate(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}${m}${d}`;
    }

    // === 管理器 ===
    const EnvManager = {
        get: () => CURRENT_API_URL,
        showDialog: () => new Promise(resolve => {
            const contentHtml = `
            <h3 class="qt-dialog-title">選擇查詢環境</h3>
            <div style="display:flex; gap:10px; justify-content:center;">
                <button id="qt-env-uat" class="qt-dialog-btn qt-dialog-btn-green" style="flex-grow:1;">測試 (UAT)</button>
                <button id="qt-env-prod" class="qt-dialog-btn qt-dialog-btn-orange" style="flex-grow:1;">正式 (PROD)</button>
            </div>
            <div style="text-align:center; margin-top:15px;">
                <button id="qt-env-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button>
            </div>
            `;
            const {
                overlay
            } = createDialogBase('_EnvSelect', contentHtml, '300px', 'auto');
            const closeDialog = (value) => {
                overlay.remove();
                resolve(value);
            };
            overlay.querySelector('#qt-env-uat').onclick = () => closeDialog('test');
            overlay.querySelector('#qt-env-prod').onclick = () => closeDialog('prod');
            overlay.querySelector('#qt-env-cancel').onclick = () => closeDialog(null);
        }),
        set(env) {
            CURRENT_API_URL = API_URLS[env];
        }
    };

    const TokenManager = {
        get: () => apiAuthToken,
        set: (token) => {
            apiAuthToken = token;
            localStorage.setItem(TOKEN_STORAGE_KEY, token);
        },
        clear: () => {
            apiAuthToken = null;
            localStorage.removeItem(TOKEN_STORAGE_KEY);
        },
        init: () => {
            apiAuthToken = localStorage.getItem(TOKEN_STORAGE_KEY);
            return !!apiAuthToken;
        },
        showDialog: (attempt = 1) => new Promise(resolve => {
            const contentHtml = `
            <h3 class="qt-dialog-title">API TOKEN 設定</h3>
            <input type="password" id="qt-token-input" class="qt-input" placeholder="請輸入您的 API TOKEN">
            ${attempt > 1 ? `<p style="color:red; font-size:12px; text-align:center; margin-bottom:10px;">Token驗證失敗，請重新輸入。</p>` : ''}
            <div class="qt-dialog-flex-between">
                <button id="qt-token-skip" class="qt-dialog-btn qt-dialog-btn-orange">略過</button>
                <div>
                    <button id="qt-token-close-tool" class="qt-dialog-btn qt-dialog-btn-red">關閉工具</button>
                    <button id="qt-token-ok" class="qt-dialog-btn qt-dialog-btn-blue">${attempt > 1 ? '重試' : '確定'}</button>
                </div>
            </div>
            `;
            const {
                overlay
            } = createDialogBase('_Token', contentHtml, '320px', 'auto');
            const inputEl = overlay.querySelector('#qt-token-input');
            inputEl.focus();

            if (attempt > 2) {
                const okBtn = overlay.querySelector('#qt-token-ok');
                okBtn.disabled = true;
                okBtn.style.opacity = '0.5';
                okBtn.style.cursor = 'not-allowed';
                displaySystemNotification('Token多次驗證失敗', true, 4000);
            }

            const closeDialog = (value) => {
                overlay.remove();
                resolve(value);
            };

            overlay.querySelector('#qt-token-ok').onclick = () => closeDialog(inputEl.value.trim());
            overlay.querySelector('#qt-token-close-tool').onclick = () => closeDialog('_close_tool_');
            overlay.querySelector('#qt-token-skip').onclick = () => closeDialog('_skip_token_');
        })
    };

    const CSVManager = {
        async importCSV(file, purpose) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = e => {
                    try {
                        const text = e.target.result;
                        const lines = text.split('\n').filter(line => line.trim());
                        if (lines.length < 2) {
                            reject(new Error('CSV檔案格式不正確'));
                            return;
                        }

                        const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
                        const data = lines.slice(1).map(line => {
                            const cells = line.split(',').map(c => c.trim().replace(/"/g, ''));
                            const row = {};
                            headers.forEach((header, index) => {
                                row[header] = cells[index] || '';
                            });
                            return row;
                        });

                        StateManager.csvImport.fileName = file.name;
                        StateManager.csvImport.rawHeaders = headers;
                        StateManager.csvImport.rawData = data;

                        resolve({
                            headers,
                            data,
                            purpose
                        });
                    } catch (error) {
                        reject(error);
                    }
                };
                reader.onerror = () => reject(new Error('檔案讀取失敗'));
                reader.readAsText(file, 'UTF-8');
            });
        },

        async showColumnSelectDialog(headers, purpose, isMultiple = false) {
            return new Promise(resolve => {
                const title = purpose === 'query' ? '選擇查詢值欄位' : '選擇A17合併欄位';
                const inputType = isMultiple ? 'checkbox' : 'radio';
                const columnsHtml = headers.map((header, index) =>
                    `<label style="display:block;margin:8px 0;cursor:pointer;">
                        <input type="${inputType}" name="csvCol" value="${index}" style="margin-right:8px;">
                        ${escapeHtml(header)}
                    </label>`
                ).join('');

                const contentHtml = `
                <h3 class="qt-dialog-title">${title}</h3>
                <div style="max-height:300px;overflow-y:auto;border:1px solid #ddd;padding:10px;border-radius:4px;">
                    ${columnsHtml}
                </div>
                <div class="qt-dialog-flex-end">
                    <button id="qt-csv-col-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button>
                    <button id="qt-csv-col-ok" class="qt-dialog-btn qt-dialog-btn-blue">確定</button>
                </div>
                `;

                const {
                    overlay
                } = createDialogBase('_CSVColSelect', contentHtml, '400px', 'auto');

                const closeDialog = (value) => {
                    overlay.remove();
                    resolve(value);
                };

                overlay.querySelector('#qt-csv-col-ok').onclick = () => {
                    const selected = Array.from(overlay.querySelectorAll('input[name="csvCol"]:checked'))
                        .map(input => parseInt(input.value));
                    if (selected.length === 0) {
                        displaySystemNotification('請選擇至少一個欄位', true);
                        return;
                    }
                    closeDialog(selected);
                };

                overlay.querySelector('#qt-csv-col-cancel').onclick = () => closeDialog(null);
            });
        }
    };

    // === API查詢函數 ===
    async function performApiQuery(queryValue, queryApiKey) {
        try {
            const requestBody = {
                currentPage: 1,
                pageSize: 10,
                [queryApiKey]: queryValue
            };

            const headers = {
                'Content-Type': 'application/json'
            };

            if (apiAuthToken) {
                headers['SSO-TOKEN'] = apiAuthToken;
            }

            let retries = 1;
            while (retries >= 0) {
                try {
                    const response = await fetch(CURRENT_API_URL, {
                        method: 'POST',
                        headers: headers,
                        body: JSON.stringify(requestBody)
                    });

                    if (response.status === 401) {
                        apiAuthToken = null;
                        TokenManager.clear();
                        return {
                            error: 'token_invalid',
                            data: null
                        };
                    }

                    if (!response.ok) {
                        throw new Error(`HTTP ${response.status}`);
                    }

                    const data = await response.json();

                    if (data && data.records && data.records.length > 0) {
                        return {
                            success: true,
                            data
                        };
                    } else {
                        return {
                            success: false,
                            data: null
                        };
                    }
                } catch (error) {
                    if (retries > 0) {
                        await new Promise(r => setTimeout(r, 2000));
                        retries--;
                    } else {
                        throw error;
                    }
                }
            }
        } catch (error) {
            console.error('API查詢錯誤:', error);
            return {
                error: 'network_error',
                data: null
            };
        }
    }

    // === 查詢設定對話框 ===
    function createQuerySetupDialog() {
        return new Promise(resolve => {
            const queryButtonsHtml = QUERYABLE_FIELD_DEFINITIONS.map(def =>
                `<button class="qt-querytype-btn" data-apikey="${def.queryApiKey}" style="background-color:${def.color};color:white;">${escapeHtml(def.queryDisplayName)}</button>`
            ).join('');

            const contentHtml = `
            <h3 class="qt-dialog-title">查詢條件設定</h3>
            <div style="margin-bottom:10px;font-size:13px;color:#555;">選擇查詢欄位類型：</div>
            <div id="qt-querytype-buttons" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:15px;">
                ${queryButtonsHtml}
            </div>
            <div style="margin-bottom:5px;font-size:13px;color:#555;">輸入查詢值（多個值請用空格、逗號或換行分隔）：</div>
            <textarea id="qt-queryvalue-input" class="qt-textarea" placeholder="請輸入查詢值..."></textarea>
            <div style="margin-bottom:15px;">
                <input type="file" id="qt-csv-file-input" accept=".csv,.txt" style="display:none;">
                <button id="qt-csv-import-btn" class="qt-dialog-btn qt-dialog-btn-grey">從CSV/TXT匯入...</button>
                <span id="qt-csv-filename" style="margin-left:10px;font-size:12px;color:#666;"></span>
            </div>
            <div style="display:flex;justify-content:space-between;margin-top:15px;">
                <button id="qt-clear-input" class="qt-dialog-btn qt-dialog-btn-orange">清除所有輸入</button>
                <div>
                    <button id="qt-query-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button>
                    <button id="qt-query-ok" class="qt-dialog-btn qt-dialog-btn-blue">開始查詢</button>
                </div>
            </div>
            `;

            const {
                overlay
            } = createDialogBase('_QuerySetup', contentHtml, '450px', 'auto');

            let selectedApiKey = QUERYABLE_FIELD_DEFINITIONS[0].queryApiKey;
            const queryTypeButtons = overlay.querySelectorAll('.qt-querytype-btn');
            const queryValueInput = overlay.querySelector('#qt-queryvalue-input');
            const csvFileInput = overlay.querySelector('#qt-csv-file-input');
            const csvImportBtn = overlay.querySelector('#qt-csv-import-btn');
            const csvFilename = overlay.querySelector('#qt-csv-filename');

            function setActiveButton(apiKey) {
                queryTypeButtons.forEach(btn => {
                    const isSelected = btn.dataset.apikey === apiKey;
                    btn.style.border = isSelected ? `2px solid ${btn.style.backgroundColor}` : '2px solid transparent';
                    btn.style.boxShadow = isSelected ? `0 0 6px ${btn.style.backgroundColor}80` : 'none';
                });

                // 更新placeholder
                const selectedDef = QUERYABLE_FIELD_DEFINITIONS.find(def => def.queryApiKey === apiKey);
                if (selectedDef) {
                    queryValueInput.placeholder = `請輸入${selectedDef.queryDisplayName}...`;
                }
            }

            queryTypeButtons.forEach(btn => {
                btn.onclick = () => {
                    selectedApiKey = btn.dataset.apikey;
                    setActiveButton(selectedApiKey);
                };
            });

            setActiveButton(selectedApiKey);

            // CSV匯入處理
            csvImportBtn.onclick = () => csvFileInput.click();
            csvFileInput.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) return;

                try {
                    // 詢問用途
                    const purpose = await showCSVPurposeDialog();
                    if (!purpose) return;

                    const result = await CSVManager.importCSV(file, purpose);
                    csvFilename.textContent = file.name;

                    if (purpose === 'query') {
                        // 選擇查詢值欄位
                        const selectedCols = await CSVManager.showColumnSelectDialog(result.headers, 'query', false);
                        if (selectedCols && selectedCols.length > 0) {
                            const colIndex = selectedCols[0];
                            const queryValues = result.data.map(row => row[result.headers[colIndex]]).filter(Boolean);
                            queryValueInput.value = queryValues.join('\n');
                            StateManager.csvImport.selectedColForQueryName = result.headers[colIndex];
                        }
                    } else if (purpose === 'a17merge') {
                        // 選擇A17合併欄位
                        const selectedCols = await CSVManager.showColumnSelectDialog(result.headers, 'a17merge', true);
                        if (selectedCols && selectedCols.length > 0) {
                            StateManager.csvImport.selectedColsForA17Merge = selectedCols.map(i => result.headers[i]);
                            StateManager.csvImport.isA17CsvPrepared = true;
                            displaySystemNotification(`已選擇${selectedCols.length}個欄位供A17合併`, false);
                        }
                    }
                } catch (error) {
                    displaySystemNotification(`CSV匯入失敗：${error.message}`, true);
                }
            };

            // 清除輸入
            overlay.querySelector('#qt-clear-input').onclick = () => {
                queryValueInput.value = '';
                csvFilename.textContent = '';
                csvFileInput.value = '';
                StateManager.csvImport = {
                    fileName: '',
                    rawHeaders: [],
                    rawData: [],
                    selectedColForQueryName: null,
                    selectedColsForA17Merge: [],
                    isA17CsvPrepared: false,
                };
            };

            overlay.querySelector('#qt-query-ok').onclick = () => {
                const queryValues = queryValueInput.value.trim();
                if (!queryValues) {
                    displaySystemNotification('請輸入查詢值', true);
                    return;
                }
                overlay.remove();
                resolve({
                    selectedApiKey,
                    queryValues
                });
            };

            overlay.querySelector('#qt-query-cancel').onclick = () => {
                overlay.remove();
                resolve(null);
            };
        });
    }

    // === CSV用途選擇對話框 ===
    function showCSVPurposeDialog() {
        return new Promise(resolve => {
            const contentHtml = `
            <h3 class="qt-dialog-title">CSV匯入用途</h3>
            <div style="margin-bottom:15px;font-size:14px;color:#555;">請選擇CSV檔案的用途：</div>
            <div style="display:flex;flex-direction:column;gap:10px;">
                <button id="qt-csv-purpose-query" class="qt-dialog-btn qt-dialog-btn-blue" style="width:100%;text-align:left;">
                    📋 將CSV某欄作為查詢值
                </button>
                <button id="qt-csv-purpose-a17" class="qt-dialog-btn qt-dialog-btn-green" style="width:100%;text-align:left;">
                    📊 勾選CSV欄位供A17合併顯示
                </button>
            </div>
            <div style="text-align:center;margin-top:15px;">
                <button id="qt-csv-purpose-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button>
            </div>
            `;

            const {
                overlay
            } = createDialogBase('_CSVPurpose', contentHtml, '350px', 'auto');

            const closeDialog = (value) => {
                overlay.remove();
                resolve(value);
            };

            overlay.querySelector('#qt-csv-purpose-query').onclick = () => closeDialog('query');
            overlay.querySelector('#qt-csv-purpose-a17').onclick = () => closeDialog('a17merge');
            overlay.querySelector('#qt-csv-purpose-cancel').onclick = () => closeDialog(null);
        });
    }

    // === 單筆重試對話框 ===
    async function showRetryEditDialog(row) {
        return new Promise(resolve => {
            const queryButtonsHtml = QUERYABLE_FIELD_DEFINITIONS.map(def =>
                `<button class="qt-querytype-btn" data-apikey="${def.queryApiKey}" style="background-color:${def.color};color:white;">${escapeHtml(def.queryDisplayName)}</button>`
            ).join('');

            const contentHtml = `
            <h3 class="qt-dialog-title">調整單筆查詢條件</h3>
            <div style="margin-bottom:10px;font-size:13px;color:#555;">選擇查詢欄位類型：</div>
            <div id="qt-retry-querytype-buttons" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:15px;">
                ${queryButtonsHtml}
            </div>
            <div style="margin-bottom:5px;font-size:13px;color:#555;">輸入查詢值：</div>
            <input type="text" id="qt-retry-queryvalue-input" class="qt-input" placeholder="請輸入查詢值..." value="${escapeHtml(row[FIELD_DISPLAY_NAMES_MAP._queriedValue_]||'')}">
            <div class="qt-dialog-flex-end">
                <button id="qt-retry-edit-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button>
                <button id="qt-retry-edit-ok" class="qt-dialog-btn qt-dialog-btn-blue">查詢</button>
            </div>
            `;

            const {
                overlay
            } = createDialogBase('_RetryEdit', contentHtml, '400px', 'auto');

            let selectedApiKey = row._originalQueryApiKey || QUERYABLE_FIELD_DEFINITIONS[0].queryApiKey;
            const queryTypeButtons = overlay.querySelectorAll('.qt-querytype-btn');

            function setActiveButton(apiKey) {
                queryTypeButtons.forEach(btn => {
                    const isSelected = btn.dataset.apikey === apiKey;
                    btn.style.border = isSelected ? `2px solid ${btn.style.backgroundColor}` : '2px solid transparent';
                    btn.style.boxShadow = isSelected ? `0 0 6px ${btn.style.backgroundColor}80` : 'none';
                });
            }

            queryTypeButtons.forEach(btn => {
                btn.onclick = () => {
                    selectedApiKey = btn.dataset.apikey;
                    setActiveButton(selectedApiKey);
                };
            });

            setActiveButton(selectedApiKey);

            overlay.querySelector('#qt-retry-edit-ok').onclick = () => {
                const queryValue = overlay.querySelector('#qt-retry-queryvalue-input').value.trim();
                if (!queryValue) {
                    displaySystemNotification('請輸入查詢值', true);
                    return;
                }
                overlay.remove();
                resolve({
                    apiKey: selectedApiKey,
                    queryValue
                });
            };

            overlay.querySelector('#qt-retry-edit-cancel').onclick = () => {
                overlay.remove();
                resolve(null);
            };
        });
    }

    // === 表格渲染函數 ===
    function populateTableRows(data) {
        if (!StateManager.currentTable.tableBodyElement || !StateManager.currentTable.tableHeadElement) {
            return;
        }

        const tbody = StateManager.currentTable.tableBodyElement;
        const thead = StateManager.currentTable.tableHeadElement;

        // 清空表格
        tbody.innerHTML = '';

        if (!data || data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="100%" style="text-align:center;color:#666;padding:20px;">暫無資料</td></tr>';
            return;
        }

        // 渲染表格行
        data.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');

            StateManager.currentTable.currentHeaders.forEach(headerKey => {
                const td = document.createElement('td');
                let cellValue = row[headerKey];

                if (cellValue === null || cellValue === undefined) {
                    cellValue = '';
                }

                // 特殊處理查詢結果欄位
                if (headerKey === FIELD_DISPLAY_NAMES_MAP._apiQueryStatus) {
                    td.innerHTML = escapeHtml(String(cellValue));

                    // 檢查是否需要添加重試按鈕
                    if (String(cellValue).includes('查詢失敗') || String(cellValue).includes('查無資料') || String(cellValue).includes('TOKEN失效') || String(cellValue).includes('網路錯誤')) {
                        const retryBtn = document.createElement('button');
                        retryBtn.className = 'qt-retry-btn';
                        retryBtn.textContent = '重新撈取';
                        retryBtn.dataset.rowIndex = rowIndex;
                        td.appendChild(document.createElement('br'));
                        td.appendChild(retryBtn);

                        if (row._retryFailed) {
                            const editBtn = document.createElement('button');
                            editBtn.className = 'qt-retry-edit-btn';
                            editBtn.textContent = '調整查詢條件';
                            editBtn.dataset.rowIndex = rowIndex;
                            td.appendChild(editBtn);
                        }
                    }
                } else if (headerKey === FIELD_DISPLAY_NAMES_MAP.statusCombined) {
                    // 狀態欄位特殊處理
                    td.innerHTML = String(cellValue);
                } else if (isEditMode && headerKey !== FIELD_DISPLAY_NAMES_MAP.NO && headerKey !== FIELD_DISPLAY_NAMES_MAP._queriedValue_) {
                    // 編輯模式下的可編輯儲存格
                    td.className = 'qt-editable-cell';
                    td.innerHTML = escapeHtml(String(cellValue));
                    td.onclick = () => editCell(td, row, headerKey);
                } else {
                    td.innerHTML = escapeHtml(String(cellValue));
                }

                tr.appendChild(td);
            });

            // 編輯模式下添加刪除按鈕
            if (isEditMode) {
                const deleteTd = document.createElement('td');
                const deleteBtn = document.createElement('button');
                deleteBtn.className = 'qt-delete-btn';
                deleteBtn.innerHTML = '🗑️';
                deleteBtn.onclick = () => {
                    if (confirm('確定要刪除這一列嗎？')) {
                        StateManager.pushSnapshot('刪除列');
                        data.splice(rowIndex, 1);
                        populateTableRows(data);
                        displaySystemNotification('已刪除列', false);
                    }
                };
                deleteTd.appendChild(deleteBtn);
                tr.appendChild(deleteTd);
            }

            tbody.appendChild(tr);
        });

        // 綁定重試按鈕事件
        tbody.querySelectorAll('.qt-retry-btn').forEach(btn => {
            btn.onclick = async (e) => {
                e.stopPropagation();
                const rowIndex = Number(btn.dataset.rowIndex);
                const row = data[rowIndex];
                const apiKey = row._originalQueryApiKey || selectedQueryDefinitionGlobal.queryApiKey;
                const queryValue = row._originalQueryValue || row[FIELD_DISPLAY_NAMES_MAP._queriedValue_];

                btn.disabled = true;
                btn.textContent = '查詢中...';

                const apiResult = await performApiQuery(queryValue, apiKey);

                if (apiResult.error === 'token_invalid') {
                    row._retryFailed = true;
                    row[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = '❌ TOKEN失效';
                    displaySystemNotification('TOKEN失效，請重新設定', true);
                } else if (apiResult.success) {
                    // 成功，更新該列資料
                    Object.assign(row, apiResult.data.records[0]);
                    ALL_DISPLAY_FIELDS_API_KEYS_MAIN.forEach(dKey => {
                        const displayName = FIELD_DISPLAY_NAMES_MAP[dKey] || dKey;
                        let cellValue = apiResult.data.records[0][dKey];

                        if (cellValue === null || cellValue === undefined) {
                            cellValue = '';
                        } else {
                            cellValue = String(cellValue);
                        }

                        if (dKey === 'statusCombined') {
                            const mainS = apiResult.data.records[0].mainStatus || '';
                            const subS = apiResult.data.records[0].subStatus || '';
                            row[displayName] = `<span style="font-weight:bold;">${escapeHtml(mainS)}</span>` + (subS ? ` <span style="color:#777;">(${escapeHtml(subS)})</span>` : '');
                        } else if (dKey === UNIT_MAP_FIELD_API_KEY) {
                            const unitCodePrefix = getFirstLetter(cellValue);
                            const mappedUnitName = UNIT_CODE_MAPPINGS[unitCodePrefix] || cellValue;
                            row[displayName] = unitCodePrefix && UNIT_CODE_MAPPINGS[unitCodePrefix] ? `${unitCodePrefix}-${mappedUnitName.replace(/^[A-Z]-/,'')}` : mappedUnitName;
                        } else if (dKey === 'uwApprover' || dKey === 'approvalUser') {
                            row[displayName] = extractName(cellValue);
                        } else {
                            row[displayName] = cellValue;
                        }
                    });
                    row[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = '✔️ 成功';
                    row._retryFailed = false;
                    displaySystemNotification('單筆重查成功', false);
                } else {
                    row._retryFailed = true;
                    row[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = '❌ 查詢失敗';
                    displaySystemNotification('單筆重查失敗，請調整條件', true);
                }

                populateTableRows(data);
            };
        });

        // 綁定調整條件按鈕事件
        tbody.querySelectorAll('.qt-retry-edit-btn').forEach(btn => {
            btn.onclick = async (e) => {
                e.stopPropagation();
                const rowIndex = Number(btn.dataset.rowIndex);
                const row = data[rowIndex];

                const result = await showRetryEditDialog(row);
                if (result) {
                    btn.disabled = true;
                    btn.textContent = '查詢中...';

                    const apiResult = await performApiQuery(result.queryValue, result.apiKey);

                    if (apiResult.error === 'token_invalid') {
                        row._retryFailed = true;
                        row[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = '❌ TOKEN失效';
                        displaySystemNotification('TOKEN失效，請重新設定', true);
                    } else if (apiResult.success) {
                        // 成功，更新該列資料
                        Object.assign(row, apiResult.data.records[0]);
                        ALL_DISPLAY_FIELDS_API_KEYS_MAIN.forEach(dKey => {
                            const displayName = FIELD_DISPLAY_NAMES_MAP[dKey] || dKey;
                            let cellValue = apiResult.data.records[0][dKey];

                            if (cellValue === null || cellValue === undefined) {
                                cellValue = '';
                            } else {
                                cellValue = String(cellValue);
                            }

                            if (dKey === 'statusCombined') {
                                const mainS = apiResult.data.records[0].mainStatus || '';
                                const subS = apiResult.data.records[0].subStatus || '';
                                row[displayName] = `<span style="font-weight:bold;">${escapeHtml(mainS)}</span>` + (subS ? ` <span style="color:#777;">(${escapeHtml(subS)})</span>` : '');
                            } else if (dKey === UNIT_MAP_FIELD_API_KEY) {
                                const unitCodePrefix = getFirstLetter(cellValue);
                                const mappedUnitName = UNIT_CODE_MAPPINGS[unitCodePrefix] || cellValue;
                                row[displayName] = unitCodePrefix && UNIT_CODE_MAPPINGS[unitCodePrefix] ? `${unitCodePrefix}-${mappedUnitName.replace(/^[A-Z]-/,'')}` : mappedUnitName;
                            } else if (dKey === 'uwApprover' || dKey === 'approvalUser') {
                                row[displayName] = extractName(cellValue);
                            } else {
                                row[displayName] = cellValue;
                            }
                        });
                        row[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = '✔️ 成功';
                        row._retryFailed = false;
                        row._originalQueryApiKey = result.apiKey;
                        row._originalQueryValue = result.queryValue;
                        displaySystemNotification('調整條件後查詢成功', false);
                    } else {
                        row[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = '❌ 查詢失敗';
                        displaySystemNotification('調整條件後查詢失敗', true);
                    }

                    populateTableRows(data);
                }
            };
        });
    }

    // === 儲存格編輯函數 ===
    function editCell(td, row, headerKey) {
        if (td.querySelector('input') || td.querySelector('select')) return;

        const originalValue = row[headerKey] || '';
        const originalHtml = td.innerHTML;

        if (headerKey === FIELD_DISPLAY_NAMES_MAP.uwApproverUnit) {
            // 單位欄位使用下拉選單
            const select = document.createElement('select');
            select.className = 'qt-select';
            select.style.width = '100%';
            select.style.margin = '0';

            // 添加選項
            Object.entries(UNIT_CODE_MAPPINGS).forEach(([code, name]) => {
                const option = document.createElement('option');
                option.value = `${code}-${name}`;
                option.textContent = `${code}-${name}`;
                if (originalValue.includes(code)) option.selected = true;
                select.appendChild(option);
            });

            td.innerHTML = '';
            td.appendChild(select);
            select.focus();

            const saveEdit = () => {
                const newValue = select.value;
                row[headerKey] = newValue;
                td.innerHTML = escapeHtml(newValue);
                displaySystemNotification('已更新儲存格', false);
            };

            const cancelEdit = () => {
                td.innerHTML = originalHtml;
            };

            select.onblur = saveEdit;
            select.onchange = saveEdit;
            select.onkeydown = e => {
                if (e.key === 'Enter') saveEdit();
                if (e.key === 'Escape') cancelEdit();
            };
        } else {
            // 一般文字欄位使用輸入框
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'qt-input';
            input.style.width = '100%';
            input.style.margin = '0';
            input.value = originalValue;

            td.innerHTML = '';
            td.appendChild(input);
            input.focus();
            input.select();

            const saveEdit = () => {
                const newValue = input.value.trim();
                row[headerKey] = newValue;
                td.innerHTML = escapeHtml(newValue);
                displaySystemNotification('已更新儲存格', false);
            };

            const cancelEdit = () => {
                td.innerHTML = originalHtml;
            };

            input.onblur = saveEdit;
            input.onkeydown = e => {
                if (e.key === 'Enter') saveEdit();
                if (e.key === 'Escape') cancelEdit();
            };
        }
    }

    // === 主結果表格UI ===
    function renderResultsTableUI(data) {
        // 移除現有主視窗
        document.getElementById(TOOL_MAIN_CONTAINER_ID)?.remove();

        const mainContainer = document.createElement('div');
        mainContainer.id = TOOL_MAIN_CONTAINER_ID;
        mainContainer.style.cssText = `
        position:fixed;
        top:50%;
        left:50%;
        transform:translate(-50%,-50%);
        width:90vw;
        max-width:1200px;
        height:80vh;
        background:#fff;
        border-radius:8px;
        box-shadow:0 10px 30px rgba(0,0,0,0.3);
        z-index:${Z_INDEX.MAIN_UI};
        font-family:'Microsoft JhengHei',Arial,sans-serif;
        display:flex;
        flex-direction:column;
        overflow:hidden;
        `;

        // 標題列
        const titleBar = document.createElement('div');
        titleBar.style.cssText = 'background:#f8f9fa;padding:15px 20px;border-bottom:1px solid #dee2e6;cursor:move;user-select:none;font-weight:600;font-size:16px;color:#333;';
        titleBar.textContent = '凱基人壽案件查詢結果';

        // 拖曳功能
        let isDragging = false;
        let dragOffset = {
            x: 0,
            y: 0
        };

        titleBar.onmousedown = e => {
            isDragging = true;
            dragOffset.x = e.clientX - mainContainer.offsetLeft;
            dragOffset.y = e.clientY - mainContainer.offsetTop;
            document.addEventListener('mousemove', onDrag);
            document.addEventListener('mouseup', onDragEnd);
        };

        const onDrag = e => {
            if (!isDragging) return;
            const newX = e.clientX - dragOffset.x;
            const newY = e.clientY - dragOffset.y;
            mainContainer.style.left = newX + 'px';
            mainContainer.style.top = newY + 'px';
            mainContainer.style.transform = 'none';
        };

        const onDragEnd = () => {
            isDragging = false;
            document.removeEventListener('mousemove', onDrag);
            document.removeEventListener('mouseup', onDragEnd);
        };

        // 控制列
        const controlBar = document.createElement('div');
        controlBar.style.cssText = 'padding:10px 20px;border-bottom:1px solid #dee2e6;display:flex;justify-content:space-between;align-items:center;background:#fff;';

        const leftControls = document.createElement('div');
        leftControls.style.display = 'flex';
        leftControls.style.gap = '8px';

        const rightControls = document.createElement('div');
        rightControls.style.display = 'flex';
        rightControls.style.gap = '8px';
        rightControls.style.alignItems = 'center';

        // 左側按鈕
        const clearBtn = createControlButton('清除條件', '#6c757d');
        const requeryBtn = createControlButton('重新查詢', '#fd7e14');
        const a17Btn = createControlButton('A17作業', '#6f42c1');
        const copyBtn = createControlButton('複製表格', '#28a745');
        const editBtn = createControlButton(isEditMode ? '結束編輯' : '編輯模式', isEditMode ? '#dc3545' : '#007bff');

        leftControls.appendChild(clearBtn);
        leftControls.appendChild(requeryBtn);
        leftControls.appendChild(a17Btn);
        leftControls.appendChild(copyBtn);
        leftControls.appendChild(editBtn);

        // 編輯模式下的新增按鈕
        if (isEditMode) {
            const addBtn = createControlButton('+ 新增列', '#007bff');
            addBtn.onclick = () => {
                StateManager.pushSnapshot('新增列');
                const newRow = {};
                StateManager.currentTable.currentHeaders.forEach(header => {
                    newRow[header] = header === FIELD_DISPLAY_NAMES_MAP.NO ? String(data.length + 1) : '';
                });
                data.push(newRow);
                populateTableRows(data);
                displaySystemNotification('已新增列', false);
            };
            leftControls.appendChild(addBtn);
        }

        // 右側控制
        const searchInput = document.createElement('input');
        searchInput.type = 'text';
        searchInput.placeholder = '篩選表格內容...';
        searchInput.style.cssText = 'padding:6px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;width:200px;';

        const closeBtn = createControlButton('關閉工具', '#dc3545');

        rightControls.appendChild(searchInput);
        rightControls.appendChild(closeBtn);

        controlBar.appendChild(leftControls);
        controlBar.appendChild(rightControls);

        // A17模式控制區
        const a17Controls = document.createElement('div');
        a17Controls.style.cssText = 'padding:10px 20px;border-bottom:1px solid #dee2e6;background:#f8f9fa;display:none;';

        if (StateManager.a17Mode.isActive) {
            a17Controls.style.display = 'block';
            renderA17Controls(a17Controls);
        }

        // 表格容器
        const tableContainer = document.createElement('div');
        tableContainer.style.cssText = 'flex:1;overflow:auto;padding:0;';

        const table = document.createElement('table');
        table.className = 'qt-table';
        table.style.cssText = 'width:100%;border-collapse:collapse;margin:0;';

        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        table.appendChild(thead);
        table.appendChild(tbody);
        tableContainer.appendChild(table);

        // 組裝主視窗
        mainContainer.appendChild(titleBar);
        mainContainer.appendChild(controlBar);
        mainContainer.appendChild(a17Controls);
        mainContainer.appendChild(tableContainer);

        document.body.appendChild(mainContainer);

        // 儲存引用
        StateManager.currentTable.mainUIElement = mainContainer;
        StateManager.currentTable.tableBodyElement = tbody;
        StateManager.currentTable.tableHeadElement = thead;
        StateManager.currentTable.a17UnitButtonsContainer = a17Controls;

        // 設定表頭
        setupTableHeaders(data);

        // 填充資料
        populateTableRows(data);

        // 綁定事件
        bindControlEvents(clearBtn, requeryBtn, a17Btn, copyBtn, editBtn, closeBtn, searchInput);

        // 搜尋功能
        searchInput.oninput = () => {
            const searchTerm = searchInput.value.toLowerCase();
            const rows = tbody.querySelectorAll('tr');
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(searchTerm) ? '' : 'none';
            });
        };
    }

    function createControlButton(text, color) {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.style.cssText = `
        background:${color};
        color:white;
        border:none;
        padding:6px 12px;
        border-radius:4px;
        font-size:13px;
        cursor:pointer;
        transition:opacity 0.2s ease;
        font-weight:500;
        `;
        btn.onmouseover = () => btn.style.opacity = '0.85';
        btn.onmouseout = () => btn.style.opacity = '1';
        return btn;
    }

    // === 補足缺失的函數 ===
    function setupTableHeaders(data) {
        const thead = StateManager.currentTable.tableHeadElement;
        thead.innerHTML = '';

        if (!data || data.length === 0) {
            StateManager.currentTable.currentHeaders = [];
            return;
        }

        // 確定表頭欄位
        const baseHeaders = [
            FIELD_DISPLAY_NAMES_MAP._queriedValue_,
            FIELD_DISPLAY_NAMES_MAP.NO,
            ...ALL_DISPLAY_FIELDS_API_KEYS_MAIN.map(key => FIELD_DISPLAY_NAMES_MAP[key]),
            FIELD_DISPLAY_NAMES_MAP._apiQueryStatus
        ];

        // A17模式下加入CSV合併欄位
        if (StateManager.a17Mode.isActive && StateManager.csvImport.selectedColsForA17Merge.length > 0) {
            baseHeaders.splice(-1, 0, ...StateManager.csvImport.selectedColsForA17Merge);
        }

        // 編輯模式下加入操作欄
        if (isEditMode) {
            baseHeaders.push('操作');
        }

        StateManager.currentTable.currentHeaders = baseHeaders;

        const headerRow = document.createElement('tr');
        baseHeaders.forEach(headerKey => {
            const th = document.createElement('th');
            th.textContent = headerKey;
            th.style.position = 'relative';

            // 排序功能
            if (headerKey !== '操作') {
                th.style.cursor = 'pointer';
                th.onclick = () => {
                    const currentDirection = StateManager.currentTable.sortDirections[headerKey] || 'asc';
                    const newDirection = currentDirection === 'asc' ? 'desc' : 'asc';
                    StateManager.currentTable.sortDirections[headerKey] = newDirection;

                    // 清除其他欄位的排序方向
                    Object.keys(StateManager.currentTable.sortDirections).forEach(key => {
                        if (key !== headerKey) delete StateManager.currentTable.sortDirections[key];
                    });

                    // 排序資料
                    const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
                    currentData.sort((a, b) => {
                        let aVal = a[headerKey] || '';
                        let bVal = b[headerKey] || '';

                        // 數字欄位特殊處理
                        if (headerKey === FIELD_DISPLAY_NAMES_MAP.NO) {
                            aVal = parseInt(aVal) || 0;
                            bVal = parseInt(bVal) || 0;
                        } else {
                            aVal = String(aVal).toLowerCase();
                            bVal = String(bVal).toLowerCase();
                        }

                        if (newDirection === 'asc') {
                            return aVal > bVal ? 1 : (aVal < bVal ? -1 : 0);
                        } else {
                            return aVal < bVal ? 1 : (aVal > bVal ? -1 : 0);
                        }
                    });

                    populateTableRows(currentData);

                    // 更新排序箭頭
                    thead.querySelectorAll('th').forEach(thEl => {
                        thEl.querySelector('.sort-arrow')?.remove();
                    });

                    const arrow = document.createElement('span');
                    arrow.className = 'sort-arrow';
                    arrow.innerHTML = newDirection === 'asc' ? ' ↑' : ' ↓';
                    arrow.style.cssText = 'color:#007bff;font-weight:bold;';
                    th.appendChild(arrow);
                };
            }

            headerRow.appendChild(th);
        });

        thead.appendChild(headerRow);
    }

    // === A17控制區渲染 ===
    function renderA17Controls(container) {
        container.innerHTML = '';

        // 單位篩選按鈕群
        const unitButtonsContainer = document.createElement('div');
        unitButtonsContainer.style.cssText = 'display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px;';

        A17_UNIT_BUTTONS_DEFS.forEach(unitDef => {
            const btn = document.createElement('button');
            btn.className = 'qt-unit-filter-btn';
            btn.dataset.unitId = unitDef.id;
            btn.style.backgroundColor = unitDef.color;
            btn.style.color = 'white';

            // 計算該單位的資料筆數
            const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
            let count = 0;
            if (unitDef.id === 'UNDEF') {
                count = currentData.filter(row => {
                    const unit = row[FIELD_DISPLAY_NAMES_MAP.uwApproverUnit] || '';
                    return !Object.keys(UNIT_CODE_MAPPINGS).some(code => unit.includes(code));
                }).length;
            } else {
                count = currentData.filter(row => {
                    const unit = row[FIELD_DISPLAY_NAMES_MAP.uwApproverUnit] || '';
                    return unit.includes(unitDef.id);
                }).length;
            }

            btn.textContent = `${unitDef.label} (${count})`;

            // 無資料時禁用
            if (count === 0) {
                btn.disabled = true;
                btn.style.opacity = '0.5';
                btn.style.cursor = 'not-allowed';
            } else {
                // 檢查是否已選中
                if (StateManager.a17Mode.selectedUnits.has(unitDef.id)) {
                    btn.classList.add('active');
                }

                btn.onclick = () => {
                    if (StateManager.a17Mode.selectedUnits.has(unitDef.id)) {
                        StateManager.a17Mode.selectedUnits.delete(unitDef.id);
                        btn.classList.remove('active');
                    } else {
                        StateManager.a17Mode.selectedUnits.add(unitDef.id);
                        btn.classList.add('active');
                    }

                    // 篩選表格資料
                    filterA17TableByUnits();
                };
            }

            unitButtonsContainer.appendChild(btn);
        });

        // A17通知文控制
        const textControlsContainer = document.createElement('div');
        textControlsContainer.style.cssText = 'display:flex;align-items:center;gap:10px;';

        const includeTextCheckbox = document.createElement('input');
        includeTextCheckbox.type = 'checkbox';
        includeTextCheckbox.id = 'qt-a17-include-text';
        includeTextCheckbox.checked = true;

        const includeTextLabel = document.createElement('label');
        includeTextLabel.htmlFor = 'qt-a17-include-text';
        includeTextLabel.textContent = 'A17含通知文';
        includeTextLabel.style.cursor = 'pointer';

        const editTextBtn = createControlButton('編輯通知文', '#007bff');
        editTextBtn.onclick = () => showA17TextSettingsDialog();

        textControlsContainer.appendChild(includeTextCheckbox);
        textControlsContainer.appendChild(includeTextLabel);
        textControlsContainer.appendChild(editTextBtn);

        container.appendChild(unitButtonsContainer);
        container.appendChild(textControlsContainer);
    }

    // === A17單位篩選 ===
    function filterA17TableByUnits() {
        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;

        if (StateManager.a17Mode.selectedUnits.size === 0) {
            populateTableRows(currentData);
            return;
        }

        const filteredData = currentData.filter(row => {
            const unit = row[FIELD_DISPLAY_NAMES_MAP.uwApproverUnit] || '';

            return Array.from(StateManager.a17Mode.selectedUnits).some(selectedUnit => {
                if (selectedUnit === 'UNDEF') {
                    return !Object.keys(UNIT_CODE_MAPPINGS).some(code => unit.includes(code));
                } else {
                    return unit.includes(selectedUnit);
                }
            });
        });

        populateTableRows(filteredData);
    }

    // === A17通知文設定對話框 ===
    function showA17TextSettingsDialog() {
        return new Promise(resolve => {
            const settings = StateManager.a17Mode.textSettings;

            const contentHtml = `
        <h3 class="qt-dialog-title">A17 通知文設定</h3>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;">
            <div>
                <label style="display:block;margin-bottom:5px;font-weight:500;">主文案內容：</label>
                <textarea id="qt-a17-main-content" class="qt-textarea" style="height:120px;">${escapeHtml(settings.mainContent)}</textarea>
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">字體大小 (pt)：</label>
                <input type="number" id="qt-a17-main-font-size" class="qt-input" min="8" max="24" value="${settings.mainFontSize}">
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">行高：</label>
                <input type="number" id="qt-a17-main-line-height" class="qt-input" min="1" max="3" step="0.1" value="${settings.mainLineHeight}">
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">字體顏色：</label>
                <input type="color" id="qt-a17-main-color" class="qt-input" value="${settings.mainFontColor}">
            </div>
            <div>
                <label style="display:block;margin-bottom:5px;font-weight:500;">產檔時間偏移 (天)：</label>
                <input type="number" id="qt-a17-gen-offset" class="qt-input" value="${settings.genDateOffset}">
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">對比時間偏移 (天)：</label>
                <input type="number" id="qt-a17-comp-offset" class="qt-input" value="${settings.compDateOffset}">
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">日期字體大小 (pt)：</label>
                <input type="number" id="qt-a17-date-font-size" class="qt-input" min="6" max="16" value="${settings.dateFontSize}">
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">日期行高：</label>
                <input type="number" id="qt-a17-date-line-height" class="qt-input" min="1" max="3" step="0.1" value="${settings.dateLineHeight}">
                
                <label style="display:block;margin-bottom:5px;font-weight:500;">日期顏色：</label>
                <input type="color" id="qt-a17-date-color" class="qt-input" value="${settings.dateFontColor}">
            </div>
        </div>
        
        <div style="margin-top:15px;">
            <label style="display:block;margin-bottom:5px;font-weight:500;">預覽區（可臨時編輯）：</label>
            <div id="qt-a17-preview" contenteditable="true" style="border:1px solid #ddd;padding:10px;border-radius:4px;min-height:100px;background:#f9f9f9;font-family:'Microsoft JhengHei',Arial,sans-serif;"></div>
        </div>
        
        <div class="qt-dialog-flex-between">
            <button id="qt-a17-reset" class="qt-dialog-btn qt-dialog-btn-orange">重設預設</button>
            <div>
                <button id="qt-a17-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button>
                <button id="qt-a17-save" class="qt-dialog-btn qt-dialog-btn-blue">儲存設定</button>
            </div>
        </div>
        `;

            const {
                overlay
            } = createDialogBase('_A17TextSettings', contentHtml, '700px', 'auto');

            const previewEl = overlay.querySelector('#qt-a17-preview');

            function updatePreview() {
                const mainContent = overlay.querySelector('#qt-a17-main-content').value;
                const mainFontSize = overlay.querySelector('#qt-a17-main-font-size').value;
                const mainLineHeight = overlay.querySelector('#qt-a17-main-line-height').value;
                const mainColor = overlay.querySelector('#qt-a17-main-color').value;
                const genOffset = parseInt(overlay.querySelector('#qt-a17-gen-offset').value) || 0;
                const compOffset = parseInt(overlay.querySelector('#qt-a17-comp-offset').value) || 0;
                const dateFontSize = overlay.querySelector('#qt-a17-date-font-size').value;
                const dateLineHeight = overlay.querySelector('#qt-a17-date-line-height').value;
                const dateColor = overlay.querySelector('#qt-a17-date-color').value;

                const genDate = new Date();
                genDate.setDate(genDate.getDate() + genOffset);
                const compDate = new Date();
                compDate.setDate(compDate.getDate() + compOffset);

                previewEl.innerHTML = `
            <div style="font-size:${mainFontSize}px;line-height:${mainLineHeight};color:${mainColor};">
                ${escapeHtml(mainContent).replace(/\n/g,'<br>')}
            </div>
            <div style="margin-top:15px;font-size:${dateFontSize}px;line-height:${dateLineHeight};color:${dateColor};">
                產檔時間：${formatDate(genDate)}<br>
                對比時間：${formatDate(compDate)}
            </div>
            `;
            }

            // 初始預覽
            updatePreview();

            // 綁定輸入事件
            ['#qt-a17-main-content', '#qt-a17-main-font-size', '#qt-a17-main-line-height', '#qt-a17-main-color',
                '#qt-a17-gen-offset', '#qt-a17-comp-offset', '#qt-a17-date-font-size', '#qt-a17-date-line-height', '#qt-a17-date-color'
            ]
            .forEach(selector => {
                overlay.querySelector(selector).oninput = updatePreview;
            });

            overlay.querySelector('#qt-a17-reset').onclick = () => {
                overlay.querySelector('#qt-a17-main-content').value = A17_DEFAULT_TEXT_CONTENT;
                overlay.querySelector('#qt-a17-main-font-size').value = '12';
                overlay.querySelector('#qt-a17-main-line-height').value = '1.5';
                overlay.querySelector('#qt-a17-main-color').value = '#333333';
                overlay.querySelector('#qt-a17-gen-offset').value = '-3';
                overlay.querySelector('#qt-a17-comp-offset').value = '0';
                overlay.querySelector('#qt-a17-date-font-size').value = '8';
                overlay.querySelector('#qt-a17-date-line-height').value = '1.2';
                overlay.querySelector('#qt-a17-date-color').value = '#555555';
                updatePreview();
            };

            overlay.querySelector('#qt-a17-save').onclick = () => {
                StateManager.a17Mode.textSettings = {
                    mainContent: overlay.querySelector('#qt-a17-main-content').value,
                    mainFontSize: parseInt(overlay.querySelector('#qt-a17-main-font-size').value),
                    mainLineHeight: parseFloat(overlay.querySelector('#qt-a17-main-line-height').value),
                    mainFontColor: overlay.querySelector('#qt-a17-main-color').value,
                    dateFontSize: parseInt(overlay.querySelector('#qt-a17-date-font-size').value),
                    dateLineHeight: parseFloat(overlay.querySelector('#qt-a17-date-line-height').value),
                    dateFontColor: overlay.querySelector('#qt-a17-date-color').value,
                    genDateOffset: parseInt(overlay.querySelector('#qt-a17-gen-offset').value),
                    compDateOffset: parseInt(overlay.querySelector('#qt-a17-comp-offset').value),
                };
                StateManager.saveA17Settings();
                overlay.remove();
                displaySystemNotification('A17通知文設定已儲存', false);
                resolve(true);
            };

            overlay.querySelector('#qt-a17-cancel').onclick = () => {
                overlay.remove();
                resolve(false);
            };
        });
    }

    // === 控制事件綁定 ===
    function bindControlEvents(clearBtn, requeryBtn, a17Btn, copyBtn, editBtn, closeBtn, searchInput) {
        clearBtn.onclick = () => {
            if (confirm('確定要清除所有查詢條件和結果嗎？')) {
                StateManager.originalQueryResults = [];
                StateManager.baseA17MasterData = [];
                StateManager.csvImport = {
                    fileName: '',
                    rawHeaders: [],
                    rawData: [],
                    selectedColForQueryName: null,
                    selectedColsForA17Merge: [],
                    isA17CsvPrepared: false,
                };
                StateManager.a17Mode.isActive = false;
                StateManager.a17Mode.selectedUnits.clear();
                isEditMode = false;

                StateManager.currentTable.mainUIElement?.remove();
                displaySystemNotification('已清除所有條件', false);
            }
        };

        requeryBtn.onclick = async () => {
            const querySetupResult = await createQuerySetupDialog();
            if (querySetupResult) {
                await executeQuery(querySetupResult);
            }
        };

        a17Btn.onmousedown = () => {
            a17ButtonLongPressTimer = setTimeout(() => {
                // 長按強制進入A17模式
                StateManager.pushSnapshot('強制進入A17模式');
                StateManager.a17Mode.isActive = true;
                StateManager.a17Mode.selectedUnits.clear();

                // 重新渲染UI
                renderResultsTableUI(StateManager.originalQueryResults);
                displaySystemNotification('已強制進入A17模式', false);
            }, 1000);
        };

        a17Btn.onmouseup = () => {
            if (a17ButtonLongPressTimer) {
                clearTimeout(a17ButtonLongPressTimer);
                a17ButtonLongPressTimer = null;
            }
        };

        a17Btn.onclick = () => {
            if (a17ButtonLongPressTimer) {
                clearTimeout(a17ButtonLongPressTimer);
                a17ButtonLongPressTimer = null;
            }

            if (!StateManager.csvImport.isA17CsvPrepared) {
                displaySystemNotification('請先匯入CSV並選擇A17合併欄位', true);
                return;
            }

            StateManager.pushSnapshot('切換A17模式');
            StateManager.a17Mode.isActive = !StateManager.a17Mode.isActive;

            if (StateManager.a17Mode.isActive) {
                // 合併CSV資料到查詢結果
                mergeA17Data();
            }

            renderResultsTableUI(StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults);
            displaySystemNotification(`已${StateManager.a17Mode.isActive?'進入':'退出'}A17模式`, false);
        };

        copyBtn.onclick = () => {
            const includeText = document.querySelector('#qt-a17-include-text')?.checked || false;
            copyTableToClipboard(includeText);
        };

        editBtn.onclick = () => {
            StateManager.pushSnapshot('切換編輯模式');
            isEditMode = !isEditMode;
            editBtn.textContent = isEditMode ? '結束編輯' : '編輯模式';
            editBtn.style.backgroundColor = isEditMode ? '#dc3545' : '#007bff';

            const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
            renderResultsTableUI(currentData);
            displaySystemNotification(`已${isEditMode?'進入':'退出'}編輯模式`, false);
        };

        closeBtn.onclick = () => {
            if (confirm('確定要關閉查詢工具嗎？')) {
                StateManager.currentTable.mainUIElement?.remove();
                displaySystemNotification('查詢工具已關閉', false);
            }
        };

        // Ctrl+Z復原功能
        document.addEventListener('keydown', e => {
            if (e.ctrlKey && e.key === 'z' && StateManager.currentTable.mainUIElement) {
                e.preventDefault();
                StateManager.undo();
            }
        });
    }

    // === A17資料合併 ===
    function mergeA17Data() {
        if (!StateManager.csvImport.isA17CsvPrepared || StateManager.csvImport.selectedColsForA17Merge.length === 0) {
            StateManager.baseA17MasterData = [...StateManager.originalQueryResults];
            return;
        }

        StateManager.baseA17MasterData = StateManager.originalQueryResults.map(queryRow => {
            const mergedRow = {
                ...queryRow
            };

            // 尋找對應的CSV資料
            const queryValue = queryRow[FIELD_DISPLAY_NAMES_MAP._queriedValue_];
            const matchingCsvRow = StateManager.csvImport.rawData.find(csvRow => {
                return Object.values(csvRow).some(value => String(value).trim() === String(queryValue).trim());
            });

            if (matchingCsvRow) {
                // 合併選中的CSV欄位
                StateManager.csvImport.selectedColsForA17Merge.forEach(colName => {
                    mergedRow[colName] = matchingCsvRow[colName] || '';
                });
            } else {
                // 沒有匹配的CSV資料，填入空值
                StateManager.csvImport.selectedColsForA17Merge.forEach(colName => {
                    mergedRow[colName] = '';
                });
            }

            return mergedRow;
        });
    }

    // === 複製表格功能 ===
    function copyTableToClipboard(includeA17Text = false) {
        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
        const headers = StateManager.currentTable.currentHeaders;

        if (!currentData || currentData.length === 0) {
            displaySystemNotification('沒有資料可複製', true);
            return;
        }

        let content = '';

        // A17模式且包含通知文
        if (StateManager.a17Mode.isActive && includeA17Text) {
            const settings = StateManager.a17Mode.textSettings;
            const genDate = new Date();
            genDate.setDate(genDate.getDate() + settings.genDateOffset);
            const compDate = new Date();
            compDate.setDate(compDate.getDate() + settings.compDateOffset);

            content += `<div style="font-size:${settings.mainFontSize}px;line-height:${settings.mainLineHeight};color:${settings.mainFontColor};margin-bottom:20px;">`;
            content += settings.mainContent.replace(/\n/g, '<br>');
            content += `</div>`;
            content += `<div style="font-size:${settings.dateFontSize}px;line-height:${settings.dateLineHeight};color:${settings.dateFontColor};margin-bottom:20px;">`;
            content += `產檔時間：${formatDate(genDate)}<br>對比時間：${formatDate(compDate)}`;
            content += `</div>`;
        }

        // 表格HTML
        content += `<table border="1" style="border-collapse:collapse;width:100%;">`;
        content += `<thead><tr>`;
        headers.forEach(header => {
            content += `<th style="background:#f8f9fa;padding:8px;text-align:left;">${escapeHtml(header)}</th>`;
        });
        content += `</tr></thead><tbody>`;

        currentData.forEach(row => {
            content += `<tr>`;
            headers.forEach(header => {
                let cellValue = row[header];
                if (cellValue === null || cellValue === undefined) cellValue = '';
                content += `<td style="padding:8px;border:1px solid #ddd;">${String(cellValue)}</td>`;
            });
            content += `</tr>`;
        });

        content += `</tbody></table>`;

        // 複製到剪貼簿
        const blob = new Blob([content], {
            type: 'text/html'
        });
        const item = new ClipboardItem({
            'text/html': blob
        });

        navigator.clipboard.write([item]).then(() => {
            displaySystemNotification('表格已複製到剪貼簿', false);
        }).catch(() => {
            // 降級處理
            const textContent = headers.join('\t') + '\n' +
                currentData.map(row => headers.map(h => row[h] || '').join('\t')).join('\n');
            navigator.clipboard.writeText(textContent).then(() => {
                displaySystemNotification('表格已複製到剪貼簿（純文字格式）', false);
            }).catch(() => {
                displaySystemNotification('複製失敗，請手動選擇表格內容', true);
            });
        });
    }

    // === 主執行函數 ===
    async function executeQuery(querySetupResult) {
        const queryValues = querySetupResult.queryValues.split(/[\s,;\n]+/).map(x => x.trim().toUpperCase()).filter(Boolean);

        if (queryValues.length === 0) {
            displaySystemNotification('未輸入有效查詢值', true);
            return;
        }

        selectedQueryDefinitionGlobal = QUERYABLE_FIELD_DEFINITIONS.find(qdf => qdf.queryApiKey === querySetupResult.selectedApiKey);

        // 顯示載入對話框
        const loadingDialog = createDialogBase('_Loading', `
        <h3 class="qt-dialog-title" id="${TOOL_MAIN_CONTAINER_ID}_LoadingTitle">查詢中...</h3>
        <p id="${TOOL_MAIN_CONTAINER_ID}_LoadingMsg" style="text-align:center;font-size:13px;color:#555;">處理中...</p>
        <div style="width:40px;height:40px;border:4px solid #f3f3f3;border-top:4px solid #3498db;border-radius:50%;margin:15px auto;animation:qtSpin 1s linear infinite;"></div>
        <style>@keyframes qtSpin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}</style>
    `, '300px', 'auto', 'text-align:center;');

        const loadingTitleEl = loadingDialog.dialog.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_LoadingTitle`);
        const loadingMsgEl = loadingDialog.dialog.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_LoadingMsg`);

        StateManager.originalQueryResults = [];
        let currentQueryCount = 0;

        for (const singleQueryValue of queryValues) {
            currentQueryCount++;
            if (loadingTitleEl) loadingTitleEl.textContent = `查詢中 (${currentQueryCount}/${queryValues.length})`;
            if (loadingMsgEl) loadingMsgEl.textContent = `正在處理: ${singleQueryValue}`;

            const resultRowBase = {
                [FIELD_DISPLAY_NAMES_MAP.NO]: String(currentQueryCount),
                [FIELD_DISPLAY_NAMES_MAP._queriedValue_]: singleQueryValue,
                _originalQueryApiKey: selectedQueryDefinitionGlobal.queryApiKey,
                _originalQueryValue: singleQueryValue,
                _retryFailed: false
            };

            const apiResult = await performApiQuery(singleQueryValue, selectedQueryDefinitionGlobal.queryApiKey);
            let apiQueryStatusText = '❌ 查詢失敗';

            if (apiResult.error === 'token_invalid') {
                apiQueryStatusText = '❌ TOKEN失效';
            } else if (apiResult.success) {
                apiQueryStatusText = '✔️ 成功';
            } else if (!apiResult.error) {
                apiQueryStatusText = '➖ 查無資料';
            }

            resultRowBase[FIELD_DISPLAY_NAMES_MAP._apiQueryStatus] = apiQueryStatusText;

            if (apiResult.success && apiResult.data.records) {
                apiResult.data.records.forEach(rec => {
                    const populatedRow = {
                        ...resultRowBase
                    };
                    ALL_DISPLAY_FIELDS_API_KEYS_MAIN.forEach(dKey => {
                        const displayName = FIELD_DISPLAY_NAMES_MAP[dKey] || dKey;
                        let cellValue = rec[dKey] === null || rec[dKey] === undefined ? '' : String(rec[dKey]);

                        if (dKey === 'statusCombined') {
                            const mainS = rec.mainStatus || '';
                            const subS = rec.subStatus || '';
                            populatedRow[displayName] = `<span style="font-weight:bold;">${escapeHtml(mainS)}</span>` + (subS ? ` <span style="color:#777;">(${escapeHtml(subS)})</span>` : '');
                        } else if (dKey === UNIT_MAP_FIELD_API_KEY) {
                            const unitCodePrefix = getFirstLetter(cellValue);
                            const mappedUnitName = UNIT_CODE_MAPPINGS[unitCodePrefix] || cellValue;
                            populatedRow[displayName] = unitCodePrefix && UNIT_CODE_MAPPINGS[unitCodePrefix] ? `${unitCodePrefix}-${mappedUnitName.replace(/^[A-Z]-/,'')}` : mappedUnitName;
                        } else if (dKey === 'uwApprover' || dKey === 'approvalUser') {
                            populatedRow[displayName] = extractName(cellValue);
                        } else {
                            populatedRow[displayName] = cellValue;
                        }
                    });
                    StateManager.originalQueryResults.push(populatedRow);
                });
            } else {
                ALL_DISPLAY_FIELDS_API_KEYS_MAIN.forEach(dKey => {
                    resultRowBase[FIELD_DISPLAY_NAMES_MAP[dKey] || dKey] = '-';
                });
                StateManager.originalQueryResults.push(resultRowBase);
            }
        }

        loadingDialog.overlay.remove();

        if (StateManager.originalQueryResults.length > 0) {
            StateManager.originalQueryResults.sort((a, b) => (parseInt(a[FIELD_DISPLAY_NAMES_MAP.NO]) || 0) - (parseInt(b[FIELD_DISPLAY_NAMES_MAP.NO]) || 0));
        }

        StateManager.a17Mode.isActive = false;
        isEditMode = false;

        renderResultsTableUI(StateManager.originalQueryResults);
        displaySystemNotification(`查詢完成！共處理 ${queryValues.length} 個查詢值`, false, 3500);
    }

    // === 主執行函數 ===
    async function executeCaseQueryTool() {
        // 檢查是否已有工具開啟
        if (document.getElementById(TOOL_MAIN_CONTAINER_ID)) {
            displaySystemNotification('工具已開啟', true);
            return;
        }

        const selectedEnv = await EnvManager.showDialog();
        if (!selectedEnv) {
            displaySystemNotification('操作已取消', true);
            return;
        }

        EnvManager.set(selectedEnv);
        displaySystemNotification(`環境: ${selectedEnv === 'prod' ? '正式' : '測試'}`, false);

        if (!TokenManager.init()) {
            let tokenAttempt = 1;
            while (true) {
                const tokenResult = await TokenManager.showDialog(tokenAttempt);
                if (tokenResult === '_close_tool_') {
                    displaySystemNotification('工具已關閉', false);
                    return;
                }
                if (tokenResult === '_skip_token_') {
                    TokenManager.clear();
                    displaySystemNotification('已略過Token輸入', false);
                    break;
                }
                if (tokenResult === '_token_dialog_cancel_') {
                    displaySystemNotification('Token輸入已取消', true);
                    return;
                }
                if (tokenResult && tokenResult.trim() !== '') {
                    TokenManager.set(tokenResult.trim());
                    displaySystemNotification('Token已設定', false);
                    break;
                } else {
                    tokenAttempt++;
                }
            }
        }

        const querySetupResult = await createQuerySetupDialog();
        if (!querySetupResult) {
            displaySystemNotification('操作已取消', true);
            return;
        }

        await executeQuery(querySetupResult);
    }

    // === 主函數執行 ===
    (async function main() {
        document.getElementById(TOOL_MAIN_CONTAINER_ID)?.remove();

        ['_EnvSelect_overlay', '_Token_overlay', '_QuerySetup_overlay',
            '_A17TextSettings_overlay', '_Loading_overlay', '_CSVPurpose_overlay',
            '_CSVColSelect_overlay', '_CSVCheckbox_overlay', '_RetryEdit_overlay'
        ]
        .forEach(suffix => {
            const el = document.getElementById(TOOL_MAIN_CONTAINER_ID + suffix);
            if (el) el.remove();
        });

        document.getElementById(TOOL_MAIN_CONTAINER_ID + '_Notification')?.remove();

        StateManager.loadA17Settings();

        await executeCaseQueryTool();
    })();

})();
