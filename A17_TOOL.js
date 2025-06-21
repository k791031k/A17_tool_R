javascript: (async () => {
    /*
     * =======================================================================
     * 凱基人壽案件查詢工具 - 終極整合版
     * * 所有 UI 元素皆在工具主視窗內部動態生成與管理。
     * * 整合了自動環境判斷、固定環境視覺提示，以及優化後的 CSV 處理與 UI 邏輯。
     * =======================================================================
     */

    /* --- 主要常數與設定 --- */
    const API_URLS = {
        test: 'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisb/api/caseQuery/query',
        prod: 'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisb/api/caseQuery/query'
    };
    const TOKEN_STORAGE_KEY = 'euisToken';
    const A17_TEXT_SETTINGS_STORAGE_KEY = 'kgilifeQueryTool_A17TextSettings_vFinal';
    const TOOL_MAIN_CONTAINER_ID = 'kgilifeQueryToolMainContainer_vFinal'; // 主視窗的ID

    const Z_INDEX = {
        OVERLAY: 2147483640, // 模態對話框（如Token, 查詢設定）
        MAIN_UI: 2147483630, // 工具主視窗
        NOTIFICATION: 2147483647 // 系統通知
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
    const ALL_DISPLAY_FIELDS_API_KEYS_MAIN = ['applyNumber', 'policyNumber', 'approvalNumber', 'receiptNumber', 'insuredId', 'statusCombined', 'uwApproverUnit', 'uwApprover', 'approvalUser'];

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
        }, {
            id: 'B',
            label: 'B-北一',
            color: '#28a745'
        },
        {
            id: 'P',
            label: 'P-北二',
            color: '#ffc107'
        }, {
            id: 'T',
            label: 'T-桃竹',
            color: '#17a2b8'
        },
        {
            id: 'C',
            label: 'C-台中',
            color: '#fd7e14'
        }, {
            id: 'N',
            label: 'N-台南',
            color: '#6f42c1'
        },
        {
            id: 'K',
            label: 'K-高雄',
            color: '#e83e8c'
        }, {
            id: 'UNDEF',
            label: '查無單位',
            color: '#546e7a'
        }
    ];
    const UNIT_MAP_FIELD_API_KEY = 'uwApproverUnit';

    const A17_DEFAULT_TEXT_CONTENT = "DEAR,\n\n依據【管理報表：A17 新契約異常帳務】所載內容，報表中列示之送金單號碼，涉及多項帳務異常情形，例如：溢繳、短收、取消件需退費、以及無相對應之案件等問題。\n\n本週我們已逐筆查詢該等異常帳務，結果顯示，這些送金單應對應至下表所列之新契約案件。為利後續帳務處理，敬請協助確認各案件之實際帳務狀況，並如有需調整或處理事項，請一併協助辦理，謝謝。";

    /* --- 全域變數與狀態管理 --- */
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
    let toolMainContainerEl = null; // 用於儲存主工具視窗的 DOM 引用

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
                        if (parsed.hasOwnProperty(key)) this.a17Mode.textSettings[key] = parsed[key];
                    }
                } catch (e) {
                    console.error("載入A17文本設定失敗:", e);
                }
            }
        },
        // 使用 structuredClone 進行深度複製，更穩健
        pushSnapshot(description) {
            try {
                // 只複製需要復原的狀態部分
                const snapshot = structuredClone({
                    originalQueryResults: this.originalQueryResults,
                    baseA17MasterData: this.baseA17MasterData,
                    csvImport: this.csvImport, // 假設 csvImport 也需要復原
                    a17Mode: { // 只複製必要的部分，避免循環引用或不必要的複雜性
                        isActive: this.a17Mode.isActive,
                        selectedUnits: new Set(this.a17Mode.selectedUnits),
                        textSettings: structuredClone(this.a17Mode.textSettings)
                    },
                    // currentTable 相關的 DOM 元素不應該被快照化，因為它們會隨著 UI 重新渲染而改變
                    // history 也不應被快照化以避免無限循環
                    isEditMode: isEditMode // 外部的 isEditMode 也要納入考量
                });
                this.history.push({
                    description,
                    snapshot,
                    timestamp: Date.now()
                });
                if (this.history.length > 20) this.history.shift(); // 增加歷史紀錄上限
            } catch (e) {
                console.error("創建快照失敗:", e);
                displaySystemNotification("無法保存歷史狀態，復原功能可能受限。", true);
            }
        },
        undo() {
            if (this.history.length === 0) {
                displaySystemNotification("沒有更多操作可復原", true);
                return;
            }
            try {
                const lastState = this.history.pop();
                // 恢復狀態
                this.originalQueryResults = lastState.snapshot.originalQueryResults;
                this.baseA17MasterData = lastState.snapshot.baseA17MasterData;
                this.csvImport = lastState.snapshot.csvImport;
                this.a17Mode.isActive = lastState.snapshot.a17Mode.isActive;
                this.a17Mode.selectedUnits = new Set(lastState.snapshot.a17Mode.selectedUnits);
                this.a17Mode.textSettings = lastState.snapshot.a17Mode.textSettings;
                isEditMode = lastState.snapshot.isEditMode; // 恢復外部的 isEditMode 變數

                // 重新渲染UI以反映恢復的狀態
                // 需要判斷當前是哪種模式來決定渲染哪個數據源
                if (StateManager.currentTable.mainUIElement) { // 檢查主UI是否已存在
                    renderResultsTableUI(this.a17Mode.isActive ? this.baseA17MasterData : this.originalQueryResults);
                }
                displaySystemNotification(`已復原：${lastState.description}`, false);
            } catch (e) {
                console.error("復原操作失敗:", e);
                displaySystemNotification("復原操作執行失敗，狀態可能不一致。", true);
            }
        }
    };

    /* --- 輔助函數 --- */
    function escapeHtml(unsafe) {
        if (typeof unsafe !== 'string') return unsafe === null || unsafe === undefined ? '' : String(unsafe);
        // 修正了`字符的替換，並使其更完整
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
        // 加入關閉按鈕
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

    // 修改 createDialogBase，使其對話框附加到工具主容器而不是 body
    function createDialogBase(idSuffix, contentHtml, minWidth = '350px', maxWidth = '600px', customStyles = '') {
        // 確保主工具視窗已存在
        if (!toolMainContainerEl) {
            console.error("主工具視窗尚未初始化，無法創建對話框。");
            return; // 或者選擇fallback到document.body，但這會打破UI一致性
        }

        const id = TOOL_MAIN_CONTAINER_ID + idSuffix;
        toolMainContainerEl.querySelector(`#${id}_overlay`)?.remove(); // 在主容器內查找並移除舊的 overlay

        const overlay = document.createElement('div');
        overlay.id = id + '_overlay';
        // Overlay 樣式現在是相對於 toolMainContainerEl 進行定位
        overlay.style.cssText = `position:absolute;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:${Z_INDEX.OVERLAY};display:flex;align-items:center;justify-content:center;font-family:'Microsoft JhengHei',Arial,sans-serif;backdrop-filter:blur(2px);`;
        const dialog = document.createElement('div');
        dialog.id = id + '_dialog';
        dialog.style.cssText = `background:#fff;padding:20px 25px;border-radius:8px;box-shadow:0 5px 20px rgba(0,0,0,0.25);min-width:${minWidth};max-width:${maxWidth};width:auto;animation:qtDialogAppear 0.2s ease-out;${customStyles}`;
        dialog.innerHTML = contentHtml;
        overlay.appendChild(dialog);

        // 將對話框附加到主工具視窗的內容區
        toolMainContainerEl.appendChild(overlay);

        const styleEl = document.createElement('style');
        // 修正了 styleEl.textContent 中的`字符錯誤，並將重複的按鈕樣式提取到這裡
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
      .qt-dialog-btn-purple{background:#6f42c1;color:white;} /* A17按鈕的顏色 */
      .qt-dialog-title{margin:0 0 15px 0;color:#333;font-size:18px;text-align:center;font-weight:600;}

      /* 關鍵修改：強制設定輸入欄位的顏色 */
      .qt-input,.qt-textarea,.qt-select{
        width:calc(100% - 18px);
        padding:9px;
        border:1px solid #ccc;
        border-radius:4px;
        font-size:13px;
        margin-bottom:15px;
        box-sizing:border-box;
        background-color:#ffffff !important;
        color:#333333 !important;
        -webkit-appearance:none;
        appearance:none;
      }

      /* 針對深色模式的特殊處理 */
      .qt-input:focus,.qt-textarea:focus,.qt-select:focus{
        background-color:#ffffff !important;
        color:#333333 !important;
        border-color:#007bff;
        outline:none;
      }

      /* 確保placeholder文字可見 */
      .qt-input::placeholder,.qt-textarea::placeholder{
        color:#666666 !important;
        opacity:1;
      }

      /* 針對color input的特殊處理 */
      input[type="color"].qt-input{
        height:40px;
        padding:2px;
        background-color:#ffffff !important;
      }

      /* 針對number input的特殊處理 */
      input[type="number"].qt-input{
        background-color:#ffffff !important;
        color:#333333 !important;
      }

      .qt-textarea{min-height:70px;resize:vertical;}
      .qt-dialog-flex-end{display:flex;justify-content:flex-end;margin-top:15px;}
      .qt-dialog-flex-between{display:flex;justify-content:space-between;align-items:center;margin-top:15px;}
      .qt-retry-btn, .qt-retry-edit-btn { background:#17a2b8; color:white; border:none; padding:4px 8px; border-radius:3px; font-size:11px; cursor:pointer; margin:2px; }
      .qt-retry-edit-btn { background:#fd7e14; }
      .qt-retry-btn:hover, .qt-retry-edit-btn:hover { opacity:0.8; }
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

      /* 強制設定對話框背景色 */
      .qt-dialog-title{color:#333333 !important;}

      /* 針對預覽區的特殊處理 */
      #qt-a17-preview{
        background-color:#f9f9f9 !important;
        color:#333333 !important;
        border:1px solid #ddd !important;
      }
    `;
        dialog.appendChild(styleEl);

        // ESC鍵關閉監聽器獨立於此函數之外
        const escListener = (e) => {
            if (e.key === 'Escape') {
                overlay.remove();
                document.removeEventListener('keydown', escListener); // 移除監聽器以避免衝突
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
        if (!unitString || typeof unitString !== 'string') return 'UNDEF'; // 修正為 UNDEF 統一
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

    /* --- 環境管理器 (整合了自動判斷與手動切換，並增加了視覺提示) --- */
    const EnvManager = {
        get: () => CURRENT_API_URL,

        // 設定環境並更新顯示
        set: (env) => {
            if (env === 'test') {
                CURRENT_API_URL = API_URLS.test;
                EnvManager.updateDisplay('UAT', '#28a745'); // 綠色
            } else if (env === 'prod') {
                CURRENT_API_URL = API_URLS.prod;
                EnvManager.updateDisplay('PROD', '#dc3545'); // 紅色
            }
        },

        // 自動判斷環境並設定
        autoDetectAndSet: () => {
            const hostname = window.location.hostname;
            if (hostname.includes('uat') || hostname.includes('test')) {
                EnvManager.set('test'); // 自動設定為 UAT
            } else {
                EnvManager.set('prod'); // 自動設定為 Production
            }
        },

        // 顯示環境選擇對話框 (用於手動切換)
        showDialog: async () => {
            return new Promise(resolve => {
                const contentHtml = `<h3 class="qt-dialog-title">選擇查詢環境</h3><div style="display:flex; gap:10px; justify-content:center;"><button id="qt-env-uat" class="qt-dialog-btn qt-dialog-btn-green" style="flex-grow:1;">測試 (UAT)</button><button id="qt-env-prod" class="qt-dialog-btn qt-dialog-btn-orange" style="flex-grow:1;">正式 (PROD)</button></div><div style="text-align:center; margin-top:15px;"><button id="qt-env-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button></div>`;
                const {
                    overlay
                } = createDialogBase('_EnvSelect', contentHtml, '300px', 'auto');

                // 獨立的 ESC 監聽器，避免與 createDialogBase 的衝突
                const dialogEscListener = (e) => {
                    if (e.key === 'Escape') {
                        overlay.remove();
                        document.removeEventListener('keydown', dialogEscListener);
                        resolve(null); // 取消操作
                    }
                };
                document.addEventListener('keydown', dialogEscListener);

                const closeDialog = (value) => {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                    resolve(value);
                };
                overlay.querySelector('#qt-env-uat').onclick = () => closeDialog('test');
                overlay.querySelector('#qt-env-prod').onclick = () => closeDialog('prod');
                overlay.querySelector('#qt-env-cancel').onclick = () => closeDialog(null);
            });
        },

        // 更新主工具視窗內環境顯示
        updateDisplay: (envName, color) => {
            // 確保主工具視窗和其內的 envDisplay 元素已存在
            if (!toolMainContainerEl) return;
            let envDisplay = toolMainContainerEl.querySelector('#qt-env-display-fixed');
            if (!envDisplay) {
                envDisplay = document.createElement('div');
                envDisplay.id = 'qt-env-display-fixed';
                envDisplay.style.cssText = `
                position:absolute; top:10px; right:10px; padding:5px 10px; 
                border-radius:5px; font-weight:bold; color:white; z-index:${Z_INDEX.MAIN_UI + 1};
                font-size:12px; pointer-events:none; /* 不阻擋點擊 */
            `;
                toolMainContainerEl.appendChild(envDisplay);
            }
            envDisplay.textContent = `${envName} 環境`;
            envDisplay.style.backgroundColor = color;
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
            const contentHtml = `<h3 class="qt-dialog-title">API TOKEN 設定</h3><input type="password" id="qt-token-input" class="qt-input" placeholder="請輸入您的 API TOKEN">${attempt > 1 ? `<p style="color:red; font-size:12px; text-align:center; margin-bottom:10px;">Token驗證失敗，請重新輸入。</p>` : ''}<div class="qt-dialog-flex-between"><button id="qt-token-skip" class="qt-dialog-btn qt-dialog-btn-orange">略過</button><div><button id="qt-token-close-tool" class="qt-dialog-btn qt-dialog-btn-red">關閉工具</button><button id="qt-token-ok" class="qt-dialog-btn qt-dialog-btn-blue">${attempt > 1 ? '重試' : '確定'}</button></div></div>`;
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

            // 獨立的 ESC 監聽器
            const dialogEscListener = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener);
                    resolve('_token_dialog_cancel_'); // 取消操作
                }
            };
            document.addEventListener('keydown', dialogEscListener);

            const closeDialog = (value) => {
                overlay.remove();
                document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                resolve(value);
            };
            overlay.querySelector('#qt-token-ok').onclick = () => closeDialog(inputEl.value.trim());
            overlay.querySelector('#qt-token-close-tool').onclick = () => closeDialog('_close_tool_');
            overlay.querySelector('#qt-token-skip').onclick = () => closeDialog('_skip_token_');
        })
    };

    /* --- 對話框建立函數 (CSV 相關功能模組化) --- */
    function createQuerySetupDialog() {
        return new Promise(resolve => {
            const queryButtonsHtml = QUERYABLE_FIELD_DEFINITIONS.map(def => `<button class="qt-querytype-btn" data-apikey="${def.queryApiKey}" style="background-color:${def.color}; color:white;">${escapeHtml(def.queryDisplayName)}</button>`).join('');
            const contentHtml = `<h3 class="qt-dialog-title">查詢條件設定</h3><div style="margin-bottom:10px; font-size:13px; color:#555;">選擇查詢欄位類型：</div><div id="qt-querytype-buttons" style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:15px;">${queryButtonsHtml}</div><div style="margin-bottom:5px; font-size:13px; color:#555;">輸入查詢值 (可多筆，以換行/空格/逗號/分號分隔)：</div><textarea id="qt-queryvalues-input" class="qt-textarea" placeholder="請先選擇上方查詢欄位類型"></textarea><div style="margin-bottom:15px;"><button id="qt-csv-import-btn" class="qt-dialog-btn qt-dialog-btn-grey" style="margin-left:0;">從CSV/TXT匯入...</button><span id="qt-csv-filename-display" style="font-size:12px; color:#666; margin-left:10px;"></span></div><div class="qt-dialog-flex-between"><button id="qt-clear-all-input-btn" class="qt-dialog-btn qt-dialog-btn-orange">清除所有輸入</button><div><button id="qt-querysetup-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button><button id="qt-querysetup-ok" class="qt-dialog-btn qt-dialog-btn-blue">開始查詢</button></div><button id="qt-env-switch-btn" class="qt-dialog-btn qt-dialog-btn-blue">環境</button></div><input type="file" id="qt-file-input-hidden" accept=".csv,.txt" style="display:none;">`; // 新增「環境」按鈕

            const {
                overlay
            } = createDialogBase('_QuerySetup', contentHtml, '480px', 'auto');
            const queryValuesInput = overlay.querySelector('#qt-queryvalues-input');
            const typeButtons = overlay.querySelectorAll('.qt-querytype-btn');
            const csvImportBtn = overlay.querySelector('#qt-csv-import-btn');
            const fileInputHidden = overlay.querySelector('#qt-file-input-hidden');
            const csvFilenameDisplay = overlay.querySelector('#qt-csv-filename-display');
            const envSwitchBtn = overlay.querySelector('#qt-env-switch-btn'); // 取得「環境」按鈕

            // 綁定「環境」按鈕事件
            if (envSwitchBtn) {
                envSwitchBtn.onclick = async () => {
                    const selectedEnv = await EnvManager.showDialog(); // 彈出選擇對話框
                    if (selectedEnv) {
                        EnvManager.set(selectedEnv); // 手動設定環境
                        displaySystemNotification(`環境已手動切換為: ${selectedEnv === 'prod' ? '正式' : '測試'}`, false);
                    } else {
                        displaySystemNotification('環境切換已取消', true);
                    }
                };
            }

            function setActiveButton(apiKey) {
                typeButtons.forEach(btn => {
                    const isSelected = btn.dataset.apikey === apiKey;
                    btn.style.border = isSelected ? `2px solid ${btn.style.backgroundColor}` : '2px solid transparent';
                    btn.style.boxShadow = isSelected ? `0 0 8px ${btn.style.backgroundColor}70` : 'none';
                    if (isSelected) {
                        selectedQueryDefinitionGlobal = QUERYABLE_FIELD_DEFINITIONS.find(d => d.queryApiKey === apiKey);
                        queryValuesInput.placeholder = `請輸入${selectedQueryDefinitionGlobal.queryDisplayName}(可多筆...)`;
                    }
                });
            }

            typeButtons.forEach(btn => btn.onclick = () => {
                setActiveButton(btn.dataset.apikey);
                queryValuesInput.focus();
            });
            setActiveButton(selectedQueryDefinitionGlobal.queryApiKey);
            csvImportBtn.onclick = () => fileInputHidden.click();
            fileInputHidden.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) return;
                csvFilenameDisplay.textContent = `已選: ${file.name}`;
                try {
                    const text = await file.text();
                    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
                    if (lines.length === 0) {
                        displaySystemNotification('CSV檔案為空', true);
                        return;
                    }
                    const headers = lines[0].split(/,|;|\t/).map(h => h.trim().replace(/^"|"$/g, ''));
                    const purpose = await createCSVPurposeDialog();
                    if (!purpose) {
                        csvFilenameDisplay.textContent = '';
                        fileInputHidden.value = '';
                        return;
                    }
                    if (purpose === 'fillQueryValues') {
                        const columnIndex = await createCSVColumnSelectionDialog(headers, "選擇包含查詢值的欄位：");
                        if (columnIndex === null || columnIndex === columnIndex) {
                            csvFilenameDisplay.textContent = '';
                            fileInputHidden.value = '';
                            return;
                        } //修正了這裡的邏輯
                        const values = [];
                        for (let i = 1; i < lines.length; i++) {
                            const cols = lines[i].split(/,|;|\t/).map(c => c.trim().replace(/^"|"$/g, ''));
                            if (cols[columnIndex] && cols[columnIndex].trim() !== "") values.push(cols[columnIndex].trim());
                        }
                        queryValuesInput.value = Array.from(new Set(values)).join('\n');
                        displaySystemNotification('查詢值已從CSV填入', false);
                        StateManager.csvImport = {
                            ...StateManager.csvImport,
                            fileName: file.name,
                            rawHeaders: headers,
                            rawData: lines.slice(1).map(line => line.split(/,|;|\t/).map(c => c.trim().replace(/^"|"$/g, ''))),
                            selectedColForQueryName: headers[columnIndex],
                            isA17CsvPrepared: false,
                            selectedColsForA17Merge: []
                        };
                    } else if (purpose === 'prepareA17Merge') {
                        const selectedHeadersForA17 = await createCSVColumnCheckboxDialog(headers, "勾選要在A17表格中顯示的CSV欄位：");
                        if (!selectedHeadersForA17 || selectedHeadersForA17.length === 0) {
                            csvFilenameDisplay.textContent = '';
                            fileInputHidden.value = '';
                            return;
                        }
                        StateManager.csvImport = {
                            ...StateManager.csvImport,
                            fileName: file.name,
                            rawHeaders: headers,
                            rawData: lines.slice(1).map(line => line.split(/,|;|\t/).map(c => c.trim().replace(/^"|"$/g, ''))),
                            selectedColsForA17Merge: selectedHeadersForA17,
                            isA17CsvPrepared: true,
                            selectedColForQueryName: null
                        };
                        displaySystemNotification(`已選 ${selectedHeadersForA17.length} 個CSV欄位供A17合併`, false);
                    }
                } catch (err) {
                    console.error("處理CSV錯誤:", err);
                    displaySystemNotification('讀取CSV失敗', true);
                    csvFilenameDisplay.textContent = '';
                }
                fileInputHidden.value = '';
            };

            overlay.querySelector('#qt-clear-all-input-btn').onclick = () => {
                queryValuesInput.value = '';
                csvFilenameDisplay.textContent = '';
                StateManager.csvImport = {
                    fileName: '',
                    rawHeaders: [],
                    rawData: [],
                    selectedColForQueryName: null,
                    selectedColsForA17Merge: [],
                    isA17CsvPrepared: false
                };
                fileInputHidden.value = '';
                displaySystemNotification('所有輸入已清除', false);
            };

            // 獨立的 ESC 監聽器
            const dialogEscListener = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener);
                    resolve(null); // 取消操作
                }
            };
            document.addEventListener('keydown', dialogEscListener);

            const closeDialog = (value) => {
                overlay.remove();
                document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                resolve(value);
            };
            overlay.querySelector('#qt-querysetup-ok').onclick = () => {
                const values = queryValuesInput.value.trim();
                if (!selectedQueryDefinitionGlobal) {
                    displaySystemNotification('請選查詢欄位類型', true);
                    return;
                }
                if (!values) {
                    displaySystemNotification(`請輸入${selectedQueryDefinitionGlobal.queryDisplayName}`, true);
                    queryValuesInput.focus();
                    return;
                }
                closeDialog({
                    selectedApiKey: selectedQueryDefinitionGlobal.queryApiKey,
                    queryValues: values
                });
            };
            overlay.querySelector('#qt-querysetup-cancel').onclick = () => closeDialog(null);
        });
    }

    function createCSVPurposeDialog() {
        return new Promise(resolve => {
            const contentHtml = `<h3 class="qt-dialog-title">選擇CSV檔案用途</h3><div style="display:flex; flex-direction:column; gap:10px;"><button id="qt-csv-purpose-query" class="qt-dialog-btn qt-dialog-btn-blue" style="margin-left:0;">將CSV某欄作為查詢值</button><button id="qt-csv-purpose-a17" class="qt-dialog-btn qt-dialog-btn-green" style="margin-left:0;">勾選CSV欄位供A17合併顯示</button></div><div style="text-align:center; margin-top:15px;"><button id="qt-csv-purpose-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button></div>`;
            const {
                overlay
            } = createDialogBase('_CSVPurpose', contentHtml, '300px', 'auto');

            // 獨立的 ESC 監聽器
            const dialogEscListener = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener);
                    resolve(null); // 取消操作
                }
            };
            document.addEventListener('keydown', dialogEscListener);

            const closeDialog = (value) => {
                overlay.remove();
                document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                resolve(value);
            };
            overlay.querySelector('#qt-csv-purpose-query').onclick = () => closeDialog('fillQueryValues');
            overlay.querySelector('#qt-csv-purpose-a17').onclick = () => closeDialog('prepareA17Merge');
            overlay.querySelector('#qt-csv-purpose-cancel').onclick = () => closeDialog(null);
        });
    }

    function createCSVColumnSelectionDialog(headers, title) {
        return new Promise(resolve => {
            let optionsHtml = headers.map((header, index) => `<button class="qt-dialog-btn qt-dialog-btn-blue" data-index="${index}" style="margin:5px; width: calc(50% - 10px); text-overflow: ellipsis; overflow: hidden; white-space: nowrap;">${escapeHtml(header)}</button>`).join('');
            const contentHtml = `<h3 class="qt-dialog-title">${escapeHtml(title)}</h3><div style="display:flex; flex-wrap:wrap; justify-content:center; max-height:300px; overflow-y:auto; margin-bottom:15px;">${optionsHtml}</div><div style="text-align:center;"><button id="qt-csvcol-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button></div>`;
            const {
                overlay
            } = createDialogBase('_CSVColSelect', contentHtml, '400px', 'auto');

            // 獨立的 ESC 監聽器
            const dialogEscListener = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener);
                    resolve(null); // 取消操作
                }
            };
            document.addEventListener('keydown', dialogEscListener);

            const closeDialog = (value) => {
                overlay.remove();
                document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                resolve(value);
            };
            overlay.querySelectorAll('.qt-dialog-btn[data-index]').forEach(btn => {
                btn.onclick = () => closeDialog(parseInt(btn.dataset.index));
            });
            overlay.querySelector('#qt-csvcol-cancel').onclick = () => closeDialog(null);
        });
    }

    function createCSVColumnCheckboxDialog(headers, title) {
        return new Promise(resolve => {
            let checkboxesHtml = headers.map((header, index) => `<div style="margin-bottom: 8px; display:flex; align-items:center;"><input type="checkbox" id="qt-csv-header-cb-${index}" value="${escapeHtml(header)}" style="margin-right:8px; transform:scale(1.2);"><label for="qt-csv-header-cb-${index}" style="font-size:14px;">${escapeHtml(header)}</label></div>`).join('');
            const contentHtml = `<h3 class="qt-dialog-title">${escapeHtml(title)}</h3><div style="max-height: 300px; overflow-y: auto; margin-bottom: 15px; border: 1px solid #eee; padding: 10px; border-radius: 4px;">${checkboxesHtml}</div><div class="qt-dialog-flex-end"><button id="qt-csvcb-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button><button id="qt-csvcb-ok" class="qt-dialog-btn qt-dialog-btn-blue">確定勾選</button></div>`;
            const {
                overlay
            } = createDialogBase('_CSVCheckbox', contentHtml, '400px', 'auto');

            // 獨立的 ESC 監聽器
            const dialogEscListener = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener);
                    resolve(null); // 取消操作
                }
            };
            document.addEventListener('keydown', dialogEscListener);

            const closeDialog = (value) => {
                overlay.remove();
                document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                resolve(value);
            };
            overlay.querySelector('#qt-csvcb-ok').onclick = () => {
                const selected = [];
                overlay.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => selected.push(cb.value)); // 修正為 overlay.querySelectorAll
                if (selected.length === 0) {
                    displaySystemNotification('請至少勾選一個欄位', true);
                    return;
                }
                closeDialog(selected);
            };
            overlay.querySelector('#qt-csvcb-cancel').onclick = () => closeDialog(null);
        });
    }

    function createA17TextSettingDialog() {
        return new Promise(resolve => {
            const s = StateManager.a17Mode.textSettings;
            const contentHtml = `<h3 class="qt-dialog-title">A17 通知文本設定</h3><div style="display:grid; grid-template-columns: 1fr; gap: 15px;"><div><label for="qt-a17-mainContent" style="font-weight:bold; font-size:13px; display:block; margin-bottom:5px;">主文案內容：</label><textarea id="qt-a17-mainContent" class="qt-textarea" style="height:150px;">${escapeHtml(s.mainContent)}</textarea><div style="display:flex; gap:10px; margin-top:5px; flex-wrap:wrap;"><label style="font-size:12px;">字體大小: <input type="number" id="qt-a17-mainFontSize" value="${s.mainFontSize}" min="8" max="24" step="0.5" class="qt-input" style="width:60px; padding:3px; margin-bottom:0;"> pt</label><label style="font-size:12px;">行高: <input type="number" id="qt-a17-mainLineHeight" value="${s.mainLineHeight}" min="1" max="3" step="0.1" class="qt-input" style="width:60px; padding:3px; margin-bottom:0;"> 倍</label><label style="font-size:12px;">顏色: <input type="color" id="qt-a17-mainFontColor" value="${s.mainFontColor}" style="padding:1px; height:25px; vertical-align:middle;"></label></div></div><div><label style="font-weight:bold; font-size:13px; display:block; margin-bottom:5px;">動態日期設定 (相對於今天)：</label><div style="display:flex; gap:15px; align-items:center; margin-bottom:5px;flex-wrap:wrap;"><label style="font-size:12px;">產檔時間偏移: <input type="number" id="qt-a17-genDateOffset" value="${s.genDateOffset}" class="qt-input" style="width:60px; padding:3px; margin-bottom:0;"> 天</label><label style="font-size:12px;">比對時間偏移: <input type="number" id="qt-a17-compDateOffset" value="${s.compDateOffset}" class="qt-input" style="width:60px; padding:3px; margin-bottom:0;"> 天</label></div><div style="display:flex; gap:10px; flex-wrap:wrap;"><label style="font-size:12px;">日期字體大小: <input type="number" id="qt-a17-dateFontSize" value="${s.dateFontSize}" min="6" max="16" step="0.5" class="qt-input" style="width:60px; padding:3px; margin-bottom:0;"> pt</label><label style="font-size:12px;">日期行高: <input type="number" id="qt-a17-dateLineHeight" value="${s.dateLineHeight}" min="1" max="3" step="0.1" class="qt-input" style="width:60px; padding:3px; margin-bottom:0;"> 倍</label><label style="font-size:12px;">日期顏色: <input type="color" id="qt-a17-dateFontColor" value="${s.dateFontColor}" style="padding:1px; height:25px; vertical-align:middle;"></label></div></div><div><label style="font-weight:bold; font-size:13px; display:block; margin-bottom:5px;">預覽效果 (此區可臨時編輯，僅影響當次複製)：</label><div id="qt-a17-preview" contenteditable="true" style="border:1px solid #ccc; padding:10px; min-height:100px; max-height:200px; overflow-y:auto; font-size:${s.mainFontSize}pt; line-height:${s.mainLineHeight}; color:${s.mainFontColor}; background:#f9f9f9; border-radius:4px;"></div></div></div><div class="qt-dialog-flex-between" style="margin-top:20px;"><button id="qt-a17-text-reset" class="qt-dialog-btn qt-dialog-btn-orange">重設預設</button><div><button id="qt-a17-text-cancel" class="qt-dialog-btn qt-dialog-btn-grey">取消</button><button id="qt-a17-text-save" class="qt-dialog-btn qt-dialog-btn-blue">儲存設定</button></div></div>`;
            const {
                overlay
            } = createDialogBase('_A17TextSettings', contentHtml, '550px', 'auto');
            const previewEl = overlay.querySelector('#qt-a17-preview');
            const getSettingsFromUI = () => ({
                mainContent: overlay.querySelector('#qt-a17-mainContent').value,
                mainFontSize: parseFloat(overlay.querySelector('#qt-a17-mainFontSize').value),
                mainLineHeight: parseFloat(overlay.querySelector('#qt-a17-mainLineHeight').value),
                mainFontColor: overlay.querySelector('#qt-a17-mainFontColor').value,
                dateFontSize: parseFloat(overlay.querySelector('#qt-a17-dateFontSize').value),
                dateLineHeight: parseFloat(overlay.querySelector('#qt-a17-dateLineHeight').value),
                dateFontColor: overlay.querySelector('#qt-a17-dateFontColor').value,
                genDateOffset: parseInt(overlay.querySelector('#qt-a17-genDateOffset').value),
                compDateOffset: parseInt(overlay.querySelector('#qt-a17-compDateOffset').value)
            });
            const updatePreview = () => {
                const currentUISettings = getSettingsFromUI();
                const today = new Date();
                const genDate = new Date(today);
                genDate.setDate(today.getDate() + currentUISettings.genDateOffset);
                const compDate = new Date(today);
                compDate.setDate(today.getDate() + currentUISettings.compDateOffset);
                const genDateStr = formatDate(genDate);
                const compDateStr = formatDate(compDate);
                let previewContent = escapeHtml(currentUISettings.mainContent).replace(/\n/g, '<br>') + `<br><br><span class="qt-a17-dynamic-date" style="font-size:${currentUISettings.dateFontSize}pt; line-height:${currentUISettings.dateLineHeight}; color:${currentUISettings.dateFontColor};">產檔時間：${genDateStr}<br>比對時間：${compDateStr}</span>`;
                previewEl.innerHTML = previewContent;
                previewEl.style.fontSize = currentUISettings.mainFontSize + 'pt';
                previewEl.style.lineHeight = currentUISettings.mainLineHeight;
                previewEl.style.color = currentUISettings.mainFontColor;
            };
            ['#qt-a17-mainContent', '#qt-a17-mainFontSize', '#qt-a17-mainLineHeight', '#qt-a17-mainFontColor', '#qt-a17-dateFontSize', '#qt-a17-dateLineHeight', '#qt-a17-dateFontColor', '#qt-a17-genDateOffset', '#qt-a17-compDateOffset'].forEach(selector => {
                const el = overlay.querySelector(selector);
                if (el.type === 'color') el.onchange = updatePreview;
                else el.oninput = updatePreview;
            });
            updatePreview();

            // 獨立的 ESC 監聽器
            const dialogEscListener = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', dialogEscListener);
                    resolve(null); // 取消操作
                }
            };
            document.addEventListener('keydown', dialogEscListener);

            const closeDialog = (value) => {
                overlay.remove();
                document.removeEventListener('keydown', dialogEscListener); // 確保移除監聽器
                resolve(value);
            };
            overlay.querySelector('#qt-a17-text-save').onclick = () => {
                const newSettings = getSettingsFromUI();
                if (!newSettings.mainContent.trim()) {
                    displaySystemNotification('主文案內容不可為空', true);
                    return;
                }
                StateManager.a17Mode.textSettings = newSettings;
                localStorage.setItem(A17_TEXT_SETTINGS_STORAGE_KEY, JSON.stringify(newSettings));
                displaySystemNotification('A17文本設定已儲存', false);
                closeDialog(true);
            };
            overlay.querySelector('#qt-a17-text-cancel').onclick = () => closeDialog(null);
            overlay.querySelector('#qt-a17-text-reset').onclick = () => {
                overlay.querySelector('#qt-a17-mainContent').value = A17_DEFAULT_TEXT_CONTENT;
                overlay.querySelector('#qt-a17-mainFontSize').value = 12;
                overlay.querySelector('#qt-a17-mainLineHeight').value = 1.5;
                overlay.querySelector('#qt-a17-mainFontColor').value = '#333333';
                overlay.querySelector('#qt-a17-dateFontSize').value = 8;
                overlay.querySelector('#qt-a17-dateLineHeight').value = 1.2;
                overlay.querySelector('#qt-a17-dateFontColor').value = '#555555';
                overlay.querySelector('#qt-a17-genDateOffset').value = -3;
                overlay.querySelector('#qt-a17-compDateOffset').value = 0;
                updatePreview();
            };
        });
    }

    async function performApiQuery(queryValue, apiKey) {
        const reqBody = {
            currentPage: 1,
            pageSize: 10
        };
        reqBody[apiKey] = queryValue;
        const fetchOpts = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(reqBody),
        };
        if (apiAuthToken) fetchOpts.headers['SSO-TOKEN'] = apiAuthToken;
        let retries = 1;
        while (retries >= 0) {
            try {
                const res = await fetch(CURRENT_API_URL, fetchOpts);
                const data = await res.json();
                if (res.status === 401) {
                    apiAuthToken = null;
                    TokenManager.clear();
                    return {
                        error: 'token_invalid',
                        data: null
                    };
                }
                if (!res.ok) {
                    throw new Error(`API請求錯誤: ${res.status} ${res.statusText}`);
                }
                return {
                    error: null,
                    data: data,
                    success: data && data.records && data.records.length > 0
                };
            } catch (e) {
                console.error(`查詢 ${queryValue} 錯誤 (嘗試 ${2-retries}):`, e);
                if (retries > 0) {
                    displaySystemNotification(`查詢 ${queryValue} 失敗，2秒後重試...`, true, 1800);
                    await new Promise(r => setTimeout(r, 2000));
                    retries--;
                } else {
                    return {
                        error: 'network_error',
                        data: null
                    };
                }
            }
        }
    }

    // --- 表格渲染及操作函數 (保持版本 B 的優化) ---
    function renderResultsTableUI(dataToRender) {
        // 不再移除 toolMainContainerEl，因為它是整個工具的根元素
        // StateManager.currentTable.mainUIElement?.remove(); // 這行應該被移除或處理為清空內容

        // 清空主視窗內容，以便重新渲染表格
        toolMainContainerEl.innerHTML = `
        <div class="qt-main-title-bar">凱基人壽案件查詢結果</div>
        <div class="qt-main-content-wrapper">
            <div class="qt-main-controls-header">
                <div class="qt-summary-section"></div>
                <div class="qt-buttons-group-left"></div>
                <div class="qt-filter-input-wrapper">
                    <input type="text" placeholder="篩選表格內容..." class="qt-input qt-filter-input">
                </div>
                <div class="qt-buttons-group-right"></div>
            </div>
            <div class="qt-a17-unit-buttons-container"></div>
            <div class="qt-a17-text-controls-container"></div>
            <div class="qt-table-scroll-wrapper">
                <table class="qt-table-results">
                    <thead></thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
    `;

        // 重新獲取新的 DOM 元素引用
        StateManager.currentTable.mainUIElement = toolMainContainerEl; // 指向整個工具容器
        StateManager.currentTable.tableHeadElement = toolMainContainerEl.querySelector('.qt-table-results thead');
        StateManager.currentTable.tableBodyElement = toolMainContainerEl.querySelector('.qt-table-results tbody');
        StateManager.currentTable.a17UnitButtonsContainer = toolMainContainerEl.querySelector('.qt-a17-unit-buttons-container');

        // 環境提示已經存在於 toolMainContainerEl，無需重新創建或更新其引用
        // EnvManager.updateDisplay(); // 可在此處再次呼叫確保顯示狀態

        const titleBar = toolMainContainerEl.querySelector('.qt-main-title-bar');
        const controlsHeader = toolMainContainerEl.querySelector('.qt-main-controls-header');
        const summarySec = toolMainContainerEl.querySelector('.qt-summary-section');
        const filterInput = toolMainContainerEl.querySelector('.qt-filter-input');
        const buttonsGroupLeft = toolMainContainerEl.querySelector('.qt-buttons-group-left');
        const buttonsGroupRight = toolMainContainerEl.querySelector('.qt-buttons-group-right');
        const a17UnitBtnsCtr = toolMainContainerEl.querySelector('.qt-a17-unit-buttons-container');
        const a17TextControls = toolMainContainerEl.querySelector('.qt-a17-text-controls-container');

        // 重新綁定拖曳功能（因為內容被清空了）
        titleBar.onmousedown = (e) => {
            if (e.target !== titleBar) return;
            e.preventDefault();
            dragState.dragging = true;
            dragState.startX = e.clientX;
            dragState.startY = e.clientY;
            dragState.initialX = toolMainContainerEl.offsetLeft;
            dragState.initialY = toolMainContainerEl.offsetTop;
            titleBar.style.cursor = 'grabbing';
            toolMainContainerEl.style.transform = 'none';
        };
        document.onmousemove = (e) => {
            if (dragState.dragging) {
                const dx = e.clientX - dragState.startX;
                const dy = e.clientY - dragState.startY;
                toolMainContainerEl.style.left = (dragState.initialX + dx) + 'px';
                toolMainContainerEl.style.top = (dragState.initialY + dy) + 'px';
            }
        };
        document.onmouseup = () => {
            if (dragState.dragging) {
                dragState.dragging = false;
                titleBar.style.cursor = 'grab';
            }
        };

        // 清除並重新創建按鈕，以便重新綁定事件
        buttonsGroupLeft.innerHTML = '';
        buttonsGroupRight.innerHTML = '';

        [{
                id: 'ClearConditions',
                text: '清除條件',
                cls: 'qt-dialog-btn-grey',
                group: buttonsGroupLeft
            },
            {
                id: 'Requery',
                text: '重新查詢',
                cls: 'qt-dialog-btn-orange',
                group: buttonsGroupLeft
            },
            {
                id: 'A17',
                text: 'A17作業',
                cls: 'qt-dialog-btn-purple',
                group: buttonsGroupLeft
            },
            {
                id: 'CopyTable',
                text: '複製表格',
                cls: 'qt-dialog-btn-green',
                group: buttonsGroupLeft
            },
            {
                id: 'EditMode',
                text: '編輯模式',
                cls: 'qt-dialog-btn-blue',
                group: buttonsGroupLeft
            },
            {
                id: 'AddRow',
                text: '+ 新增列',
                cls: 'qt-dialog-btn-blue',
                group: buttonsGroupLeft,
                style: 'display:none;'
            }
        ].forEach(cfg => {
            const btn = document.createElement('button');
            btn.id = TOOL_MAIN_CONTAINER_ID + '_btn' + cfg.id;
            btn.textContent = cfg.text;
            btn.className = `qt-dialog-btn ${cfg.cls}`;
            if (cfg.style) btn.style.cssText += cfg.style;
            cfg.group.appendChild(btn)
        });

        const closeBtn = document.createElement('button');
        closeBtn.id = TOOL_MAIN_CONTAINER_ID + '_btnCloseTool';
        closeBtn.textContent = '關閉工具';
        closeBtn.className = 'qt-dialog-btn qt-dialog-btn-red';
        buttonsGroupRight.appendChild(closeBtn);

        // --- 控制按鈕事件綁定 ---
        filterInput.oninput = () => applyTableFilter();
        toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnClearConditions`).onclick = handleClearConditions;
        toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnRequery`).onclick = () => {
            // 重新啟動工具，會再次觸發環境自動判斷和初始化
            executeCaseQueryTool();
        };

        const a17Btn = toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnA17`);
        a17Btn.onmousedown = (e) => {
            if (e.button !== 0) return;
            a17ButtonLongPressTimer = setTimeout(() => {
                a17ButtonLongPressTimer = null;
                toggleA17Mode(true);
            }, 700);
        };
        a17Btn.onmouseup = () => {
            if (a17ButtonLongPressTimer) {
                clearTimeout(a17ButtonLongPressTimer);
                a17ButtonLongPressTimer = null;
                toggleA17Mode(false);
            }
        };
        a17Btn.onmouseleave = () => {
            if (a17ButtonLongPressTimer) {
                clearTimeout(a17ButtonLongPressTimer);
                a17ButtonLongPressTimer = null;
            }
        };
        toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnCopyTable`).onclick = handleCopyTable;
        toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnEditMode`).onclick = toggleEditMode;
        toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnAddRow`).onclick = handleAddRowToTable;
        closeBtn.onclick = () => {
            if (confirm('確定要關閉查詢工具嗎？')) {
                toolMainContainerEl.remove(); // 移除整個工具主視窗
                document.removeEventListener('keydown', mainUIEscListener); // 移除ESC監聽
                displaySystemNotification('查詢工具已關閉', false);
            }
        };
        toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnA17EditText`).onclick = async () => {
            await createA17TextSettingDialog();
        };

        // --- 主 UI 的 ESC 關閉監聽器 ---
        // 每次渲染 UI 時綁定，舊的會被新的覆蓋，或者確保移除舊的
        document.removeEventListener('keydown', mainUIEscListener); // 確保移除舊的
        document.addEventListener('keydown', mainUIEscListener); // 綁定新的

        updateSummaryCount(dataToRender.length);
        if (StateManager.a17Mode.isActive) { // 修正為 StateManager.a17Mode.isActive
            renderA17ModeUI();
            populateTableRows(StateManager.baseA17MasterData);
            updateA17UnitButtonCounts();
        } else {
            renderNormalModeUI();
            populateTableRows(dataToRender);
        }
    }

    // --- 主 UI ESC 關閉的獨立監聽器 ---
    const mainUIEscListener = (e) => {
        // 檢查是否是 ESC 鍵，並且沒有其他對話框開啟
        // 查找 toolMainContainerEl 內的 overlay
        if (e.key === 'Escape' && toolMainContainerEl && !toolMainContainerEl.querySelector(`[id^="${TOOL_MAIN_CONTAINER_ID}_"][id$="_overlay"]`)) {
            if (confirm('確定要關閉查詢工具嗎？')) {
                toolMainContainerEl.remove(); // 移除整個工具主視窗
                document.removeEventListener('keydown', mainUIEscListener);
                displaySystemNotification('查詢工具已關閉', false);
            }
        }
    };

    // --- 表格相關輔助函數 ---
    function updateSummaryCount(visibleRowCount) {
        const summaryEl = toolMainContainerEl?.querySelector('.qt-summary-section'); // 修正為 toolMainContainerEl 內查找
        if (!summaryEl) return;
        let baseDataCount = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData.length : StateManager.originalQueryResults.length;
        let text = `查詢結果：<strong>${baseDataCount}</strong>筆`;
        const filterInput = toolMainContainerEl.querySelector('.qt-filter-input'); // 修正為 toolMainContainerEl 內查找
        const isFiltered = (filterInput && filterInput.value.trim() !== '') || (StateManager.a17Mode.isActive && StateManager.a17Mode.selectedUnits.size > 0);
        if (isFiltered && visibleRowCount !== baseDataCount) {
            text += ` (篩選後顯示 <strong>${visibleRowCount}</strong> 筆)`;
        }
        summaryEl.innerHTML = text;
    }


    function sortTableByColumn(headerKey) {
        const currentDirection = StateManager.currentTable.sortDirections[headerKey] || 'asc';
        const newDirection = currentDirection === 'asc' ? 'desc' : 'asc';
        StateManager.currentTable.sortDirections = {}; // 清除其他欄位的排序方向
        StateManager.currentTable.sortDirections[headerKey] = newDirection;

        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
        currentData.sort((a, b) => {
            let aVal = a[headerKey] || '';
            let bVal = b[headerKey] || '';

            // 數值類型轉換 (針對序號)
            if (headerKey === FIELD_DISPLAY_NAMES_MAP.NO) {
                aVal = parseInt(String(aVal)) || 0;
                bVal = parseInt(String(bVal)) || 0;
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
        const thead = StateManager.currentTable.tableHeadElement;
        thead.querySelectorAll('th .sort-arrow').forEach(arrow => arrow.remove());
        const targetTh = Array.from(thead.querySelectorAll('th')).find(th => th.textContent.includes(headerKey));
        if (targetTh) {
            const arrow = document.createElement('span');
            arrow.className = 'sort-arrow';
            arrow.innerHTML = newDirection === 'asc' ? ' &#9650;' : ' &#9660;'; // 使用 unicode 三角形箭頭
            arrow.style.cssText = 'color:#007bff;font-weight:bold;';
            targetTh.appendChild(arrow);
        }
    }

    function startCellEdit(td, row, headerKey, rowIndex) {
        if (td.querySelector('input') || td.querySelector('select')) return; // 避免重複編輯

        StateManager.pushSnapshot('編輯儲存格'); // 在編輯前保存狀態

        const originalValue = row[headerKey] === null || row[headerKey] === undefined ? '' : String(row[headerKey]);
        const originalHtml = td.innerHTML;
        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
        const currentRowElement = StateManager.currentTable.tableBodyElement.children[rowIndex];
        if (currentRowElement) {
            currentRowElement.classList.add('qt-editing-row'); // 標記正在編輯的行
            currentRowElement.style.backgroundColor = '#fffacd'; // 編輯中的顏色
        }


        const saveEdit = () => {
            let newValue;
            if (select) {
                newValue = select.value;
            } else {
                newValue = input.value.trim();
            }

            row[headerKey] = newValue; // 更新數據模型
            td.innerHTML = escapeHtml(newValue); // 更新顯示
            displaySystemNotification('已更新儲存格', false);

            if (currentRowElement) {
                currentRowElement.classList.remove('qt-editing-row');
                // 恢復原始行顏色
                currentRowElement.style.backgroundColor = rowIndex % 2 ? '#f8f9fa' : '#ffffff';
            }

            // 移除事件監聽器
            cleanupListeners();
        };

        const cancelEdit = () => {
            td.innerHTML = originalHtml; // 恢復原始內容
            displaySystemNotification('編輯已取消', true, 1500);
            if (currentRowElement) {
                currentRowElement.classList.remove('qt-editing-row');
                currentRowElement.style.backgroundColor = rowIndex % 2 ? '#f8f9fa' : '#ffffff';
            }
            cleanupListeners();
        };

        let input, select;

        if (headerKey === FIELD_DISPLAY_NAMES_MAP.uwApproverUnit) {
            select = document.createElement('select');
            select.className = 'qt-select';
            select.style.cssText = 'width:100%; margin:0;';

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
            select.onblur = saveEdit;
            select.onchange = saveEdit; // 選擇後立即保存
            select.onkeydown = e => {
                if (e.key === 'Enter') saveEdit();
                if (e.key === 'Escape') cancelEdit();
            };
        } else {
            input = document.createElement('input');
            input.type = 'text';
            input.className = 'qt-input';
            input.style.cssText = 'width:100%; margin:0; box-sizing:border-box;'; // 添加 box-sizing
            input.value = originalValue;

            td.innerHTML = '';
            td.appendChild(input);
            input.focus();
            input.select();
            input.onblur = saveEdit;
            input.onkeydown = e => {
                if (e.key === 'Enter') saveEdit();
                if (e.key === 'Escape') cancelEdit();
            };
        }

        // 清理監聽器
        const cleanupListeners = () => {
            if (input) {
                input.onblur = null;
                input.onkeydown = null;
            }
            if (select) {
                select.onblur = null;
                select.onchange = null;
                select.onkeydown = null;
            }
        };
    }

    function handleDeleteRow(rowIndex) {
        if (confirm('確定要刪除這一列嗎？此操作不可逆。')) {
            StateManager.pushSnapshot('刪除列');
            const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
            currentData.splice(rowIndex, 1);
            // 更新序號 (如果顯示序號)
            currentData.forEach((row, idx) => {
                row[FIELD_DISPLAY_NAMES_MAP.NO] = String(idx + 1);
            });
            populateTableRows(currentData);
            displaySystemNotification('已刪除列', false);
        }
    }

    function handleClearConditions() {
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

            toolMainContainerEl.remove(); // 移除整個工具主視窗
            displaySystemNotification('已清除所有條件', false);
        }
    }

    function handleCopyTable() {
        const includeText = toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_cbA17IncludeText`)?.checked || false; // 修正為 toolMainContainerEl 內查找
        copyTableToClipboard(includeText);
    }

    function toggleEditMode() {
        StateManager.pushSnapshot('切換編輯模式');
        isEditMode = !isEditMode;
        const editBtn = toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnEditMode`); // 修正為 toolMainContainerEl 內查找
        editBtn.textContent = isEditMode ? '結束編輯' : '編輯模式';
        editBtn.style.backgroundColor = isEditMode ? '#dc3545' : '#007bff';

        const addRowBtn = toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_btnAddRow`); // 修正為 toolMainContainerEl 內查找
        addRowBtn.style.display = isEditMode ? 'inline-block' : 'none';

        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
        renderResultsTableUI(currentData);
        displaySystemNotification(`已${isEditMode ? '進入' : '退出'}編輯模式`, false);
    }

    function handleAddRowToTable() {
        StateManager.pushSnapshot('新增列');
        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;
        const newRow = {
            _isNewRow: true
        }; // 標記為新行以便編輯
        StateManager.currentTable.currentHeaders.forEach(header => {
            newRow[header] = header === FIELD_DISPLAY_NAMES_MAP.NO ? String(currentData.length + 1) : '';
        });
        currentData.push(newRow);
        populateTableRows(currentData);
        displaySystemNotification('已新增列', false);
    }


    function applyTableFilter() {
        const filterInput = toolMainContainerEl?.querySelector('.qt-filter-input'); // 修正為 toolMainContainerEl 內查找
        if (!filterInput) return;

        const searchTerm = filterInput.value.toLowerCase();
        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;

        let filteredData = currentData;
        if (StateManager.a17Mode.isActive && StateManager.a17Mode.selectedUnits.size > 0) {
            filteredData = filteredData.filter(row => { // 在已篩選的數據上再次篩選
                const unit = row[FIELD_DISPLAY_NAMES_MAP.uwApproverUnit] || '';
                return Array.from(StateManager.a17Mode.selectedUnits).some(selectedUnit => {
                    if (selectedUnit === 'UNDEF') {
                        return !Object.keys(UNIT_CODE_MAPPINGS).some(code => unit.includes(code));
                    } else {
                        return unit.includes(selectedUnit);
                    }
                });
            });
        }

        if (searchTerm) {
            filteredData = filteredData.filter(row => {
                return Object.values(row).some(value => {
                    if (value === null || value === undefined) return false;
                    // 對於 statusCombined 欄位，搜尋其內部 HTML 文本
                    if (FIELD_DISPLAY_NAMES_MAP.statusCombined && row[FIELD_DISPLAY_NAMES_MAP.statusCombined] === value) {
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = value;
                        return tempDiv.textContent.toLowerCase().includes(searchTerm);
                    }
                    return String(value).toLowerCase().includes(searchTerm);
                });
            });
        }
        populateTableRows(filteredData);
    }

    function renderA17ModeUI() {
        const a17UnitBtnsCtr = toolMainContainerEl.querySelector('.qt-a17-unit-buttons-container'); // 修正為 toolMainContainerEl 內查找
        const a17TextControls = toolMainContainerEl.querySelector('.qt-a17-text-controls-container'); // 修正為 toolMainContainerEl 內查找
        a17UnitBtnsCtr.style.display = 'flex';
        a17TextControls.style.display = 'flex';
        updateA17UnitButtonCounts(); // 確保按鈕顯示最新計數

        // 綁定A17單位篩選按鈕事件
        a17UnitBtnsCtr.innerHTML = ''; // 清空舊按鈕
        A17_UNIT_BUTTONS_DEFS.forEach(unitDef => {
            const btn = document.createElement('button');
            btn.className = 'qt-unit-filter-btn';
            btn.dataset.unitId = unitDef.id;
            btn.style.backgroundColor = unitDef.color;
            btn.style.color = 'white';
            btn.textContent = `${unitDef.label} (0)`; // 初始為0，稍後更新

            if (StateManager.a17Mode.selectedUnits.has(unitDef.id)) {
                btn.classList.add('active');
            }

            btn.onclick = () => {
                StateManager.pushSnapshot('A17單位篩選');
                if (StateManager.a17Mode.selectedUnits.has(unitDef.id)) {
                    StateManager.a17Mode.selectedUnits.delete(unitDef.id);
                    btn.classList.remove('active');
                } else {
                    StateManager.a17Mode.selectedUnits.add(unitDef.id);
                    btn.classList.add('active');
                }
                applyTableFilter(); // 使用統一的過濾函數
            };
            a17UnitBtnsCtr.appendChild(btn);
        });

        // 綁定A17通知文包含選項和編輯按鈕
        const includeTextCheckbox = a17TextControls.querySelector(`#${TOOL_MAIN_CONTAINER_ID}_cbA17IncludeText`);
        includeTextCheckbox.onchange = () => {
            // 不做快照，這個變化不影響數據
            displaySystemNotification(`A17通知文${includeTextCheckbox.checked ? '將' : '不'}包含在複製內容中`, false, 1500);
        };

        renderTableHeaders([
            FIELD_DISPLAY_NAMES_MAP._queriedValue_,
            FIELD_DISPLAY_NAMES_MAP.NO,
            ...ALL_DISPLAY_FIELDS_API_KEYS_MAIN.map(key => FIELD_DISPLAY_NAMES_MAP[key]),
            ...StateManager.csvImport.selectedColsForA17Merge, // 合併 CSV 欄位
            FIELD_DISPLAY_NAMES_MAP._apiQueryStatus
        ]);
    }

    function renderNormalModeUI() {
        const a17UnitBtnsCtr = toolMainContainerEl.querySelector('.qt-a17-unit-buttons-container'); // 修正為 toolMainContainerEl 內查找
        const a17TextControls = toolMainContainerEl.querySelector('.qt-a17-text-controls-container'); // 修正為 toolMainContainerEl 內查找
        a17UnitBtnsCtr.style.display = 'none';
        a17TextControls.style.display = 'none';
        StateManager.a17Mode.selectedUnits.clear(); // 清空A17模式選中的單位

        renderTableHeaders([
            FIELD_DISPLAY_NAMES_MAP._queriedValue_,
            FIELD_DISPLAY_NAMES_MAP.NO,
            ...ALL_DISPLAY_FIELDS_API_KEYS_MAIN.map(key => FIELD_DISPLAY_NAMES_MAP[key]),
            FIELD_DISPLAY_NAMES_MAP._apiQueryStatus
        ]);
    }


    function updateA17UnitButtonCounts() {
        if (!toolMainContainerEl || !StateManager.currentTable.a17UnitButtonsContainer || !StateManager.a17Mode.isActive) return; // 修正為 toolMainContainerEl 檢查

        const currentData = StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults;

        StateManager.currentTable.a17UnitButtonsContainer.querySelectorAll('.qt-unit-filter-btn').forEach(btn => {
            const unitDefId = btn.dataset.unitId;
            let count = 0;
            if (unitDefId === 'UNDEF') {
                count = currentData.filter(row => {
                    const unit = row[FIELD_DISPLAY_NAMES_MAP.uwApproverUnit] || '';
                    return !Object.keys(UNIT_CODE_MAPPINGS).some(code => unit.includes(code));
                }).length;
            } else {
                count = currentData.filter(row => {
                    const unit = row[FIELD_DISPLAY_NAMES_MAP.uwApproverUnit] || '';
                    return unit.includes(unitDefId);
                }).length;
            }
            btn.textContent = `${A17_UNIT_BUTTONS_DEFS.find(d => d.id === unitDefId)?.label || unitDefId} (${count})`;
            btn.disabled = (count === 0);
            btn.style.opacity = (count === 0) ? '0.5' : '1';
            btn.style.cursor = (count === 0) ? 'not-allowed' : 'pointer';
        });
    }


    function toggleA17Mode(isLongPress = false) {
        if (isLongPress) {
            StateManager.pushSnapshot('強制進入A17模式');
            StateManager.a17Mode.isActive = true;
            StateManager.a17Mode.selectedUnits.clear();
            mergeA17Data(); // 長按直接合併數據
            renderResultsTableUI(StateManager.baseA17MasterData);
            displaySystemNotification('已強制進入A17模式', false);
        } else {
            if (!StateManager.csvImport.isA17CsvPrepared && !StateManager.a17Mode.isActive) { // 如果沒有CSV且不是退出模式
                displaySystemNotification('請先匯入CSV並選擇A17合併欄位', true);
                return;
            }
            StateManager.pushSnapshot('切換A17模式');
            StateManager.a17Mode.isActive = !StateManager.a17Mode.isActive;
            if (StateManager.a17Mode.isActive) {
                mergeA17Data();
            }
            renderResultsTableUI(StateManager.a17Mode.isActive ? StateManager.baseA17MasterData : StateManager.originalQueryResults);
            displaySystemNotification(`已${StateManager.a17Mode.isActive ? '進入' : '退出'}A17模式`, false);
        }
    }


    function mergeA17Data() {
        if (!StateManager.csvImport.isA17CsvPrepared || StateManager.csvImport.selectedColsForA17Merge.length === 0) {
            StateManager.baseA17MasterData = [...StateManager.originalQueryResults];
            return;
        }

        StateManager.baseA17MasterData = StateManager.originalQueryResults.map(queryRow => {
            const mergedRow = {
                ...queryRow
            };
            const queryValue = queryRow[FIELD_DISPLAY_NAMES_MAP._queriedValue_];

            // 尋找對應的CSV資料
            const matchingCsvRow = StateManager.csvImport.rawData.find(csvRowArr => {
                // csvRowArr 是 ['value1', 'value2', ...] 這樣的陣列
                // 匹配邏輯：CSV 的任一列的值與查詢值匹配
                return csvRowArr.some(cellValue => String(cellValue).trim() === String(queryValue).trim());
            });

            // 如果找到匹配的 CSV 行，則將選定的欄位合併到查詢結果中
            if (matchingCsvRow) {
                StateManager.csvImport.selectedColsForA17Merge.forEach(colName => {
                    const colIndex = StateManager.csvImport.rawHeaders.indexOf(colName);
                    mergedRow[colName] = colIndex !== -1 ? (matchingCsvRow[colIndex] || '') : '';
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
                // 處理特殊顯示欄位，確保複製時HTML標籤正確
                if (header === FIELD_DISPLAY_NAMES_MAP.statusCombined) {
                    // 如果是狀態組合，直接使用已經包含HTML的row[header]值
                } else if (header === FIELD_DISPLAY_NAMES_MAP.uwApproverUnit) {
                    const unitCodePrefix = getFirstLetter(cellValue);
                    const mappedUnitName = UNIT_CODE_MAPPINGS[unitCodePrefix] || cellValue;
                    cellValue = unitCodePrefix !== 'UNDEF' && UNIT_CODE_MAPPINGS[unitCodePrefix] ? `${unitCodePrefix}-${mappedUnitName.replace(/^[A-Z]-/,'')}` : mappedUnitName;
                } else if (header === FIELD_DISPLAY_NAMES_MAP.uwApprover || header === FIELD_DISPLAY_NAMES_MAP.approvalUser) {
                    cellValue = extractName(cellValue);
                }

                if (cellValue === null || cellValue === undefined) cellValue = '';
                // 只有在特定欄位才允許HTML，否則進行HTML轉義
                const displayValue = (header === FIELD_DISPLAY_NAMES_MAP.statusCombined) ? String(cellValue) : escapeHtml(String(cellValue));
                content += `<td style="padding:8px;border:1px solid #ddd;">${displayValue}</td>`;
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
        }).catch((err) => {
            console.error("複製到剪貼簿 (HTML) 失敗:", err);
            // 降級處理：複製純文字格式
            const textContent = headers.join('\t') + '\n' +
                currentData.map(row => headers.map(h => {
                    let cellVal = row[h];
                    if (cellVal === null || cellVal === undefined) return '';
                    // 對於包含HTML的狀態欄位，只複製其文本內容
                    if (h === FIELD_DISPLAY_NAMES_MAP.statusCombined) {
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = String(cellVal);
                        return tempDiv.textContent;
                    }
                    return String(cellVal);
                }).join('\t')).join('\n');
            navigator.clipboard.writeText(textContent).then(() => {
                displaySystemNotification('表格已複製到剪貼簿（純文字格式）', false);
            }).catch((err2) => {
                console.error("複製到剪貼簿 (純文字) 失敗:", err2);
                displaySystemNotification('複製失敗，請手動選擇表格內容', true);
            });
        });
    }


    // === 主要執行函數 ===
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
                        ...resultRowBase,
                        ...rec
                    }; // 先合併 API 回傳的原始資料

                    // 針對顯示名稱進行格式化處理
                    ALL_DISPLAY_FIELDS_API_KEYS_MAIN.forEach(dKey => {
                        const displayName = FIELD_DISPLAY_NAMES_MAP[dKey];
                        let cellValue = rec[dKey];
                        if (cellValue === null || cellValue === undefined) cellValue = '';

                        if (dKey === 'statusCombined') {
                            const mainS = rec.mainStatus || '';
                            const subS = rec.subStatus || '';
                            populatedRow[displayName] = `<span style="font-weight:bold;">${escapeHtml(mainS)}</span>` + (subS ? ` <span style="color:#777;">(${escapeHtml(subS)})</span>` : '');
                        } else if (dKey === UNIT_MAP_FIELD_API_KEY) {
                            const unitCodePrefix = getFirstLetter(cellValue);
                            const mappedUnitName = UNIT_CODE_MAPPINGS[unitCodePrefix] || cellValue;
                            populatedRow[displayName] = unitCodePrefix !== 'UNDEF' && UNIT_CODE_MAPPINGS[unitCodePrefix] ? `${unitCodePrefix}-${mappedUnitName.replace(/^[A-Z]-/,'')}` : mappedUnitName;
                        } else if (dKey === 'uwApprover' || dKey === 'approvalUser') {
                            populatedRow[displayName] = extractName(cellValue);
                        } else {
                            populatedRow[displayName] = String(cellValue); // 確保所有值都是字串形式
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
        StateManager.pushSnapshot('初次查詢完成'); // 首次查詢結果也做快照

        renderResultsTableUI(StateManager.originalQueryResults);
        displaySystemNotification(`查詢完成！共處理 ${queryValues.length} 個查詢值`, false, 3500);
    }

    // === 工具啟動主函數 (已強化) ===
    async function executeCaseQueryTool() {
        // 1. 檢查並創建主工具視窗 (如果不存在)
        if (!toolMainContainerEl) {
            toolMainContainerEl = document.getElementById(TOOL_MAIN_CONTAINER_ID);
            if (toolMainContainerEl) {
                toolMainContainerEl.remove(); // 移除舊的，確保乾淨啟動
            }
            toolMainContainerEl = document.createElement('div');
            toolMainContainerEl.id = TOOL_MAIN_CONTAINER_ID;
            toolMainContainerEl.style.cssText = `
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
            border:1px solid #dee2e6; /* 主視窗邊框 */
            box-sizing: border-box; /* 確保padding不會撐大尺寸 */
        `;
            document.body.appendChild(toolMainContainerEl); // 將主容器附加到 body
        }

        // 2. 檢查是否已有工具開啟（通過判斷其內容是否已填充）
        // 如果工具視窗已存在且有內容（例如上次查詢結果），則不重複啟動
        if (toolMainContainerEl.children.length > 0 && toolMainContainerEl.querySelector('.qt-main-title-bar')) {
            displaySystemNotification('工具已開啟', true);
            return;
        }

        // 3. 自動判斷並設定環境（優先）
        EnvManager.autoDetectAndSet();
        // EnvManager.updateDisplay() 會在 EnvManager.set() 中被呼叫，確保在主視窗內顯示

        // 4. 清除可能殘留的舊對話框 overlay
        ['_EnvSelect_overlay', '_Token_overlay', '_QuerySetup_overlay',
            '_A17TextSettings_overlay', '_Loading_overlay', '_CSVPurpose_overlay',
            '_CSVColSelect_overlay', '_CSVCheckbox_overlay', '_RetryEdit_overlay'
        ]
        .forEach(suffix => {
            // 在 toolMainContainerEl 內查找並移除
            const el = toolMainContainerEl.querySelector(`#${TOOL_MAIN_CONTAINER_ID}${suffix}`);
            if (el) el.remove();
        });
        // 移除系統通知（如果存在於 body 層級）
        document.getElementById(TOOL_MAIN_CONTAINER_ID + '_Notification')?.remove();

        // 5. 處理 Token 邏輯
        if (!TokenManager.init()) {
            let tokenAttempt = 1;
            while (true) {
                const tokenResult = await TokenManager.showDialog(tokenAttempt);
                if (tokenResult === '_close_tool_') {
                    displaySystemNotification('工具已關閉', false);
                    if (toolMainContainerEl) toolMainContainerEl.remove(); // 關閉整個工具主視窗
                    return;
                }
                if (tokenResult === '_skip_token_') {
                    TokenManager.clear();
                    displaySystemNotification('已略過Token輸入', false);
                    break;
                }
                if (tokenResult === '_token_dialog_cancel_') {
                    displaySystemNotification('Token輸入已取消', true);
                    if (toolMainContainerEl) toolMainContainerEl.remove(); // 關閉整個工具主視窗
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

        // 6. 顯示查詢設定對話框
        const querySetupResult = await createQuerySetupDialog();
        if (!querySetupResult) {
            displaySystemNotification('操作已取消', true);
            if (toolMainContainerEl) toolMainContainerEl.remove(); // 關閉整個工具主視窗
            return;
        }

        await executeQuery(querySetupResult);
    }

    // === 主函數執行入口 ===
    (async function main() {
        // 主函數現在只負責調用 executeCaseQueryTool，所有的 UI 管理都在 executeCaseQueryTool 內部
        await executeCaseQueryTool();
    })();
})();
