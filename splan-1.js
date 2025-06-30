＊＊＊＊＊第38段程式碼開始/17商品 /＊＊＊＊＊
javascript:(async()=>{const API_URL_TEMPLATES_PQ={uat:{planCodeQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'},prod:{planCodeQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'}};const TOKEN_STORAGE_KEY_PQ='euisPlanCodeToken_v2';const TOKEN_STORAGE_KEY_PQ_FALLBACK='euisToken';const MAIN_UI_ID_PQ='productQueryBookmarkletDiv_Structured_v2';const CSS_ID_PQ='productQueryBookmarkletCss_Structured_v2';const Z_INDEX_OVERLAY_PQ=2147483640;const Z_INDEX_DIALOG_PQ=2147483641;const Z_INDEX_MAIN_UI_PQ=2147483630;const Z_INDEX_NOTIFICATION_PQ=2147483647;const PRODUCT_FIELD_DEFINITIONS_PQ={CUR:{1:
＊＊＊＊＊第38段程式碼結束/17商品 /＊＊＊＊＊




＊＊＊＊＊第39段程式碼開始/18商品多 /＊＊＊＊＊
javascript:(async function(){const API_URL_TEMPLATES_PQ={uat:{planCodeQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'},prod:{planCodeQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'}};const TOKEN_STORAGE_KEY_PQ='euisPlanCodeToken_v2';const TOKEN_STORAGE_KEY_PQ_FALLBACK='euisToken';const MAIN_UI_ID_PQ='productQueryBookmarkletDiv_Structured_v2';const CSS_ID_PQ='productQueryBookmarkletCss_Structured_v2';const Z_INDEX_OVERLAY_PQ=2147483640;const Z_INDEX_DIALOG_PQ=2147483641;const Z_INDEX_MAIN_UI_PQ=2147483630;const Z_INDEX_NOTIFICATION_PQ=2147483647;const PRODUCT_FIELD_DEFINITIONS_PQ={CUR:{1:
＊＊＊＊＊第39段程式碼結束/18商品多 /＊＊＊＊＊




＊＊＊＊＊第40段程式碼開始/19商品 /＊＊＊＊＊
javascript:(async function(){let currentPage=1;let totalPages=1;let allDataArray=[];const recordsPerPage=100;let currentToken;let displayedDataArray=[];const API_URL_PLANCODE_QUERY=
＊＊＊＊＊第40段程式碼結束/19商品 /＊＊＊＊＊




＊＊＊＊＊第41段程式碼開始/21商品 /＊＊＊＊＊
javascript:(async function(){    'use strict';          const state={        currentPage:1,totalPages:1,allDataArray:[],recordsPerPage:100,currentToken:null,        displayedDataArray:[],isShowingAllData:!1,queriedDataCache:new Map(),        currentSortColumn:null,currentSortDirection:'asc',apiAbortController:null,        currentFilters:{}   };          const API_URLS={        PLANCODE_QUERY:
＊＊＊＊＊第41段程式碼結束/21商品 /＊＊＊＊＊




＊＊＊＊＊第42段程式碼開始/55PLAN_4 /＊＊＊＊＊
javascript:(async()=>{console.log('[%E5%95%86%E5%93%81%E6%9F%A5%E8%A9%A2%E8%85%B3%E6%9C%AC] %E8%85%B3%E6%9C%AC%E9%96%8B%E5%A7%8B%E5%9F%B7%E8%A1%8C R3-debug...');const API_URL_TEMPLATES_PQ={uat:{planCodeQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'},prod:{planCodeQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'}};const TOKEN_STORAGE_KEY_PQ='euisPlanCodeToken_v2';const TOKEN_STORAGE_KEY_PQ_FALLBACK='euisToken';const MAIN_UI_ID_PQ='productQueryBookmarkletDiv_Structured_v2';const CSS_ID_PQ='productQueryBookmarkletCss_Structured_v2';const Z_INDEX_OVERLAY_PQ=2147483640;const Z_INDEX_DIALOG_PQ=2147483641;const Z_INDEX_MAIN_UI_PQ=2147483630;const Z_INDEX_NOTIFICATION_PQ=2147483647;const PRODUCT_FIELD_DEFINITIONS_PQ={CUR:{1:
＊＊＊＊＊第42段程式碼結束/55PLAN_4 /＊＊＊＊＊




＊＊＊＊＊第43段程式碼開始/56PLAN_5 /＊＊＊＊＊
javascript:(async function(){%C2%A0 %C2%A0 'use strict';%C2%A0 %C2%A0%C2%A0 %C2%A0%C2%A0 %C2%A0 const state={%C2%A0 %C2%A0 %C2%A0 %C2%A0 currentPage:1,totalPages:1,allDataArray:[],recordsPerPage:100,currentToken:null,%C2%A0 %C2%A0 %C2%A0 %C2%A0 displayedDataArray:[],isShowingAllData:!1,queriedDataCache:new Map(),%C2%A0 %C2%A0 %C2%A0 %C2%A0 currentSortColumn:null,currentSortDirection:'asc',apiAbortController:null,%C2%A0 %C2%A0 %C2%A0 %C2%A0 currentFilters:{},errorMessage:null%C2%A0 %C2%A0};%C2%A0 %C2%A0%C2%A0 %C2%A0%C2%A0 %C2%A0 const API_URLS={%C2%A0 %C2%A0 %C2%A0 %C2%A0 PLANCODE_QUERY:
＊＊＊＊＊第43段程式碼結束/56PLAN_5 /＊＊＊＊＊




＊＊＊＊＊第44段程式碼開始/60perlesty /＊＊＊＊＊
javascript:(function(){ 'use strict';  /**  * KGI Life %E5%95%86%E5%93%81%E6%9F%A5%E8%A9%A2%E5%B7%A5%E5%85%B7 - 2025 %E6%9C%80%E4%BD%B3%E5%8C%96%E7%89%88  * by Perplexity AI (for %E5%89%8D%E7%AB%AF%E4%B8%BB%E7%AE%A1/%E6%9E%B6%E6%A7%8B%E5%B8%AB/%E8%B3%87%E6%B7%B1 UI/UX %E8%A8%AD%E8%A8%88%E5%B8%AB)  */  /* ==== CONFIG ==== */ const BATCH_SIZE = 50; const CONFIG = {   API: {     prod: {       planCodeQuery:   'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',       planCodeDetail:  'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',       saleDateQuery:   'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'     },     uat: {       planCodeQuery:   'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',       planCodeDetail:  'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',       saleDateQuery:   'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'     }   },   STORAGE: {     TOKEN_KEY:
＊＊＊＊＊第44段程式碼結束/60perlesty /＊＊＊＊＊




＊＊＊＊＊第45段程式碼開始/61gemibi /＊＊＊＊＊
javascript:(function(){'use strict';const BATCH_SIZE=50;const CONFIG={API:{prod:{planCodeQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv.apps.ocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'},uat:{planCodeQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/query',planCodeDetail:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeController/queryDetail',saleDateQuery:'https://euisv-uat.apps.tocp4.kgilife.com.tw/euisw/euisbq/api/planCodeSaleDateController/query'}},STORAGE:{TOKEN_KEY:
＊＊＊＊＊第45段程式碼結束/61gemibi /＊＊＊＊＊

