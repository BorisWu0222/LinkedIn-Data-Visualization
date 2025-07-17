<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LinkedIn 數據視覺化儀表板</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background-color: #f9fafb;
            color: #111827;
        }
        
        .min-h-screen {
            min-height: 100vh;
        }
        
        .bg-gray-50 {
            background-color: #f9fafb;
        }
        
        .p-6 {
            padding: 1.5rem;
        }
        
        .max-w-7xl {
            max-width: 80rem;
        }
        
        .mx-auto {
            margin-left: auto;
            margin-right: auto;
        }
        
        .text-center {
            text-align: center;
        }
        
        .mb-8 {
            margin-bottom: 2rem;
        }
        
        .mb-4 {
            margin-bottom: 1rem;
        }
        
        .mb-2 {
            margin-bottom: 0.5rem;
        }
        
        .mb-3 {
            margin-bottom: 0.75rem;
        }
        
        .mt-2 {
            margin-top: 0.5rem;
        }
        
        .ml-4 {
            margin-left: 1rem;
        }
        
        .text-3xl {
            font-size: 1.875rem;
            line-height: 2.25rem;
        }
        
        .text-2xl {
            font-size: 1.5rem;
            line-height: 2rem;
        }
        
        .text-xl {
            font-size: 1.25rem;
            line-height: 1.75rem;
        }
        
        .text-lg {
            font-size: 1.125rem;
            line-height: 1.75rem;
        }
        
        .text-sm {
            font-size: 0.875rem;
            line-height: 1.25rem;
        }
        
        .text-xs {
            font-size: 0.75rem;
            line-height: 1rem;
        }
        
        .font-bold {
            font-weight: 700;
        }
        
        .font-semibold {
            font-weight: 600;
        }
        
        .font-medium {
            font-weight: 500;
        }
        
        .text-gray-900 {
            color: #111827;
        }
        
        .text-gray-700 {
            color: #374151;
        }
        
        .text-gray-600 {
            color: #4b5563;
        }
        
        .text-gray-500 {
            color: #6b7280;
        }
        
        .text-gray-400 {
            color: #9ca3af;
        }
        
        .text-blue-900 {
            color: #1e3a8a;
        }
        
        .text-blue-800 {
            color: #1e40af;
        }
        
        .text-white {
            color: #ffffff;
        }
        
        .text-red-700 {
            color: #b91c1c;
        }
        
        .text-purple-600 {
            color: #9333ea;
        }
        
        .text-blue-600 {
            color: #2563eb;
        }
        
        .text-green-600 {
            color: #16a34a;
        }
        
        .text-orange-600 {
            color: #ea580c;
        }
        
        .bg-white {
            background-color: #ffffff;
        }
        
        .bg-blue-50 {
            background-color: #eff6ff;
        }
        
        .bg-blue-600 {
            background-color: #2563eb;
        }
        
        .bg-blue-700 {
            background-color: #1d4ed8;
        }
        
        .bg-gray-200 {
            background-color: #e5e7eb;
        }
        
        .bg-gray-300 {
            background-color: #d1d5db;
        }
        
        .bg-red-100 {
            background-color: #fee2e2;
        }
        
        .rounded-lg {
            border-radius: 0.5rem;
        }
        
        .rounded-md {
            border-radius: 0.375rem;
        }
        
        .rounded {
            border-radius: 0.25rem;
        }
        
        .shadow-md {
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        }
        
        .shadow-lg {
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
        }
        
        .p-8 {
            padding: 2rem;
        }
        
        .p-4 {
            padding: 1rem;
        }
        
        .p-3 {
            padding: 0.75rem;
        }
        
        .px-4 {
            padding-left: 1rem;
            padding-right: 1rem;
        }
        
        .px-3 {
            padding-left: 0.75rem;
            padding-right: 0.75rem;
        }
        
        .py-2 {
            padding-top: 0.5rem;
            padding-bottom: 0.5rem;
        }
        
        .py-1 {
            padding-top: 0.25rem;
            padding-bottom: 0.25rem;
        }
        
        .py-3 {
            padding-top: 0.75rem;
            padding-bottom: 0.75rem;
        }
        
        .py-8 {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        
        .border {
            border-width: 1px;
        }
        
        .border-2 {
            border-width: 2px;
        }
        
        .border-dashed {
            border-style: dashed;
        }
        
        .border-gray-300 {
            border-color: #d1d5db;
        }
        
        .border-red-400 {
            border-color: #f87171;
        }
        
        .grid {
            display: grid;
        }
        
        .grid-cols-1 {
            grid-template-columns: repeat(1, minmax(0, 1fr));
        }
        
        .gap-4 {
            gap: 1rem;
        }
        
        .flex {
            display: flex;
        }
        
        .items-center {
            align-items: center;
        }
        
        .justify-center {
            justify-content: center;
        }
        
        .justify-between {
            justify-content: space-between;
        }
        
        .space-y-1 > * + * {
            margin-top: 0.25rem;
        }
        
        .instruction-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        
        .instruction-list li {
            margin-bottom: 0.25rem;
        }
        
        .space-x-2 > * + * {
            margin-left: 0.5rem;
        }
        
        .cursor-pointer {
            cursor: pointer;
        }
        
        .hidden {
            display: none;
        }
        
        .inline-block {
            display: inline-block;
        }
        
        .h-12 {
            height: 3rem;
        }
        
        .h-8 {
            height: 2rem;
        }
        
        .w-12 {
            width: 3rem;
        }
        
        .w-8 {
            width: 2rem;
        }
        
        .max-w-xs {
            max-width: 20rem;
        }
        
        .transition-colors {
            transition: background-color 0.15s ease-in-out;
        }
        
        .hover\:bg-blue-700:hover {
            background-color: #1d4ed8;
        }
        
        .hover\:bg-gray-300:hover {
            background-color: #d1d5db;
        }
        
        .animate-spin {
            animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
            from {
                transform: rotate(0deg);
            }
            to {
                transform: rotate(360deg);
            }
        }
        
        .border-b-2 {
            border-bottom-width: 2px;
        }
        
        .rounded-full {
            border-radius: 9999px;
        }
        
        .border-blue-600 {
            border-color: #2563eb;
        }
        
        @media (min-width: 768px) {
            .md\:grid-cols-4 {
                grid-template-columns: repeat(4, minmax(0, 1fr));
            }
        }
        
        button {
            border: none;
            cursor: pointer;
            outline: none;
            font-family: inherit;
        }
        
        button:focus {
            outline: 2px solid #3b82f6;
            outline-offset: 2px;
        }
        
        .chart-container {
            width: 100%;
            height: 400px;
            position: relative;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 1rem;
        }
        
        .data-table th,
        .data-table td {
            padding: 0.75rem;
            text-align: left;
            border-bottom: 1px solid #e5e7eb;
        }
        
        .data-table th {
            background-color: #f9fafb;
            font-weight: 600;
            color: #374151;
        }
        
        .data-table tr:hover {
            background-color: #f9fafb;
        }

        .data-table a {
            color: #2563eb;
            text-decoration: underline;
            cursor: pointer;
        }

        .data-table a:hover {
            color: #1d4ed8;
            text-decoration: underline;
        }

        .data-table a:visited {
            color: #7c3aed;
        }

        .tooltip {
            position: absolute;
            background: white;
            border: 1px solid #d1d5db;
            border-radius: 0.375rem;
            padding: 0.75rem;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            display: none;
            max-width: 20rem;
        }

        .overflow-x-auto {
            overflow-x: auto;
        }

        .pagination {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 1rem;
            margin-top: 1rem;
        }

        .pagination button {
            padding: 0.5rem 1rem;
            border: 1px solid #d1d5db;
            background-color: white;
            color: #374151;
            border-radius: 0.375rem;
            cursor: pointer;
            transition: all 0.15s ease-in-out;
        }

        .pagination button:hover:not(:disabled) {
            background-color: #f3f4f6;
            border-color: #9ca3af;
        }

        .pagination button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .pagination .page-info {
            font-size: 0.875rem;
            color: #6b7280;
        }
    </style>
</head>
<body>
    <div class="min-h-screen bg-gray-50 p-6">
        <div class="max-w-7xl mx-auto">
            <!-- 標題 -->
            <div class="text-center mb-8">
                <h1 class="text-3xl font-bold text-gray-900 mb-2">LinkedIn 數據視覺化儀表板</h1>
                <p class="text-gray-600">上傳LinkedIn匯出的Excel檔案進行數據分析</p>
            </div>

            <!-- 使用說明 -->
            <div class="bg-blue-50 rounded-lg p-6 mb-8">
                <h3 class="text-lg font-semibold text-blue-900 mb-3">使用說明</h3>
                <ul class="text-blue-800 instruction-list">
                    <li>1. 從LinkedIn匯出包含"Metrics"和"All posts"兩個工作表的Excel檔案</li>
                    <li>2. 點擊下方的"選擇Excel檔案"按鈕上傳檔案</li>
                    <li>3. 系統會自動分析數據並生成視覺化圖表</li>
                    <li>4. 可以將滑鼠懸停在圖表上查看詳細資訊</li>
                </ul>
            </div>

            <!-- 檔案上傳區域 -->
            <div class="bg-white rounded-lg shadow-md p-6 mb-8">
                <div class="flex items-center justify-center border-2 border-dashed border-gray-300 rounded-lg p-8">
                    <div class="text-center">
                        <svg class="mx-auto h-12 w-12 text-gray-400 mb-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
                        </svg>
                        <label class="cursor-pointer">
                            <span class="bg-blue-600 text-white px-4 py-2 rounded-md hover:bg-blue-700 transition-colors">
                                選擇Excel檔案
                            </span>
                            <input type="file" accept=".xlsx,.xls" id="fileInput" class="hidden" />
                        </label>
                        <p class="text-sm text-gray-500 mt-2">支援 .xlsx 和 .xls 格式</p>
                    </div>
                </div>
            </div>

            <!-- 載入狀態 -->
            <div id="loadingSection" class="text-center py-8 hidden">
                <div class="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                <p class="mt-2 text-gray-600">正在處理數據...</p>
            </div>

            <!-- 錯誤訊息 -->
            <div id="errorSection" class="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4 hidden">
                <span id="errorMessage"></span>
            </div>

            <!-- 數據統計卡片 -->
            <div id="statsSection" class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8 hidden">
                <div class="bg-white p-6 rounded-lg shadow-md">
                    <div class="flex items-center">
                        <svg class="h-8 w-8 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                        </svg>
                        <div class="ml-4">
                            <p class="text-sm font-medium text-gray-600">貼文總數</p>
                            <p class="text-2xl font-bold text-gray-900" id="totalPosts">0</p>
                        </div>
                    </div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-md">
                    <div class="flex items-center">
                        <svg class="h-8 w-8 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path>
                        </svg>
                        <div class="ml-4">
                            <p class="text-sm font-medium text-gray-600">總曝光數</p>
                            <p class="text-2xl font-bold text-gray-900" id="totalImpressions">0</p>
                        </div>
                    </div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-md">
                    <div class="flex items-center">
                        <svg class="h-8 w-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"></path>
                        </svg>
                        <div class="ml-4">
                            <p class="text-sm font-medium text-gray-600">不重複曝光數</p>
                            <p class="text-2xl font-bold text-gray-900" id="totalUniqueImpressions">0</p>
                        </div>
                    </div>
                </div>
                <div class="bg-white p-6 rounded-lg shadow-md">
                    <div class="flex items-center">
                        <svg class="h-8 w-8 text-orange-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path>
                        </svg>
                        <div class="ml-4">
                            <p class="text-sm font-medium text-gray-600">平均點閱率</p>
                            <p class="text-2xl font-bold text-gray-900" id="avgCTR">0%</p>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Metrics 圖表 -->
            <div id="metricsChartSection" class="bg-white rounded-lg shadow-md p-6 mb-8 hidden">
                <div class="flex justify-between items-center mb-4">
                    <h2 class="text-xl font-bold text-gray-900">每日曝光數與點閱率趨勢分析</h2>
                    <div class="flex space-x-2">
                        <button id="dailyBtn" class="px-3 py-1 rounded-md text-sm bg-gray-200 text-gray-700 hover:bg-gray-300">日</button>
                        <button id="monthlyBtn" class="px-3 py-1 rounded-md text-sm bg-blue-600 text-white">月</button>
                        <button id="quarterlyBtn" class="px-3 py-1 rounded-md text-sm bg-gray-200 text-gray-700 hover:bg-gray-300">季</button>
                    </div>
                </div>
                <div class="chart-container">
                    <canvas id="metricsChart"></canvas>
                </div>
            </div>

            <!-- All Posts 圖表 -->
            <div id="postsChartSection" class="bg-white rounded-lg shadow-md p-6 mb-8 hidden">
                <h2 class="text-xl font-bold text-gray-900 mb-4">每篇貼文趨勢分析</h2>
                <div class="chart-container">
                    <canvas id="postsChart"></canvas>
                </div>
            </div>

            <!-- 數據表格 -->
            <div id="metricsTableSection" class="bg-white rounded-lg shadow-md p-6 mb-8 hidden">
                <h2 class="text-xl font-bold text-gray-900 mb-4">每日詳細數據</h2>
                <div class="overflow-x-auto">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>日期</th>
                                <th>總曝光數</th>
                                <th>不重複曝光數</th>
                                <th>重複曝光數</th>
                                <th>點閱率 (%)</th>
                            </tr>
                        </thead>
                        <tbody id="metricsTableBody">
                        </tbody>
                    </table>
                    <p id="metricsTableInfo" class="text-sm text-gray-500 mt-2"></p>
                    <div id="metricsPagination" class="pagination hidden">
                        <button id="metricsPrevBtn" onclick="changeMetricsPage(-1)">上一頁</button>
                        <span id="metricsPageInfo" class="page-info"></span>
                        <button id="metricsNextBtn" onclick="changeMetricsPage(1)">下一頁</button>
                    </div>
                </div>
            </div>

            <!-- 貼文數據表格 -->
            <div id="postsTableSection" class="bg-white rounded-lg shadow-md p-6 hidden">
                <h2 class="text-xl font-bold text-gray-900 mb-4">每篇貼文詳細數據</h2>
                <div class="overflow-x-auto">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>日期</th>
                                <th>標題</th>
                                <th>曝光數</th>
                                <th>點閱率 (%)</th>
                            </tr>
                        </thead>
                        <tbody id="postsTableBody">
                        </tbody>
                    </table>
                    <p id="postsTableInfo" class="text-sm text-gray-500 mt-2"></p>
                    <div id="postsPagination" class="pagination hidden">
                        <button id="postsPrevBtn" onclick="changePostsPage(-1)">上一頁</button>
                        <span id="postsPageInfo" class="page-info"></span>
                        <button id="postsNextBtn" onclick="changePostsPage(1)">下一頁</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        // 全局變量
        let metricsData = [];
        let postsData = [];
        let loading = false;
        let error = null;
        let timeGrouping = 'monthly';
        let metricsChart = null;
        let postsChart = null;
        let metricsCurrentPage = 0;
        let postsCurrentPage = 0;
        const itemsPerPage = 10;

        // 工具函數
        function roundToTwo(num) {
            return Math.round(num * 100) / 100;
        }

        // 時間聚合函數
        function groupDataByTime(data, grouping) {
            if (grouping === 'daily') return data;
            
            const grouped = {};
            
            data.forEach(item => {
                const date = new Date(item.date);
                let key;
                
                if (grouping === 'monthly') {
                    key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                } else if (grouping === 'quarterly') {
                    const quarter = Math.floor(date.getMonth() / 3) + 1;
                    key = `${date.getFullYear()}-Q${quarter}`;
                }
                
                if (!grouped[key]) {
                    grouped[key] = {
                        date: key,
                        uniqueImpressions: 0,
                        repeatImpressions: 0,
                        totalImpressions: 0,
                        clickThroughRate: 0,
                        count: 0
                    };
                }
                
                grouped[key].uniqueImpressions += item.uniqueImpressions;
                grouped[key].repeatImpressions += item.repeatImpressions;
                grouped[key].totalImpressions += item.totalImpressions;
                grouped[key].clickThroughRate += item.clickThroughRate;
                grouped[key].count += 1;
            });
            
            return Object.values(grouped).map(item => ({
                ...item,
                clickThroughRate: roundToTwo(item.clickThroughRate / item.count)
            })).sort((a, b) => a.date.localeCompare(b.date));
        }

        // 顯示載入狀態
        function showLoading() {
            loading = true;
            document.getElementById('loadingSection').classList.remove('hidden');
            document.getElementById('errorSection').classList.add('hidden');
        }

        // 隱藏載入狀態
        function hideLoading() {
            loading = false;
            document.getElementById('loadingSection').classList.add('hidden');
        }

        // 顯示錯誤
        function showError(message) {
            error = message;
            document.getElementById('errorMessage').textContent = message;
            document.getElementById('errorSection').classList.remove('hidden');
            hideLoading();
        }

        // 隱藏錯誤
        function hideError() {
            error = null;
            document.getElementById('errorSection').classList.add('hidden');
        }

        // 更新統計數據
        function updateStats() {
            const totalPosts = postsData.length;
            const totalImpressions = metricsData.reduce((sum, item) => sum + item.totalImpressions, 0);
            const totalUniqueImpressions = metricsData.reduce((sum, item) => sum + item.uniqueImpressions, 0);
            const avgCTR = postsData.length > 0 ? roundToTwo(postsData.reduce((sum, item) => sum + item.ctr, 0) / postsData.length) : 0;

            document.getElementById('totalPosts').textContent = totalPosts;
            document.getElementById('totalImpressions').textContent = totalImpressions.toLocaleString();
            document.getElementById('totalUniqueImpressions').textContent = totalUniqueImpressions.toLocaleString();
            document.getElementById('avgCTR').textContent = avgCTR + '%';

            // 顯示統計區塊
            if (metricsData.length > 0 || postsData.length > 0) {
                document.getElementById('statsSection').classList.remove('hidden');
            }
        }

        // 更新表格
        function updateTables() {
            updateMetricsTable();
            updatePostsTable();
        }

        // 更新metrics表格
        function updateMetricsTable() {
            if (metricsData.length === 0) return;

            const metricsTableBody = document.getElementById('metricsTableBody');
            metricsTableBody.innerHTML = '';
            
            // 按日期降序排列，顯示最新的數據
            const sortedMetricsData = [...metricsData].sort((a, b) => new Date(b.date) - new Date(a.date));
            
            // 分頁邏輯
            const startIndex = metricsCurrentPage * itemsPerPage;
            const endIndex = startIndex + itemsPerPage;
            const pageData = sortedMetricsData.slice(startIndex, endIndex);
            
            pageData.forEach(item => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${item.date}</td>
                    <td>${item.totalImpressions.toLocaleString()}</td>
                    <td>${item.uniqueImpressions.toLocaleString()}</td>
                    <td>${item.repeatImpressions.toLocaleString()}</td>
                    <td>${item.clickThroughRate}%</td>
                `;
                metricsTableBody.appendChild(row);
            });

            // 更新分頁信息
            const totalPages = Math.ceil(sortedMetricsData.length / itemsPerPage);
            document.getElementById('metricsPageInfo').textContent = 
                `第 ${metricsCurrentPage + 1} 頁，共 ${totalPages} 頁 (總共 ${sortedMetricsData.length} 筆)`;
            
            // 更新按鈕狀態
            document.getElementById('metricsPrevBtn').disabled = metricsCurrentPage === 0;
            document.getElementById('metricsNextBtn').disabled = metricsCurrentPage >= totalPages - 1;
            
            // 顯示分頁控制
            if (totalPages > 1) {
                document.getElementById('metricsPagination').classList.remove('hidden');
            } else {
                document.getElementById('metricsPagination').classList.add('hidden');
            }
            
            document.getElementById('metricsTableSection').classList.remove('hidden');
        }

        // 更新posts表格
        function updatePostsTable() {
            if (postsData.length === 0) return;

            const postsTableBody = document.getElementById('postsTableBody');
            postsTableBody.innerHTML = '';
            
            // 按日期降序排列，顯示最新的數據
            const sortedPostsData = [...postsData].sort((a, b) => new Date(b.date) - new Date(a.date));
            
            // 分頁邏輯
            const startIndex = postsCurrentPage * itemsPerPage;
            const endIndex = startIndex + itemsPerPage;
            const pageData = sortedPostsData.slice(startIndex, endIndex);
            
            pageData.forEach(item => {
                const row = document.createElement('tr');
                // 為標題添加超連結
                const titleWithLink = item.link ? 
                    `<a href="${item.link}" target="_blank" rel="noopener noreferrer" style="color: #2563eb; text-decoration: underline; cursor: pointer;">${item.title}</a>` : 
                    item.title;
                
                row.innerHTML = `
                    <td>${item.date}</td>
                    <td>${titleWithLink}</td>
                    <td>${item.impressions.toLocaleString()}</td>
                    <td>${item.ctr}%</td>
                `;
                postsTableBody.appendChild(row);
            });

            // 更新分頁信息
            const totalPages = Math.ceil(sortedPostsData.length / itemsPerPage);
            document.getElementById('postsPageInfo').textContent = 
                `第 ${postsCurrentPage + 1} 頁，共 ${totalPages} 頁 (總共 ${sortedPostsData.length} 筆)`;
            
            // 更新按鈕狀態
            document.getElementById('postsPrevBtn').disabled = postsCurrentPage === 0;
            document.getElementById('postsNextBtn').disabled = postsCurrentPage >= totalPages - 1;
            
            // 顯示分頁控制
            if (totalPages > 1) {
                document.getElementById('postsPagination').classList.remove('hidden');
            } else {
                document.getElementById('postsPagination').classList.add('hidden');
            }
            
            document.getElementById('postsTableSection').classList.remove('hidden');
        }

        // 更改metrics頁面
        function changeMetricsPage(direction) {
            const sortedMetricsData = [...metricsData].sort((a, b) => new Date(b.date) - new Date(a.date));
            const totalPages = Math.ceil(sortedMetricsData.length / itemsPerPage);
            
            metricsCurrentPage += direction;
            if (metricsCurrentPage < 0) metricsCurrentPage = 0;
            if (metricsCurrentPage >= totalPages) metricsCurrentPage = totalPages - 1;
            
            updateMetricsTable();
        }

        // 更改posts頁面
        function changePostsPage(direction) {
            const sortedPostsData = [...postsData].sort((a, b) => new Date(b.date) - new Date(a.date));
            const totalPages = Math.ceil(sortedPostsData.length / itemsPerPage);
            
            postsCurrentPage += direction;
            if (postsCurrentPage < 0) postsCurrentPage = 0;
            if (postsCurrentPage >= totalPages) postsCurrentPage = totalPages - 1;
            
            updatePostsTable();
        }

        // 創建圖表
        function createMetricsChart() {
            if (metricsData.length === 0) return;

            const canvas = document.getElementById('metricsChart');
            const ctx = canvas.getContext('2d');
            
            if (metricsChart) {
                metricsChart.destroy();
            }

            const groupedData = groupDataByTime(metricsData, timeGrouping);

            metricsChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: groupedData.map(item => item.date),
                    datasets: [{
                        label: '點閱率 (%)',
                        data: groupedData.map(item => item.clickThroughRate),
                        type: 'line',
                        backgroundColor: 'rgba(239, 68, 68, 1)',
                        borderColor: 'rgba(239, 68, 68, 1)',
                        borderWidth: 2,
                        fill: false,
                        yAxisID: 'y1',
                        tension: 0.1,
                        pointRadius: 4,
                        pointBackgroundColor: 'rgba(239, 68, 68, 1)',
                        pointBorderColor: 'rgba(239, 68, 68, 1)',
                        pointBorderWidth: 2,
                        order: 1  // 確保線條在最前面
                    }, {
                        label: '不重複曝光',
                        data: groupedData.map(item => item.uniqueImpressions),
                        backgroundColor: 'rgba(59, 130, 246, 0.8)',
                        borderColor: 'rgba(59, 130, 246, 1)',
                        borderWidth: 1,
                        stack: 'impressions',
                        order: 2
                    }, {
                        label: '重複曝光',
                        data: groupedData.map(item => item.repeatImpressions),
                        backgroundColor: 'rgba(147, 197, 253, 0.8)',
                        borderColor: 'rgba(147, 197, 253, 1)',
                        borderWidth: 1,
                        stack: 'impressions',
                        order: 3
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {
                        mode: 'index',
                        intersect: false,
                    },
                    scales: {
                        x: {
                            display: true,
                            title: {
                                display: true,
                                text: timeGrouping === 'daily' ? '日期' : timeGrouping === 'monthly' ? '月份' : '季度'
                            }
                        },
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: '曝光數'
                            },
                            beginAtZero: true
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: '點閱率 (%)'
                            },
                            beginAtZero: true,
                            grid: {
                                drawOnChartArea: false,
                            },
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    let label = context.dataset.label || '';
                                    if (label) {
                                        label += ': ';
                                    }
                                    if (context.dataset.label === '點閱率 (%)') {
                                        label += context.parsed.y + '%';
                                    } else {
                                        label += context.parsed.y.toLocaleString();
                                    }
                                    return label;
                                }
                            }
                        }
                    }
                }
            });

            document.getElementById('metricsChartSection').classList.remove('hidden');
        }

        function createPostsChart() {
            if (postsData.length === 0) return;

            const canvas = document.getElementById('postsChart');
            const ctx = canvas.getContext('2d');
            
            if (postsChart) {
                postsChart.destroy();
            }

            postsChart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: postsData.map(item => item.date),
                    datasets: [{
                        label: '點閱率 (%)',
                        data: postsData.map(item => item.ctr),
                        borderColor: 'rgba(239, 68, 68, 1)',
                        backgroundColor: 'rgba(239, 68, 68, 1)',
                        borderWidth: 2,
                        fill: false,
                        tension: 0.1,
                        yAxisID: 'y1',
                        pointRadius: 4,
                        pointBackgroundColor: 'rgba(239, 68, 68, 1)',
                        pointBorderColor: 'rgba(239, 68, 68, 1)',
                        pointBorderWidth: 2,
                        order: 1  // 確保紅色線條在最前面
                    }, {
                        label: '曝光數',
                        data: postsData.map(item => item.impressions),
                        borderColor: 'rgba(16, 185, 129, 1)',
                        backgroundColor: 'rgba(16, 185, 129, 1)',
                        borderWidth: 2,
                        fill: false,
                        tension: 0.1,
                        order: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {
                        mode: 'index',
                        intersect: false,
                    },
                    scales: {
                        x: {
                            display: true,
                            title: {
                                display: true,
                                text: '日期'
                            }
                        },
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: '曝光數'
                            },
                            beginAtZero: true
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: '點閱率 (%)'
                            },
                            beginAtZero: true,
                            grid: {
                                drawOnChartArea: false,
                            },
                        }
                    },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'top'
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    let label = context.dataset.label || '';
                                    if (label) {
                                        label += ': ';
                                    }
                                    if (context.dataset.label === '點閱率 (%)') {
                                        label += context.parsed.y + '%';
                                    } 
                                    
                                    else {
                                        label += context.parsed.y.toLocaleString();
                                    }
                                    return label;
                                },
                                afterLabel: function(context) {
                                    if (context.datasetIndex === 0) {
                                        const data = postsData[context.dataIndex];
                                        return [`標題: ${data.title}`];
                                    }
                                    return null;
                                }
                            }
                        }
                    }
                }
            });

            document.getElementById('postsChartSection').classList.remove('hidden');
        }

        // 文件上傳處理
        async function handleFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            showLoading();
            hideError();

            try {
                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { cellStyles: true, cellFormulas: true, cellDates: true });
                
                // 處理Metrics表
                if (workbook.SheetNames.includes('Metrics')) {
                    const metricsSheet = workbook.Sheets['Metrics'];
                    const metricsRawData = XLSX.utils.sheet_to_json(metricsSheet, { header: 1 });
                    
                    const metricsHeaders = metricsRawData[1];
                    const metricsRows = metricsRawData.slice(2);
                    
                    const processedMetrics = metricsRows.map(row => {
                        const obj = {};
                        metricsHeaders.forEach((header, index) => {
                            obj[header] = row[index];
                        });
                        return obj;
                    }).filter(row => row.Date);
                    
                    const processedMetricsData = processedMetrics.map(row => {
                        // 處理日期格式（MM/DD/YYYY -> YYYY-MM-DD）
                        let formattedDate = row.Date;
                        if (typeof row.Date === 'string' && row.Date.includes('/')) {
                            const parts = row.Date.split('/');
                            if (parts.length === 3) {
                                const month = parts[0].padStart(2, '0');
                                const day = parts[1].padStart(2, '0');
                                const year = parts[2];
                                formattedDate = `${year}-${month}-${day}`;
                            }
                        }

                        const totalImpressions = Math.max(0, parseFloat(row['Impressions (organic)'] || 0));
                        const uniqueImpressions = Math.max(0, parseFloat(row['Unique impressions (organic)'] || 0));
                        
                        const adjustedUniqueImpressions = Math.min(uniqueImpressions, totalImpressions);
                        const repeatImpressions = Math.max(0, totalImpressions - adjustedUniqueImpressions);
                        
                        const totalClicks = parseFloat(row['Clicks (organic)'] || 0);
                        const clickThroughRate = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
                        
                        // 除錯：檢查是否有異常數據
                        if (adjustedUniqueImpressions < 0 || repeatImpressions < 0) {
                            console.log('異常數據:', {
                                date: formattedDate,
                                original_total: row['Impressions (organic)'],
                                original_unique: row['Unique impressions (organic)'],
                                processed_total: totalImpressions,
                                processed_unique: adjustedUniqueImpressions,
                                repeat: repeatImpressions
                            });
                        }
                        
                        // 除錯：檢查點閱率計算
                        if (totalClicks > 0) {
                            console.log('點閱率計算:', {
                                date: formattedDate,
                                clicks: totalClicks,
                                impressions: totalImpressions,
                                ctr: clickThroughRate
                            });
                        }
                        
                        return {
                            date: formattedDate,
                            uniqueImpressions: roundToTwo(adjustedUniqueImpressions),
                            repeatImpressions: roundToTwo(repeatImpressions),
                            totalImpressions: roundToTwo(totalImpressions),
                            clickThroughRate: roundToTwo(clickThroughRate)
                        };
                    }).filter(row => row.date && !isNaN(row.totalImpressions))
                     .sort((a, b) => new Date(a.date) - new Date(b.date));
                    
                    metricsData = processedMetricsData;
                }
                
                // 處理All posts表
                if (workbook.SheetNames.includes('All posts')) {
                    const allPostsSheet = workbook.Sheets['All posts'];
                    const allPostsRawData = XLSX.utils.sheet_to_json(allPostsSheet, { header: 1 });
                    
                    const allPostsHeaders = allPostsRawData[1];
                    const allPostsRows = allPostsRawData.slice(2);
                    
                    const processedAllPosts = allPostsRows.map(row => {
                        const obj = {};
                        allPostsHeaders.forEach((header, index) => {
                            obj[header] = row[index];
                        });
                        return obj;
                    }).filter(row => row['Created date']);
                    
                    const processedPostsData = processedAllPosts.map((row, index) => {
                        // 處理日期格式（MM/DD/YYYY -> YYYY-MM-DD）
                        let formattedDate = row['Created date'];
                        if (typeof row['Created date'] === 'string' && row['Created date'].includes('/')) {
                            const parts = row['Created date'].split('/');
                            if (parts.length === 3) {
                                const month = parts[0].padStart(2, '0');
                                const day = parts[1].padStart(2, '0');
                                const year = parts[2];
                                formattedDate = `${year}-${month}-${day}`;
                            }
                        }

                        return {
                            id: index,
                            date: formattedDate,
                            // 不保存連結信息到數據中
                            impressions: roundToTwo(parseFloat(row['Impressions'] || 0)),
                            ctr: roundToTwo(parseFloat(row['Click through rate (CTR)'] || 0) * 100),
                            title: row['Post title'] ? row['Post title'].substring(0, 50) + '...' : 'No title',
                            link: row['Post link'] // 只在表格中需要時使用
                        };
                    }).sort((a, b) => new Date(a.date) - new Date(b.date));
                    
                    postsData = processedPostsData;
                }
                
                // 重置分頁
                metricsCurrentPage = 0;
                postsCurrentPage = 0;
                
                // 更新UI
                updateStats();
                updateTables();
                createMetricsChart();
                createPostsChart();
                
            } catch (err) {
                showError('檔案讀取失敗，請確認檔案格式是否正確');
                console.error('Error processing file:', err);
            } finally {
                hideLoading();
            }
        }

        // 改變時間分組
        function changeTimeGrouping(newGrouping) {
            timeGrouping = newGrouping;
            
            // 更新按鈕樣式
            document.getElementById('dailyBtn').className = 'px-3 py-1 rounded-md text-sm ' + 
                (timeGrouping === 'daily' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300');
            document.getElementById('monthlyBtn').className = 'px-3 py-1 rounded-md text-sm ' + 
                (timeGrouping === 'monthly' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300');
            document.getElementById('quarterlyBtn').className = 'px-3 py-1 rounded-md text-sm ' + 
                (timeGrouping === 'quarterly' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300');
            
            // 重新創建圖表
            createMetricsChart();
        }

        // 初始化事件監聽器
        function initializeEventListeners() {
            // 文件上傳
            document.getElementById('fileInput').addEventListener('change', handleFileUpload);
            
            // 時間分組按鈕
            document.getElementById('dailyBtn').addEventListener('click', () => changeTimeGrouping('daily'));
            document.getElementById('monthlyBtn').addEventListener('click', () => changeTimeGrouping('monthly'));
            document.getElementById('quarterlyBtn').addEventListener('click', () => changeTimeGrouping('quarterly'));
        }

        // 初始化應用
        document.addEventListener('DOMContentLoaded', function() {
            initializeEventListeners();
        });
    </script>
</body>
</html>
