let workbookData = null;

document.addEventListener('DOMContentLoaded', function() {
    if (typeof XLSX === 'undefined') {
        alert('XLSX 库未正确加载，请确保网络连接正常');
        return;
    }
});

function handleFileUpload() {
    const fileInput = document.getElementById('fileUpload');
    const file = fileInput.files[0];
    
    if (!file) {
        alert('请先选择一个 Excel 文件');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbookData = XLSX.read(data, {type: 'array'});
            
            // 生成菜单
            generateMenu(workbookData.SheetNames);
            
            // 默认显示第一个工作表
            if (workbookData.SheetNames.length > 0) {
                showSheet(workbookData.SheetNames[0]);
            }
            
        } catch (error) {
            alert('处理 Excel 文件时出错：' + error.message);
            console.error('Excel 处理错误：', error);
        }
    };

    reader.onerror = function(error) {
        alert('读取文件时出错');
        console.error('文件读取错误：', error);
    };

    reader.readAsArrayBuffer(file);
}

function generateMenu(sheetNames) {
    const menu = document.getElementById('sheetMenu');
    menu.innerHTML = '';
    
    // 创建菜单项
    sheetNames.forEach(sheetName => {
        const menuItem = document.createElement('div');
        menuItem.className = 'menu-item';
        menuItem.textContent = sheetName;
        menuItem.onclick = () => {
            // 移除所有活动状态
            document.querySelectorAll('.menu-item').forEach(item => {
                item.classList.remove('active');
            });
            // 添加活动状态
            menuItem.classList.add('active');
            // 显示对应的工作表
            showSheet(sheetName);
        };
        menu.appendChild(menuItem);
    });
}

function showSheet(sheetName) {
    if (!workbookData) return;
    
    const tableContainer = document.getElementById('dataTable');
    tableContainer.innerHTML = '';
    
    // 创建工作表标题
    const sheetTitle = document.createElement('h2');
    sheetTitle.textContent = sheetName;
    tableContainer.appendChild(sheetTitle);
    
    // 获取当前工作表
    const worksheet = workbookData.Sheets[sheetName];
    
    // 转换HTML表格
    const htmlTable = XLSX.utils.sheet_to_html(worksheet, { editable: false });
    
    // 创建表格容器
    const tableDiv = document.createElement('div');
    tableDiv.innerHTML = htmlTable;
    tableContainer.appendChild(tableDiv);
    
    // 添加表格样式
    const table = tableDiv.getElementsByTagName('table')[0];
    if (table) {
        table.className = 'excel-table';
    }

    // 检查是否是本月收款或杂费表
    if (sheetName.includes('月分表')) {
        // 直接获取JSON数据，不使用header选项
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        console.log('读取到的数据:', jsonData);

        // 生成统计数据
        const statistics = {
            projectStats: {},
            noteStats: {}
        };
        
        let totalAmount = 0;
        
        // 处理每一行数据
        jsonData.forEach(row => {
            // 尝试不同可能的列名
            const amount = parseFloat(row['金额'] || row['金额（元）'] || row['Amount'] || 0);
            const project = row['项目'] || row['Project'] || '';
            const note = row['备注'] || row['Notes'] || '';

            console.log('正在处理行:', { amount, project, note });

            if (!isNaN(amount) && amount !== 0) {
                // 按项目统计
                if (project) {
                    statistics.projectStats[project] = (statistics.projectStats[project] || 0) + amount;
                }
                
                // 按备注统计
                if (note) {
                    statistics.noteStats[note] = (statistics.noteStats[note] || 0) + amount;
                }
                
                totalAmount += amount;
            }
        });

        console.log('统计结果:', statistics);
        
        // 更新总计显示
        if (sheetName.includes('本月收款')) {
            document.getElementById('monthlyIncome').textContent = totalAmount.toLocaleString('en-US', {
                style: 'currency',
                currency: 'USD'
            });
        } else if (sheetName.includes('杂费')) {
            document.getElementById('monthlyExpense').textContent = totalAmount.toLocaleString('en-US', {
                style: 'currency',
                currency: 'USD'
            });
        }
        
        // 显示统计表格和图表
        if (Object.keys(statistics.projectStats).length > 0 || Object.keys(statistics.noteStats).length > 0) {
            showStatistics(statistics, tableContainer);
        }
    }
}

function showStatistics(statistics, container) {
    // 创建统计容器
    const statsContainer = document.createElement('div');
    statsContainer.className = 'statistics-container';
    
    // 项目统计表格
    if (Object.keys(statistics.projectStats).length > 0) {
        const projectStatsSection = document.createElement('div');
        projectStatsSection.className = 'stats-section';
        
        const projectTitle = document.createElement('h3');
        projectTitle.textContent = '按项目统计';
        projectStatsSection.appendChild(projectTitle);
        
        // 创建项目统计表格
        const projectTable = createStatsTable(statistics.projectStats);
        projectStatsSection.appendChild(projectTable);
        
        // 创建项目统计图表
        const projectChartCanvas = document.createElement('canvas');
        projectChartCanvas.id = 'projectChart';
        projectStatsSection.appendChild(projectChartCanvas);
        
        createBarChart(statistics.projectStats, projectChartCanvas, '项目金额统计');
        
        statsContainer.appendChild(projectStatsSection);
    }
    
    // 备注统计表格
    if (Object.keys(statistics.noteStats).length > 0) {
        const noteStatsSection = document.createElement('div');
        noteStatsSection.className = 'stats-section';
        
        const noteTitle = document.createElement('h3');
        noteTitle.textContent = '按备注统计';
        noteStatsSection.appendChild(noteTitle);
        
        // 创建备注统计表格
        const noteTable = createStatsTable(statistics.noteStats);
        noteStatsSection.appendChild(noteTable);
        
        // 创建备注统计图表
        const noteChartCanvas = document.createElement('canvas');
        noteChartCanvas.id = 'noteChart';
        noteStatsSection.appendChild(noteChartCanvas);
        
        createBarChart(statistics.noteStats, noteChartCanvas, '备注金额统计');
        
        statsContainer.appendChild(noteStatsSection);
    }
    
    container.appendChild(statsContainer);
}

function createStatsTable(data) {
    const table = document.createElement('table');
    table.className = 'stats-table';
    
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['类别', '金额'].forEach(text => {
        const th = document.createElement('th');
        th.textContent = text;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    const tbody = document.createElement('tbody');
    Object.entries(data).forEach(([key, value]) => {
        const row = document.createElement('tr');
        
        const keyCell = document.createElement('td');
        keyCell.textContent = key;
        row.appendChild(keyCell);
        
        const valueCell = document.createElement('td');
        valueCell.textContent = value.toLocaleString('en-US', {
            style: 'currency',
            currency: 'USD'
        });
        row.appendChild(valueCell);
        
        tbody.appendChild(row);
    });
    table.appendChild(tbody);
    
    return table;
}

function createBarChart(data, canvas, title) {
    // 确保有数据要显示
    if (Object.keys(data).length === 0) {
        console.log('没有数据用于创建图表');
        return;
    }

    const chartData = {
        labels: Object.keys(data),
        datasets: [{
            label: '金额',
            data: Object.values(data),
            backgroundColor: 'rgba(54, 162, 235, 0.6)',
            borderColor: 'rgba(54, 162, 235, 1)',
            borderWidth: 1
        }]
    };

    console.log('图表数据:', chartData); // 调试日志

    new Chart(canvas, {
        type: 'bar',
        data: chartData,
        options: {
            responsive: true,
            maintainAspectRatio: false, // 添加这行
            plugins: {
                title: {
                    display: true,
                    text: title
                },
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw;
                            return value.toLocaleString('en-US', {
                                style: 'currency',
                                currency: 'USD'
                            });
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value.toLocaleString('en-US', {
                                style: 'currency',
                                currency: 'USD'
                            });
                        }
                    }
                }
            }
        }
    });
}

// 添加统计计算函数
function calculateMonthlyStats() {
    const currentTab = document.querySelector('.tab-content.active');
    if (!currentTab) return;
    
    let totalIncome = 0;
    let totalExpense = 0;
    
    // 获取当前表格中的所有行
    const rows = currentTab.querySelectorAll('tbody tr');
    
    rows.forEach(row => {
        // 获取收款金额（假设是第4列）
        const incomeCell = row.cells[3];
        if (incomeCell) {
            const incomeValue = parseFloat(incomeCell.textContent) || 0;
            totalIncome += incomeValue;
        }
        
        // 获取杂费金额（假设是第5列）
        const expenseCell = row.cells[4];
        if (expenseCell) {
            const expenseValue = parseFloat(expenseCell.textContent) || 0;
            totalExpense += expenseValue;
        }
    });
    
    // 更新统计显示
    document.getElementById('monthlyIncome').textContent = totalIncome.toFixed(2);
    document.getElementById('monthlyExpense').textContent = totalExpense.toFixed(2);
}

// 在切换标签页时更新统计
function switchTab(month) {
    // 原有的切换标签页代码...
    
    // 切换后计算统计
    calculateMonthlyStats();
}

// 在添加或编辑数据后更新统计
function addRow() {
    // 原有的添加行代码...
    
    // 添加行后更新统计
    calculateMonthlyStats();
}

// 在页面加载时计算初始统计
document.addEventListener('DOMContentLoaded', () => {
    calculateMonthlyStats();
}); 