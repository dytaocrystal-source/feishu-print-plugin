// 飞书多维表格打印插件主逻辑

// 全局变量
let table = null;
let currentRecords = [];
let currentFields = [];

// 初始化
async function init() {
    try {
        // 等待飞书 SDK 就绪
        await bitable.bridge.checkInit();
        console.log('飞书 SDK 初始化成功');

        // 获取当前表格
        const selection = await bitable.base.getSelection();
        table = await bitable.base.getTableById(selection.tableId);
        
        console.log('表格加载成功');
    } catch (error) {
        console.error('初始化失败:', error);
        alert('插件初始化失败，请刷新重试');
    }
}

// 获取表格数据
async function getTableData(range = 'current') {
    showLoading(true);
    
    try {
        // 获取字段列表
        currentFields = await table.getFieldMetaList();
        
        // 根据范围获取记录
        switch (range) {
            case 'current':
                // 获取当前视图的记录
                const view = await table.getActiveView();
                const recordIds = await view.getVisibleRecordIdList();
                currentRecords = await Promise.all(
                    recordIds.map(id => table.getRecordById(id))
                );
                break;
                
            case 'all':
                // 获取所有记录
                currentRecords = await table.getRecords({
                    pageSize: 5000
                });
                break;
                
            case 'selected':
                // 获取选中的记录
                const selection = await bitable.base.getSelection();
                if (selection.recordId) {
                    currentRecords = [await table.getRecordById(selection.recordId)];
                } else {
                    alert('请先选择要打印的记录');
                    currentRecords = [];
                }
                break;
        }
        
        console.log(`获取到 ${currentRecords.length} 条记录`);
        return { fields: currentFields, records: currentRecords };
        
    } catch (error) {
        console.error('获取数据失败:', error);
        alert('获取数据失败，请重试');
        return { fields: [], records: [] };
    } finally {
        showLoading(false);
    }
}

// 格式化单元格值
function formatCellValue(value, fieldType) {
    if (value === null || value === undefined) {
        return '';
    }
    
    switch (fieldType) {
        case 1: // 文本
            return String(value);
            
        case 2: // 数字
            return typeof value === 'number' ? value.toLocaleString() : value;
            
        case 3: // 单选
            return value?.text || '';
            
        case 4: // 多选
            return Array.isArray(value) ? value.map(v => v.text).join(', ') : '';
            
        case 5: // 日期
            return value ? new Date(value).toLocaleDateString('zh-CN') : '';
            
        case 7: // 复选框
            return value ? '☑' : '☐';
            
        case 11: // 人员
            return Array.isArray(value) ? value.map(v => v.name).join(', ') : value?.name || '';
            
        case 13: // 电话
        case 15: // 超链接
        case 17: // 附件
            return Array.isArray(value) ? value.map(v => v.name || v.text || v).join(', ') : String(value);
            
        default:
            return String(value);
    }
}

// 生成打印 HTML
function generatePrintHTML(fields, records, options = {}) {
    const {
        showBorder = true,
        showHeader = true,
        orientation = 'portrait'
    } = options;
    
    let html = '<div class="print-container">';
    
    // 添加标题（可选）
    html += '<div style="text-align: center; margin-bottom: 20px;">';
    html += `<h2 style="font-size: 18px; font-weight: 600;">表格打印</h2>`;
    html += `<p style="font-size: 12px; color: #999;">打印时间: ${new Date().toLocaleString('zh-CN')}</p>`;
    html += '</div>';
    
    // 生成表格
    const borderClass = showBorder ? '' : 'no-border';
    html += `<table class="print-table ${borderClass}">`;
    
    // 表头
    if (showHeader) {
        html += '<thead><tr>';
        fields.forEach(field => {
            html += `<th>${field.name}</th>`;
        });
        html += '</tr></thead>';
    }
    
    // 表体
    html += '<tbody>';
    records.forEach(record => {
        html += '<tr>';
        fields.forEach(field => {
            const value = record.fields[field.id];
            const formattedValue = formatCellValue(value, field.type);
            html += `<td>${formattedValue}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody>';
    
    html += '</table>';
    html += '</div>';
    
    return html;
}

// 预览
async function preview() {
    const range = document.getElementById('printRange').value;
    const showBorder = document.getElementById('showBorder').checked;
    const showHeader = document.getElementById('showHeader').checked;
    const orientation = document.getElementById('orientation').value;
    
    const { fields, records } = await getTableData(range);
    
    if (records.length === 0) {
        alert('没有数据可预览');
        return;
    }
    
    const html = generatePrintHTML(fields, records, {
        showBorder,
        showHeader,
        orientation
    });
    
    document.getElementById('previewContent').innerHTML = html;
    document.getElementById('previewPanel').style.display = 'block';
}

// 打印
async function print() {
    const range = document.getElementById('printRange').value;
    const showBorder = document.getElementById('showBorder').checked;
    const showHeader = document.getElementById('showHeader').checked;
    const orientation = document.getElementById('orientation').value;
    const paperSize = document.getElementById('paperSize').value;
    
    const { fields, records } = await getTableData(range);
    
    if (records.length === 0) {
        alert('没有数据可打印');
        return;
    }
    
    const html = generatePrintHTML(fields, records, {
        showBorder,
        showHeader,
        orientation
    });
    
    // 设置打印区域
    const printArea = document.getElementById('printArea');
    printArea.innerHTML = html;
    printArea.style.display = 'block';
    
    // 设置打印样式
    const style = document.createElement('style');
    style.textContent = `
        @page {
            size: ${paperSize} ${orientation};
            margin: 1.5cm;
        }
    `;
    document.head.appendChild(style);
    
    // 延迟打印，确保内容渲染完成
    setTimeout(() => {
        window.print();
        printArea.style.display = 'none';
        document.head.removeChild(style);
    }, 300);
}

// 导出 PDF
async function exportPDF() {
    const range = document.getElementById('printRange').value;
    const showBorder = document.getElementById('showBorder').checked;
    const showHeader = document.getElementById('showHeader').checked;
    const orientation = document.getElementById('orientation').value;
    
    const { fields, records } = await getTableData(range);
    
    if (records.length === 0) {
        alert('没有数据可导出');
        return;
    }
    
    showLoading(true, '正在生成 PDF...');
    
    try {
        const html = generatePrintHTML(fields, records, {
            showBorder,
            showHeader,
            orientation
        });
        
        // 创建临时容器
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;
        tempDiv.style.position = 'absolute';
        tempDiv.style.left = '-9999px';
        tempDiv.style.background = 'white';
        tempDiv.style.padding = '20px';
        document.body.appendChild(tempDiv);
        
        // 使用 html2canvas 转换为图片
        const canvas = await html2canvas(tempDiv, {
            scale: 2,
            backgroundColor: '#ffffff'
        });
        
        // 创建 PDF
        const { jsPDF } = window.jspdf;
        const imgWidth = orientation === 'landscape' ? 297 : 210; // A4 尺寸
        const imgHeight = orientation === 'landscape' ? 210 : 297;
        const imgData = canvas.toDataURL('image/png');
        
        const pdf = new jsPDF({
            orientation: orientation === 'landscape' ? 'l' : 'p',
            unit: 'mm',
            format: 'a4'
        });
        
        // 计算图片在 PDF 中的尺寸
        const pdfWidth = imgWidth - 20; // 留边距
        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
        
        pdf.addImage(imgData, 'PNG', 10, 10, pdfWidth, pdfHeight);
        
        // 下载 PDF
        const filename = `表格打印_${new Date().getTime()}.pdf`;
        pdf.save(filename);
        
        // 清理临时元素
        document.body.removeChild(tempDiv);
        
        alert('PDF 导出成功！');
        
    } catch (error) {
        console.error('导出 PDF 失败:', error);
        alert('导出 PDF 失败，请重试');
    } finally {
        showLoading(false);
    }
}

// 显示/隐藏加载动画
function showLoading(show, text = '正在加载数据...') {
    const loading = document.getElementById('loading');
    if (show) {
        loading.querySelector('p').textContent = text;
        loading.style.display = 'flex';
    } else {
        loading.style.display = 'none';
    }
}

// 事件绑定
document.addEventListener('DOMContentLoaded', () => {
    // 初始化
    init();
    
    // 预览按钮
    document.getElementById('previewBtn').addEventListener('click', preview);
    
    // 打印按钮
    document.getElementById('printBtn').addEventListener('click', print);
    
    // 导出 PDF 按钮
    document.getElementById('exportPdfBtn').addEventListener('click', exportPDF);
    
    // 关闭预览
    document.getElementById('closePreview').addEventListener('click', () => {
        document.getElementById('previewPanel').style.display = 'none';
    });
});
