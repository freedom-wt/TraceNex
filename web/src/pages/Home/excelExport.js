import * as XLSX from 'xlsx';
import * as XLSXStyle from 'xlsx-js-style';
/**
 * 前端导出Excel通用函数
 * @param {Object} options 导出配置
 * @param {Array} options.data 数据源（对象数组，必传）
 * @param {String} options.fileName 导出文件名（默认：导出数据.xlsx）
 * @param {String} options.sheetName 工作表名称（默认：Sheet1）
 * @param {Array} options.colWidths 列宽配置（[{wpx: 120}, {wpx: 80}, ...]，wpx为像素值）
 * @param {Boolean} options.needStyle 是否启用样式（默认：true）
 * @returns {void}
 */
export const exportExcel = (options) => {
  const {
    data = [],
    fileName = '导出数据.xlsx',
    sheetName = 'Sheet1',
    colWidths = [],
    needStyle = true
  } = options;

  // 步骤1：处理全量数据（解决分页场景下仅导出当前页的问题）
  // 若你的数据是分页加载的，需先聚合全量数据（示例：假设分页接口返回的全量数据数组为allData）
  const fullData = Array.isArray(data) ? [...data] : [];
  if (fullData.length === 0) {
    alert('暂无数据可导出！');
    return;
  }

  // 步骤2：将对象数组转为工作表（基础结构）
  const ws = XLSX.utils.json_to_sheet(fullData);

  // 步骤3：样式配置（列宽 + 表头/单元格样式）
  if (needStyle) {
    // 3.1 列宽配置（!cols是SheetJS定义的列宽属性）
    if (colWidths.length > 0) {
      ws['!cols'] = colWidths;
    } else {
      // 默认列宽：根据数据字段数自动分配，每个列宽120px
      const keys = Object.keys(fullData[0]);
      ws['!cols'] = keys.map(() => ({ wpx: 120 }));
    }

    // 3.2 表头样式（第一行：背景色、字体加粗、居中对齐）
    const headerRow = 0; // 表头行索引
    const headerKeys = Object.keys(fullData[0]);
    headerKeys.forEach((key, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: colIndex });
      // 设置单元格样式
      ws[cellAddress].s = {
        fill: { fgColor: { rgb: 'E8F4FD' } }, // 背景色（浅蓝）
        font: { bold: true }, // 字体加粗
        alignment: { horizontal: 'center', vertical: 'center' }, // 居中对齐
        border: { // 边框
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        }
      };
    });

    // 3.3 内容行样式（除表头外的行：居中对齐 + 边框）
    fullData.forEach((_, rowIndex) => {
      const currentRow = rowIndex + 1; // 内容行从第1行开始（表头是第0行）
      headerKeys.forEach((_, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: colIndex });
        if (ws[cellAddress]) { // 确保单元格存在
          ws[cellAddress].s = {
            alignment: { horizontal: 'center', vertical: 'center' },
            border: {
              top: { style: 'thin' },
              bottom: { style: 'thin' },
              left: { style: 'thin' },
              right: { style: 'thin' }
            }
          };
        }
      });
    });
  }

  // 步骤4：创建工作簿并加入工作表
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // 步骤5：生成Excel文件并触发下载（兼容xlsx-style的样式渲染）
  try {
    // xlsx-style需通过write方法生成二进制数据，再转Blob下载
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary' });
    function s2ab(s) {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
    // 创建Blob并触发下载
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url); // 释放URL对象
    console.log('Excel导出成功！');
  } catch (error) {
    console.error('Excel导出失败：', error);
    alert('导出失败，请检查数据格式！');
  }
};

/**
 * 从DOM表格提取数据并导出Excel（适配无数据源的场景）
 * @param {String} tableSelector 表格选择器（如#myTable）
 * @param {String} fileName 导出文件名
 * @returns {void}
 */
export const exportExcelFromTable = (tableSelector, fileName = '表格数据.xlsx') => {
  const table = document.querySelector(tableSelector);
  if (!table) {
    alert('未找到指定的表格！');
    return;
  }
  // 将DOM表格转为对象数组
  const ws = XLSX.utils.table_to_sheet(table);
  const data = XLSX.utils.sheet_to_json(ws);
  // 调用通用导出函数
  exportExcel({
    data,
    fileName,
    colWidths: Array.from({ length: data[0] ? Object.keys(data[0]).length : 0 }, () => ({ wpx: 100 }))
  });
};