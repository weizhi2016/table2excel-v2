/* eslint-disable */
let idTmr;

/**
 * 检测浏览器类型
 */
const getExplorer = () => {
	const explorer = window.navigator.userAgent;
	if (explorer.indexOf("MSIE") >= 0) return 'ie';
	if (explorer.indexOf("Firefox") >= 0) return 'Firefox';
	if (explorer.indexOf("Chrome") >= 0) return 'Chrome';
	return 'Other';
};

/**
 * 清理定时器
 */
const cleanup = () => {
	window.clearInterval(idTmr);
};

/**
 * IE浏览器导出Excel
 */
const tableToIE = (data, name) => {
	try {
		const oXL = new ActiveXObject("Excel.Application");
		const oWB = oXL.Workbooks.Add();
		const xlsheet = oWB.Worksheets(1);
		const sel = document.body.createTextRange();

		sel.moveToElementText(data);
		sel.select();
		sel.execCommand("Copy");
		xlsheet.Paste();
		oXL.Visible = true;

		const fname = oXL.Application.GetSaveAsFilename("Excel.xls", "Excel Spreadsheets (*.xls), *.xls");
		oWB.SaveAs(fname);
		oWB.Close(false);
		oXL.Quit();
		oXL = null;

		idTmr = window.setInterval(cleanup, 1);
	} catch (e) {
		console.error("IE导出失败:", e);
		alert("IE导出失败，请检查浏览器设置或使用其他浏览器");
	}
};

/**
 * 非IE浏览器导出Excel
 */
const tableToNotIE = (function() {
	const uri = 'data:application/vnd.ms-excel;base64,';
	const template = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" 
          xmlns:x="urn:schemas-microsoft-com:office:excel" 
          xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8"/>
      <!--[if gte mso 9]>
      <xml>
        <x:ExcelWorkbook>
          <x:ExcelWorksheets>
            <x:ExcelWorksheet>
              <x:Name>{worksheet}</x:Name>
              <x:WorksheetOptions>
                <x:DisplayGridlines/>
              </x:WorksheetOptions>
            </x:ExcelWorksheet>
          </x:ExcelWorksheets>
        </x:ExcelWorkbook>
      </xml>
      <![endif]-->
    </head>
    <body>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;">
        {table}
      </table>
    </body>
    </html>
  `;

	const base64 = (s) => window.btoa(unescape(encodeURIComponent(s)));
	const format = (s, c) => s.replace(/{(\w+)}/g, (m, p) => c[p]);

	return (table, name) => {
		const ctx = { worksheet: name || 'Sheet1', table };
		const url = uri + base64(format(template, ctx));

		if (navigator.userAgent.indexOf("Firefox") > -1) {
			window.location.href = url;
		} else {
			const aLink = document.createElement('a');
			aLink.href = url;
			aLink.download = (name || 'export') + '.xls';

			const event = new MouseEvent('click', {
				view: window,
				bubbles: true,
				cancelable: true
			});

			aLink.dispatchEvent(event);

			setTimeout(() => {
				window.URL.revokeObjectURL(url);
			}, 100);
		}
	};
})();

/**
 * 导出到Excel
 */
const exportToExcel = (data, name) => {
	getExplorer() === 'ie' ? tableToIE(data, name) : tableToNotIE(data, name);
};

/**
 * 生成合并单元格属性
 */
const getMergeAttributes = (mergeOptions) => {
	if (!mergeOptions) return '';
	let attrs = '';
	if (mergeOptions.rowspan > 1) attrs += ` rowspan="${mergeOptions.rowspan}"`;
	if (mergeOptions.colspan > 1) attrs += ` colspan="${mergeOptions.colspan}"`;
	return attrs;
};

/**
 * 准备合并单元格数据（终极修复版）
 */
const prepareMergedData = (data, columns) => {
	const processedData = JSON.parse(JSON.stringify(data));
	const columnKeys = columns.map(col => col.key);

	// 创建合并映射表
	const mergeMap = {};

	// 第一遍：处理行合并
	processedData.forEach((row, rowIndex) => {
		if (row.mergeOptions) {
			row.__rowMergeMap = {};
			row.__colMergeMap = {};

			Object.entries(row.mergeOptions).forEach(([key, options]) => {
				const colIndex = columnKeys.indexOf(key);
				if (colIndex >= 0) {
					// 处理行合并
					if (options.rowspan > 1) {
						for (let i = 1; i < options.rowspan; i++) {
							if (!mergeMap[rowIndex + i]) mergeMap[rowIndex + i] = {};
							mergeMap[rowIndex + i][colIndex] = true;
						}
					}

					// 处理列合并
					if (options.colspan > 1) {
						for (let i = 1; i < options.colspan; i++) {
							if (colIndex + i < columnKeys.length) {
								row.__colMergeMap[colIndex + i] = true;
							}
						}
					}
				}
			});
		}
	});

	// 第二遍：应用合并映射
	processedData.forEach((row, rowIndex) => {
		if (mergeMap[rowIndex]) {
			row.__rowMergeMap = mergeMap[rowIndex];
		}
	});

	return processedData;
};

/**
 * 生成表头单元格
 */
const generateHeaderCell = (col) => {
	return `
    <th ${getMergeAttributes(col.mergeOptions)} 
        style="background-color:#d9d9d9;height:40px;${col.width ? `width:${col.width}px;` : ''}padding:5px;text-align:center;">
      ${col.title}
    </th>
  `;
};

/**
 * 生成数据单元格（支持多种类型）
 */
const generateDataCell = (value, col, options = {}) => {
	const { width, height = 40, mergeOptions, isMerged } = options;

	if (isMerged) return ''; // 被合并的单元格留空

	// 图片类型单元格
	if (col.type === 'image' || col.type === 'images') {
		const imageList = Array.isArray(value) ? value : [value];
		return `
      <td ${getMergeAttributes(mergeOptions)} 
          style="padding:1px;height:${height}px;${width ? `width:${width}px;` : ''}">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
          <tr>
            ${imageList.map(img => `
              <td align="center" valign="middle" style="padding:1px;">
                <img src="${img}" style="display:block;height:${height-2}px;${width ? `max-width:${Math.floor(width/imageList.length)-2}px;` : 'width:auto;'}"/>
              </td>
            `).join('')}
          </tr>
        </table>
      </td>
    `;
	}

	// 普通数据单元格
	return `
    <td ${getMergeAttributes(mergeOptions)} 
        style="padding:5px;${width ? `width:${width}px;` : ''}${height ? `height:${height}px;` : ''}">
      ${value || ''}
    </td>
  `;
};

/**
 * 通用表格导出Excel主函数（终极修复版）
 */
const table2excel = (options) => {
	if (!options || !options.column || !options.data) {
		console.error('缺少必要的参数: column 和 data');
		return;
	}

	const { column, data, excelName = 'export', captionName } = options;
	const columnKeys = column.map(col => col.key);

	// 预处理数据，正确标记被合并的单元格
	const processedData = prepareMergedData(data, column);

	// 生成表头行
	const thead = `<tr>${column.map(col => generateHeaderCell(col)).join('')}</tr>`;

	// 生成数据行
	const tbody = processedData.map((row) => {
		const cells = column.map((col, colIndex) => {
			// 检查是否被行合并或列合并
			const isRowMerged = row.__rowMergeMap?.[colIndex];
			const isColMerged = row.__colMergeMap?.[colIndex];

			return generateDataCell(row[col.key], col, {
				width: col.width,
				height: col.height,
				mergeOptions: row.mergeOptions?.[col.key],
				isMerged: isRowMerged || isColMerged
			});
		}).join('');

		return `<tr>${cells}</tr>`;
	}).join('');

	// 构建完整表格
	const table = `
    ${captionName ? `<caption><b>${captionName}</b></caption>` : ''}
    <thead>${thead}</thead>
    <tbody>${tbody}</tbody>
  `;

	exportToExcel(table, excelName);
};

export default table2excel;
