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
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">{table}</table>
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
 * 生成图片单元格
 */
const generateImageCell = (images, options = {}) => {
	// 默认宽度、高度60px
	const defaultHeight = 60;
	const defaultWidth = 60;
	const { width = defaultWidth, height = defaultHeight } = options;

	// 如果是单图，转为数组形式统一处理
	const imageList = Array.isArray(images) ? images : [images];

	// 计算每张图片的宽度（如果没指定宽度，则自动平分）
	const imgWidth = width ? Math.floor(width / imageList.length) - 2 : '';

	// 生成图片HTML
	const imagesHtml = imageList.map(img => `
    <td align="center" valign="middle" ${imgWidth ? `width="${imgWidth}"` : ''}>
      <img src="${img}" ${imgWidth ? `width="${imgWidth-2}"` : ''} height="${height-2}" style="display:block;"/>
    </td>
  `).join('');

	return `
    <td ${width ? `width="${width}"` : ''} height="${height}">
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
          ${imagesHtml}
        </tr>
      </table>
    </td>
  `;
};

/**
 * 生成单元格HTML
 */
const generateCellHtml = (type, value, options = {}) => {
	const { width, height, mergeOptions } = options;

	if (type === 'image' || type === 'images') {
		return generateImageCell(value, { width, height });
	}

	const mergeAttrs = getMergeAttributes(mergeOptions);
	return `<td ${mergeAttrs} ${width ? `width="${width}"` : ''} ${height ? `height="${height}"` : ''} style="padding:20px;">${value || ''}</td>`;
};

/**
 * 表格导出Excel主函数
 */
const table2excel = (options) => {
	if (!options || !options.column || !options.data) {
		console.error('缺少必要的参数: column 和 data');
		return;
	}

	const { column, data, excelName = 'export', captionName } = options;

	// 预处理数据，标记被合并的单元格
	const processedData = JSON.parse(JSON.stringify(data));
	processedData.forEach((row, rowIndex) => {
		if (row.mergeOptions) {
			row.__mergedColumns = [];
			Object.entries(row.mergeOptions).forEach(([key, options]) => {
				if (options.rowspan > 1) {
					for (let i = 1; i < options.rowspan; i++) {
						if (processedData[rowIndex + i]) {
							processedData[rowIndex + i].__mergedColumns =
								processedData[rowIndex + i].__mergedColumns || [];
							processedData[rowIndex + i].__mergedColumns.push(key);
						}
					}
				}
			});
		}
	});

	// 生成表头
	const thead = column.reduce((html, col) => {
		const mergeAttrs = getMergeAttributes(col.mergeOptions);
		return html + `<th ${mergeAttrs} ${col.width ? `width="${col.width}"` : ''} ${col.height ? `height="${col.height}"` : 'height="40"'}  style="background-color:#d9d9d9;">${col.title}</th>`;
	}, '');

	// 生成表格行
	const tbody = processedData.map((row) => {
		const cells = column.map((col) => {
			if (row.__mergedColumns && row.__mergedColumns.includes(col.key)) {
				return '';
			}
			return generateCellHtml(
				col.type || 'text',
				row[col.key],
				{
					width: col.width,
					height: col.height,
					mergeOptions: row.mergeOptions ? row.mergeOptions[col.key] : null
				}
			);
		}).join('');
		return `<tr>${cells}</tr>`;
	}).join('');

	// 构建完整表格
	const caption = captionName ? `<caption><b>${captionName}</b></caption>` : '';
	const table = `
    ${caption}
    <thead><tr>${thead}</tr></thead>
    <tbody>${tbody}</tbody>
  `;

	// 导出表格
	exportToExcel(table, excelName);
};

export default table2excel;
