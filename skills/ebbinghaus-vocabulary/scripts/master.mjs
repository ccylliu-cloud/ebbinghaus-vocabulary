import fs from 'node:fs/promises';
const [input, output, modulePath] = process.argv.slice(2);
const { Workbook, SpreadsheetFile } = await import(modulePath);
const data = JSON.parse(await fs.readFile(input, 'utf8'));
const wb = Workbook.create();
const sheet = wb.worksheets.add('单词总表');
sheet.showGridLines = false;
const total = data.pages.length * data.rows_per_page;
const all = sheet.getRange(`A1:${data.last_column}${total}`);
all.format.font = {name:'Arial Unicode MS', size:9, color:'#111111'};
all.format.verticalAlignment = 'center';
all.format.rowHeight = 6;
all.format.wrapText = true;
data.widths.forEach((width, i) => {
  sheet.getRangeByIndexes(0, i, total, 1).format.columnWidthPx = width * 96 / 72;
});
function merged(row, col, height, width, value, kind) {
  const range = sheet.getRangeByIndexes(row - 1, col, height, width);
  range.merge();
  range.values = [[value.startsWith('=') ? "'" + value : value]];
  if (kind === 'title') range.format.font = {name:'Arial Unicode MS',size:12,bold:true};
  if (kind === 'group') {
    range.format.fill = '#E4E4E4';
    range.format.font.bold = true;
  }
  if (kind === 'header') {
    range.format.fill = '#F4F4F4';
    range.format.font.bold = true;
    range.format.horizontalAlignment = 'center';
  }
  if (kind === 'body') range.format.borders = {bottom:{style:'thin',color:'#DDDDDD'}};
}
data.pages.forEach((page, pageIndex) => {
  const offset = pageIndex * data.rows_per_page;
  merged(offset+1, 0, 3, data.widths.length, '单词总表', 'title');
  for (const block of page) {
    const col = block.side * (data.area_columns + 1);
    let position = 0;
    const columns = block.widths.map(width => {
      const left = data.edges.indexOf(position);
      position += width;
      return [left, data.edges.indexOf(position) - left];
    });
    let row = offset + block.row;
    merged(row, col, 3, data.area_columns, `第${block.group}组`, 'group');
    row += 3;
    block.headers.forEach((v,j)=>merged(row,col+columns[j][0],3,columns[j][1],v,'header'));
    row += 3;
    for (const dataRow of block.rows) {
      dataRow.values.forEach((v,j)=>merged(row,col+columns[j][0],dataRow.units,columns[j][1],v,'body'));
      row += dataRow.units;
    }
  }
});
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:sheet.name,range:`A1:${data.last_column}16`,maxChars:700,tableMaxRows:4})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?',options:{useRegex:true,maxResults:10},maxChars:500})).ndjson);
if (process.env.VOCAB_RENDER_DIR) {
  await fs.mkdir(process.env.VOCAB_RENDER_DIR,{recursive:true});
  const preview = await wb.render({sheetName:sheet.name,range:`A1:${data.last_column}${data.rows_per_page}`,scale:1.3,format:'png'});
  await fs.writeFile(`${process.env.VOCAB_RENDER_DIR}/总表预览.png`,new Uint8Array(await preview.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
