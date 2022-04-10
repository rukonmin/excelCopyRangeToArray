/**
 * When copy range data in excel the data structure is:
 * {cell value}\t{cell value}\t{cell value}\n
 * {cell value}\t{cell value}\t{cell value}\n
 * {cell value}\t{cell value}\t{cell value}\n
 *
 * where {cell value} is a value of cell in excel
 *
 * if cell containes \n or \t, then cell value will be in quotation mark ("{cell value}"),
 * and all other quotation mark will be "" (" => "")
 *
 * @param {string} data
 * @returns {string[][]}
 * @example
 * val11\tval12\tval13\n
 * val21\tval22\tval23\n
 * =>
 * [
 * [val11, val12, val13]
 * [val21, val22, val23]
 * ]
 */
export function excelCopyRangeToArray(data) {
    const result = [];
    const safeSymbolsRegExp = /[\n\t]/g;
    const safeCellRegExp = /^"([^"]*(?:""[^"]*)*)"/;
    const generalCellRegExp = /^([^\t\r\n]*)/;

    let tmp, cellValue = '', row = [];
    while (data.length) {
        switch (data[0]) {
            case '"':
                tmp = data.match(safeCellRegExp);

                if (safeSymbolsRegExp.test(tmp[1])) {
                    cellValue = tmp[1].replace(/""/g, '"');
                } else {
                    tmp = data.match(generalCellRegExp);
                    cellValue = tmp[1];
                }

                data = data.substring(tmp[0].length);
                break;
            case '\t':
                row.push(cellValue);
                cellValue = '';
                data = data.substring(1);
                break;
            case '\r':
                data = data.substring(1);
                break
            case '\n':
                row.push(cellValue);
                cellValue = '';
                data = data.substring(1);
                result.push(row)
                row = [];
                break;
            default:
                tmp = data.match(generalCellRegExp);
                cellValue = tmp[1];
                data = data.substring(tmp[0].length);
                break;
        }
    }

    return result;
}