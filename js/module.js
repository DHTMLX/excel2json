import { XLSX } from "../pkg/excel2json_wasm.js";

export { XLSX };
export function convertArray(jsonData, config = {}){
    const getStyles = config.styles === undefined ? true : config.styles;
    const xlsx = XLSX.new(jsonData);
    const styles = getStyles ? xlsx.get_styles() : null;

    let data;
    if (config.sheet) {
        data = [xlsx.get_sheet_data(sheet)];
    } else {
        const sheets = xlsx.get_sheets();
        const mode = 0 | (config.formulas ? XLSX.with_formulas(): 0);
        data = sheets.map(name => xlsx.get_sheet_data(name, mode));
    }

    return { data, styles };
}

export function convert(jsonData, config = {}){
    if (jsonData instanceof File){
        return new Promise((res) => {
            const reader =  new FileReader();
            reader.readAsArrayBuffer(jsonData);
            reader.onload = (e) => {
                res(convertArray(new Int8Array(e.target.result), config));
            };
        });
    }
}
