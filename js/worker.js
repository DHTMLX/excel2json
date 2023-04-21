import { XLSX } from '../pkg/excel2json_wasm';


onmessage = function(e) {
    const config = e.data;

    if (e.data.type === "convert") {
        const data = e.data.data;
        if (data instanceof File){
            const reader =  new FileReader();
            reader.readAsArrayBuffer(data);
            reader.onload = (e) => doConvert(new Int8Array(e.target.result), config);
        } else {
            doConvert(data, config);
        }
    }
}

function doConvert(input, config){
    const getStyles = config.styles === undefined ? true : config.styles;

    const xlsx = XLSX.new(input);
    const styles = getStyles ? xlsx.get_styles() : null;

    let sheetsData;
    if (config.sheet) {
        const data = xlsx.get_sheet_data(sheet);
        sheetsData = [data];
    } else {
        const sheets = xlsx.get_sheets();
        const mode = 0 | (config.formulas ? XLSX.with_formulas(): 0);
        sheetsData = sheets.map(name => xlsx.get_sheet_data(name, mode));
    }

    postMessage({
        uid: config.uid || (new Date()).valueOf(),
        type: "ready",
        data: sheetsData,
        styles
    });
}

postMessage({ type:"init" });