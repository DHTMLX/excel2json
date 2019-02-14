import "../node_modules/fast-text-encoding/text";
import '../wasm/xlsx_export';


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


let XLSX = null;
function doConvert(input, config){
    const path = config.wasmPath || "https://cdn.dhtmlx.com/libs/excel2json/1.wasm";
    const getStyles = config.styles === undefined ? true : config.styles;

    if (XLSX) {
        const xlsx = XLSX.new(input);
        const styles = getStyles ? xlsx.get_styles() : null;

        let sheetsData;
        if (config.sheet) {
            const data = xlsx.get_sheet_data(sheet);
            sheetsData = [data];
        } else {
            const sheets = xlsx.get_sheets();
            sheetsData = sheets.map(name => xlsx.get_sheet_data(name));
        }

        postMessage({
            uid: config.uid || (new Date()).valueOf(),
            type: "ready",
            data: sheetsData,
            styles
        });
    } else {
        wasm_bindgen(path).then(() => {
            XLSX = wasm_bindgen.XLSX;
            doConvert(input, config);
        }).catch(e => console.log(e));
    }
}