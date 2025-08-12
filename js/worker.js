
import init, { XLSX } from '../pkg/excel2json_wasm.js';

onmessage = async function (e) {
    const config = e.data;

    if (config.type === "convert") {
        const file = config.data;
        if (file instanceof File) {
            const reader = new FileReader();
            reader.readAsArrayBuffer(file);
            reader.onload = (e) => {
                const input = new Uint8Array(e.target.result);
                doConvert(input, config);
            };
        } else {
            doConvert(file, config);
        }
    }
};

async function doConvert(input, config) {
    await init();

    const getStyles = config.styles === undefined ? true : config.styles;

    const xlsx = XLSX.new(input);
    const styles = getStyles ? xlsx.get_styles() : null;

    let sheetsData;
    if (config.sheet) {
        const data = xlsx.get_sheet_data(config.sheet);
        sheetsData = [data];
    } else {
        const sheets = xlsx.get_sheets();
        const mode = 0 | (config.formulas ? XLSX.with_formulas() : 0);
        sheetsData = sheets.map(name => xlsx.get_sheet_data(name, mode));
    }

    postMessage({
        uid: config.uid || Date.now(),
        type: "ready",
        data: sheetsData,
        styles
    });
}

postMessage({ type:"init" });