import init, { XLSX } from "../pkg/excel2json_wasm.js";

let initialized = false;

async function ensureInit() {
    if (!initialized) {
        await init();
        initialized = true;
    }
}

export { XLSX };

export async function convertArray(jsonData, config = {}) {
    await ensureInit();

    const getStyles = config.styles === undefined ? true : config.styles;
    const xlsx = XLSX.new(jsonData);
    const styles = getStyles ? xlsx.get_styles() : null;

    let data;
    if (config.sheet) {
        data = [xlsx.get_sheet_data(config.sheet)];
    } else {
        const sheets = xlsx.get_sheets();
        const mode = 0 | (config.formulas ? XLSX.with_formulas() : 0);
        data = sheets.map(name => xlsx.get_sheet_data(name, mode));
    }

    return { data, styles };
}

export async function convert(jsonData, config = {}) {
    await ensureInit();

    if (jsonData instanceof File) {
        return new Promise((res) => {
            const reader = new FileReader();
            reader.readAsArrayBuffer(jsonData);
            reader.onload = async (e) => {
                const result = await convertArray(new Uint8Array(e.target.result), config);
                res(result);
            };
        });
    } else {
        return await convertArray(jsonData, config);
    }
}
