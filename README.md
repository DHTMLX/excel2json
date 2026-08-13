Excel2json
--------------

Excel2json is a Rust and WebAssembly-based library that allows easily converting files in Excel format to JSON files.

[![npm version](https://badge.fury.io/js/excel2json-wasm.svg)](https://badge.fury.io/js/excel2json-wasm)

### How to build

#### wasm

```
cargo install wasm-pack
wasm-pack build
```


### How to use via npm

- install the module

```js
yarn add excel2json-wasm
```
- import and use the module

```js
// worker.js
import {convert} "excel2json-wasm";
// convert int-array, works in sync mode
const json = convertArray(Int8Array, optional_config);
// converts file object or int-array in async mode
convert(file_object_or_typed_array, optional_config).then(json => {
    // do something
});
```


### How to use from CDN

CDN links are the following:

- https://cdn.dhtmlx.com/libs/excel2json/1.5/worker.js
- https://cdn.dhtmlx.com/libs/excel2json/1.5/module.js
- https://cdn.dhtmlx.com/libs/excel2json/1.5/excel2json_wasm_bg.wasm


You can import and use lib dynamically like

```js
const convert = import("https://cdn.dhtmlx.com/libs/excel2json/1.5/module.js");
const blob = convert(json_data_to_export);
```

or use it as web worker

```js
var url = window.URL.createObjectURL(new Blob([
    "importScripts('https://cdn.dhtmlx.com/libs/excel2json/1.5/worker.js');"
], { type: "text/javascript" }));

// you need to server worker from the same domain as the main script
var worker = new Worker("./worker.js");
worker.addEventListener("message", ev => {
    if (ev.data.type === "ready"){
        const json = ev.data.data;
        // do something
    }
});
worker.postMessage({
    type:"convert",
    data: file_object
});
worker.addEventListener("message", e => {
    if (e.data.init){
        // worker is ready
    }
})
```

if you want to load worker script from CDN and not from your domain it requires a more complicated approach, as you need to catch the moment when service inside of the worker will be fully initialized

```js
var url = window.URL.createObjectURL(new Blob([
    "importScripts('https://cdn.dhtmlx.com/libs/excel2json/1.5/worker.js');"
], { type: "text/javascript" }));

var worker = new Promise((res) => {
    const x = Worker(url);
    worker.addEventListener("message", ev => {
        if (ev.data.type === "ready"){
            const json = ev.data.data;
            // do something with result
        } else if (ev.data.type === "init"){
            // service is ready
            res(x);
        }
    });
});

worker.then(x => x.postMessage({
    type:"convert",
    data: file_object
}));
```

#### Export formulas


```js
const json = convert(data, { formulas:true });
```

or

```js
worker.postMessage({
    type: "convert",
    data: file_object_or_typed_array,
    formulas: true
});
```

### Output format

```ts
interface IConvertMessageData {
    uid?: string;
    data: Uint8Array | File;
    sheet?: string;
    styles?: boolean;
    wasmPath?: string;
}

interface IReadyMessageData {
    uid: string;
    data: ISheetData[];
    styles: IStyles[];
}

interface ISheetData {
    name: string;
    cols: IColumnData[];
    rows: IRowData[];
    cells: IDataCell[][];   // null for empty cell

    merged: IMergedCell[];
}

interface IMergedCell {
    from: IDataPoint;
    to: IDataPoint;
}

interface IDataPoint {
    column: number;
    row: number;
}

interface IColumnData {
    width: number;
}

interface IRowData {
    height: number;
}

interface IDataCell{
    v: string;
    s: number:
    formula?: IFormula;
}

// present when formulas:true; v keeps the original legacy behavior:
// cells with formula text in XLSX use "=...", shared formula followers keep cached values
interface IFormula {
    type: "normal" | "shared";
    role?: "master" | "follower";
    si?: string;
    ref?: string;
    value?: string;
}

interface IStyle {
    fontSize?: string;
    fontFamily?: string;

    background?: string;
    color?: string;

    fontWeight?: string;
    fontStyle?: string;
    textDecoration?: string;

    align?: string;
    verticalAlign?: string;

    borderLeft?: string;
    borderTop?: string;
    borderBottom?: string;
    borderRight?: string;

    format?: string;
}
```

### Versions

- **1.6.0** — Locked cells; shared formula import with per-cell formula metadata (`formulas: true`).
- **1.5.0** — Text wrapping; row height import.
- **1.4.0** — Frozen rows/columns, hidden rows/columns, data validation, and hyperlinks.
- **1.3.0** — Styles on empty cells; skip empty borders; `textAlign` renamed to `align`; phonetic runs ignored in shared strings.
- **1.2.0** — Formula export (`formulas: true`); formatted text parsing; ESM `convert` / `convertArray` API.
- **1.0.2** — Default border color.
- **1.0.1** — CDN links.
- **1.0.0** — Initial WASM Excel-to-JSON converter.

### License

MIT
