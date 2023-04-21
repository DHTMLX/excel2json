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

- https://cdn.dhtmlx.com/libs/excel2json/1.2/worker.js 
- https://cdn.dhtmlx.com/libs/excel2json/1.2/module.js 
- https://cdn.dhtmlx.com/libs/excel2json/1.2/excel2json_wasm_bg.wasm


You can import and use lib dynamically like 

```js
const convert = import("https://cdn.dhtmlx.com/libs/excel2json/1.2/module.js");
const blob = convert(json_data_to_export);
```

or use it as web worker

```js
var url = window.URL.createObjectURL(new Blob([
    "importScripts('https://cdn.dhtmlx.com/libs/excel2json/1.2/worker.js');"
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
    "importScripts('https://cdn.dhtmlx.com/libs/excel2json/1.2/worker.js');"
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
}

interface IStyle {
    fontSize?: string;
    fontFamily?: string;

    background?: string;
    color?: string;

    fontWeight?: string;
    fontStyle?: string;
    textDecoration?: string;

    textAlign?: string;
    verticalAlign?: string;

    borderLeft?: string;
    borderTop?: string;
    borderBottom?: string;
    borderRight?: string;

    format?: string;
}
```

### License 

MIT
