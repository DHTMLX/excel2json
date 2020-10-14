Excel2json
--------------

Excel2json is a Rust and WebAssembly-based library that allows easily converting files in Excel format to JSON files.

[![npm version](https://badge.fury.io/js/excel2json-wasm.svg)](https://badge.fury.io/js/excel2json-wasm) 

### How to build

#### wasm

```
cargo install wasm-pack
wasm-pack build --target web
```

#### js

```
yarn install
yarn build
```

### How to use via npm 

- install the module

```js
yarn add excel2json-wasm
```
- import the module

```js
// worker.js
import("excel2json-wasm")
```

- use the module in the app

```js
// app.js
const worker = new Worker("worker.js");

// convert excel file to json
worker.postMessage({
    type: "convert",
    data: file_object_or_typed_array
});

worker.addEventListener("message", e => {
    if (e.data.type === "ready"){
        const data = e.data.data;
        const styles = e.data.styles;

        //json data is ready
        console.log(data, styles)
    }
});
```

#### Export formulas

```
worker.postMessage({
    type: "convert",
    data: file_object_or_typed_array,
    formulas: true
});
```

### How to use from CDN

CDN links are the following:

- https://cdn.dhtmlx.com/libs/excel2json/1.1/worker.js 
- https://cdn.dhtmlx.com/libs/excel2json/1.1/lib.wasm

In case you use build system like webpack, it is advised to wrap the link to CDN source into a blob object to avoid possible 
breakdowns:

```js
var url = window.URL.createObjectURL(new Blob([
    "importScripts('https://cdn.dhtmlx.com/libs/excel2json/1.1/worker.js');"
], { type: "text/javascript" }));

var worker = new Worker(url);
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
