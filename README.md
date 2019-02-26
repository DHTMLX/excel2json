Excell to JSON
--------------

[![npm version](https://badge.fury.io/js/excel2json-wasm.svg)](https://badge.fury.io/js/excel2json-wasm) 

### How to build

```
yarn install
// build js code
yarn build
// rebuild wasm code
// rust toolchain is required
yarn build-wasm
```

### CDN links

- https://cdn.dhtmlx.com/libs/excel2json/1.0/worker.js 
- https://cdn.dhtmlx.com/libs/excel2json/1.0/lib.wasm

### How to use

```js
// worker.js
import("excel2json-wasm")
```

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

### License 

MIT

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