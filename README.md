Excell to JSON
--------------

### Output format

```ts
interface IConvertMessageData {
	uid?: string;
	data: Uint8Array | File;
	sheet?: string; // search by name, use first sheet if not provided
    styles?: boolean; // true by default
    wasmPath?: string; // use cdn by default
}

interface IReadyMessageData {
	uid: string; // same as incoming uid
	data: ISheetData[];
	styles: IStyles[];
}

interface ISheetData {
	name: string;
	cols: IColumnData[];
	rows: IRowData[];
	cells: IDataCell[][];	// null for empty cell

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
	width: number;	//int, round to px
}

interface IRowData {
	height: number;	//int, round to px
}

interface IDataCell{
	v: string; // value
	s: number: // style index, 0 - default style
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