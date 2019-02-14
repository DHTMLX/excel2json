Excell to JSON
--------------

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