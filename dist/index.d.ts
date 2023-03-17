declare const googleapis: any;
declare const sheets: any;
declare const GSUtils: any;
declare const utils: any;
declare class GoogleSheet {
    valueInputOption: string;
    sheets: any;
    auth: any;
    utils: Utils;
    spreadsheetId: string | undefined;
    constructor(auth_type: string, auth_data: any);
    getCellName([row, col]: [number, number]): string;
    createSpreadSheet(spreadsheetName?: string): Promise<any>;
    setSpreadSheetToWorkWith(spreadsheetId: string): Promise<void>;
    addSheet(title?: string, index?: number): Promise<any>;
    append(value: any, range?: string): Promise<any>;
    read(range?: string, raw_data?: boolean): Promise<any>;
    searchText(text: string, range?: string): Promise<string[]>;
    update(data?: never[], range?: string): Promise<any>;
    clear(range?: string): Promise<any>;
    getSheetId(sheetName?: string): Promise<any>;
    deleteRowsOrColumn({ dimension, sheetName, sheetId, indexesToBeDelete }: {
        dimension: string;
        sheetName: string;
        sheetId: any;
        indexesToBeDelete: {
            startIndex: number;
            endIndex: number;
        };
    }): Promise<any>;
}
