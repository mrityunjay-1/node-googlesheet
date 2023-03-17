const googleapis = require("googleapis");
const sheets = googleapis.google.sheets("v4");

const GSUtils = require("./utils");

const utils: any = new GSUtils();

class GoogleSheet {

    valueInputOption: string;
    sheets: any;
    auth: any;
    utils: Utils;
    spreadsheetId: string | undefined;

    constructor(auth_type: string, auth_data: any) {
        let oauth;

        switch (auth_type) {

            case "client_id_client_secret":
                oauth = new googleapis.google.auth.OAuth2(auth_data.client_id, auth_data.client_secret, auth_data.redirect_uri);
                oauth.setCredentials({ refresh_token: auth_data.refresh_token });
                break;

            case "google-service-account":

                oauth = new googleapis.google.JWT(
                    auth_data.client_email,
                    null,
                    auth_data.private_key,
                    [
                        "https://www.googleapis.com/auth/spreadsheets",
                        "https://www.googleapis.com/auth/drive"
                    ]
                );

                break;

            default:
                throw new Error("please do authentication first...");
        }

        this.valueInputOption = "RAW";
        this.sheets = sheets;
        this.auth = oauth;
        this.utils = utils;
    }


    getCellName([row, col]: [number, number]) {
        let column = "";
        let col_repeatation = Math.floor(col / 27);
        if (col_repeatation > 0) column = column + String.fromCharCode(64 + col_repeatation);
        let remainder = col % 26;
        if (remainder === 0) remainder = 26;
        column = column + (String.fromCharCode(64 + remainder));
        return column + "-" + row;
    }

    async createSpreadSheet(spreadsheetName = "SpreadSheet1") {
        try {

            let req = {
                auth: this.auth,
                resource: {
                    properties: {
                        title: spreadsheetName

                    }
                }
            }

            const res = await sheets.spreadsheets.create(req);

            return res;


        } catch (err) {
            console.log(err);
        }
    }

    async setSpreadSheetToWorkWith(spreadsheetId: string) {

        try {

            if (!spreadsheetId) throw new Error("please pass spreadsheetId by calling this function named: setSpreadSheetToWorkWith");

            this.spreadsheetId = spreadsheetId;

        } catch (err) {
            console.log(err);
        }

    }

    async addSheet(title = "Sheet-" + (new Date().getTime()), index = 0) {
        try {

            let request = {
                auth: this.auth,
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [
                        {
                            addSheet: {
                                properties: {
                                    title,
                                    index
                                }
                            }
                        }
                    ]
                }
            }

            const res = (await sheets.spreadsheets.batchUpdate(request)).data;

            return res;

        } catch (err: any) {
            console.log(err);
            return err.message;
        }
    }

    async append(value: any, range = "Sheet1") {

        let request = {
            auth: this.auth,
            valueInputOption: "RAW",
            range,
            spreadsheetId: this.spreadsheetId,
            resource: {
                values: [
                    value
                ]
            }
        };

        const res = (await sheets.spreadsheets.values.append(request)).data;

        return res;

    }

    async read(range = "Sheet1", raw_data = false) {
        try {

            const request = {
                spreadsheetId: this.spreadsheetId,
                range,
                auth: this.auth
            }

            const res = (await sheets.spreadsheets.values.get(request)).data;

            if (!raw_data) {
                let formattedData = utils.formatDataLikeJson(res.values);
                return formattedData;
            }

            return res.values;

        } catch (err) {
            console.log(err);
        }
    }

    async searchText(text: string, range = "Sheet1") {

        const data = await this.read(range, true);

        let foundRecords = [];

        for (let i = 0; i < data.length; ++i) {
            if (data[i].includes(text)) {
                let index = data[i].indexOf(text);
                let cellname = this.getCellName([i + 1, index + 1]);
                foundRecords.push(cellname);
            }
        }

        return foundRecords;

    }

    async update(data = [], range = "Sheet1") {
        try {

            let request = {
                spreadsheetId: this.spreadsheetId,
                range,
                valueInputOption: this.valueInputOption,
                resource: {
                    values: [
                        data
                    ]
                },
                auth: this.auth
            }

            const res = (await sheets.spreadsheets.values.update(request)).data;

            return res;

        } catch (err) {
            console.log(err);
        }
    }

    async clear(range = "Sheet1") {
        try {

            let request = {
                spreadsheetId: this.spreadsheetId,
                range,
                resource: {},
                auth: this.auth
            }

            const res = (await sheets.spreadsheets.values.clear(request)).data;

            return res;

        } catch (err) {
            console.log(err);
        }
    }

    async getSheetId(sheetName = "Sheet1") {
        try {

            let request = {
                spreadsheetId: this.spreadsheetId,
                auth: this.auth
            };

            const sheetData = (await sheets.spreadsheets.get(request)).data;

            if (sheetData?.sheets?.length > 0) {
                const sheetFound = sheetData?.sheets?.find((sheet: any) => (sheet.properties.title === sheetName));

                if (sheetFound) {
                    return sheetFound.properties.sheetId;
                }
            }

            return -1;

        } catch (err) {
            console.log(err);
        }
    }

    async deleteRowsOrColumn({
        dimension = "ROWS",
        sheetName = "",
        sheetId = undefined,
        indexesToBeDelete
    }: {
        dimension: string,
        sheetName: string,
        sheetId: any,
        indexesToBeDelete: {
            startIndex: number, endIndex: number
        }
    }) {
        try {

            if (dimension !== "COLUMN") {
                throw new Error("dimension value should be either ROWS or COLUMN");
            }

            if (!(indexesToBeDelete && indexesToBeDelete.startIndex && indexesToBeDelete.endIndex)) {
                throw new Error("please pass indexes values as {startIndex: 0, endIndex: 1}");
            }

            if (!isNaN(sheetId)) {
                sheetId = Number(sheetId);
            }

            if (!sheetId && !sheetName) {
                throw new Error("please pass sheetId or sheetName while calling the deleteRowsOrColumn method...");
            }

            if (!sheetId && sheetName) {
                sheetId = await this.getSheetId(sheetName);

                if (sheetId < 0) {
                    throw new Error("No sheet id found with sheetname: " + sheetName);
                }
            }

            let request = {
                spreadsheetId: this.spreadsheetId,
                resource: {
                    requests: [
                        {
                            deleteDimension: {
                                range: {
                                    ...indexesToBeDelete,
                                    sheetId,
                                    dimension
                                }
                            }
                        }
                    ]
                },
                auth: this.auth
            };

            const deleteRes = (await sheets.spreadsheets.batchUpdate(request)).data;

            return deleteRes;

        } catch (err) {
            console.log(err);
        }
    }

}

module.exports = {
    GoogleSheet
}