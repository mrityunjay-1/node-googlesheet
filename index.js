const googleapis = require("googleapis");
const sheets = googleapis.google.sheets("v4");

class GoogleSheet {

    constructor(auth_type, auth_data) {
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
    }

    formatDataLikeJson(data) {

        let res = [], obj = {}, headers = data[0];

        for (let i = 1; i < data.length; ++i) {

            for (let j = 0; j < headers.length; ++j) {
                obj[`${headers[j]}`] = data[i][j];
            }

            res.push(obj);

            obj = {};
        }

        return res;
    }

    getCellName([row, col]) {
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

    async setSpreadSheetToWorkWith(spreadsheetId) {

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

        } catch (err) {
            console.log(err);
            return err.message;
        }
    }

    async append(value, range = "Sheet1") {

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
                let formattedData = this.formatDataLikeJson(res.values);
                return formattedData;
            }

            return res.values;

        } catch (err) {
            console.log(err);
        }
    }

    async searchText(text, range = "Sheet1") {

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

}

module.exports = {
    GoogleSheet
}