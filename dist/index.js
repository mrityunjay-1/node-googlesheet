"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var googleapis = require("googleapis");
var sheets = googleapis.google.sheets("v4");
var GSUtils = require("./utils");
var utils = new GSUtils();
var GoogleSheet = /** @class */ (function () {
    function GoogleSheet(auth_type, auth_data) {
        var oauth;
        switch (auth_type) {
            case "client_id_client_secret":
                oauth = new googleapis.google.auth.OAuth2(auth_data.client_id, auth_data.client_secret, auth_data.redirect_uri);
                oauth.setCredentials({ refresh_token: auth_data.refresh_token });
                break;
            case "google-service-account":
                oauth = new googleapis.google.JWT(auth_data.client_email, null, auth_data.private_key, [
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]);
                break;
            default:
                throw new Error("please do authentication first...");
        }
        this.valueInputOption = "RAW";
        this.sheets = sheets;
        this.auth = oauth;
        this.utils = utils;
    }
    GoogleSheet.prototype.getCellName = function (_a) {
        var row = _a[0], col = _a[1];
        var column = "";
        var col_repeatation = Math.floor(col / 27);
        if (col_repeatation > 0)
            column = column + String.fromCharCode(64 + col_repeatation);
        var remainder = col % 26;
        if (remainder === 0)
            remainder = 26;
        column = column + (String.fromCharCode(64 + remainder));
        return column + "-" + row;
    };
    GoogleSheet.prototype.createSpreadSheet = function (spreadsheetName) {
        if (spreadsheetName === void 0) { spreadsheetName = "SpreadSheet1"; }
        return __awaiter(this, void 0, void 0, function () {
            var req, res, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        req = {
                            auth: this.auth,
                            resource: {
                                properties: {
                                    title: spreadsheetName
                                }
                            }
                        };
                        return [4 /*yield*/, sheets.spreadsheets.create(req)];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, res];
                    case 2:
                        err_1 = _a.sent();
                        console.log(err_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GoogleSheet.prototype.setSpreadSheetToWorkWith = function (spreadsheetId) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    if (!spreadsheetId)
                        throw new Error("please pass spreadsheetId by calling this function named: setSpreadSheetToWorkWith");
                    this.spreadsheetId = spreadsheetId;
                }
                catch (err) {
                    console.log(err);
                }
                return [2 /*return*/];
            });
        });
    };
    GoogleSheet.prototype.addSheet = function (title, index) {
        if (title === void 0) { title = "Sheet-" + (new Date().getTime()); }
        if (index === void 0) { index = 0; }
        return __awaiter(this, void 0, void 0, function () {
            var request, res, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        request = {
                            auth: this.auth,
                            spreadsheetId: this.spreadsheetId,
                            resource: {
                                requests: [
                                    {
                                        addSheet: {
                                            properties: {
                                                title: title,
                                                index: index
                                            }
                                        }
                                    }
                                ]
                            }
                        };
                        return [4 /*yield*/, sheets.spreadsheets.batchUpdate(request)];
                    case 1:
                        res = (_a.sent()).data;
                        return [2 /*return*/, res];
                    case 2:
                        err_2 = _a.sent();
                        console.log(err_2);
                        return [2 /*return*/, err_2.message];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GoogleSheet.prototype.append = function (value, range) {
        if (range === void 0) { range = "Sheet1"; }
        return __awaiter(this, void 0, void 0, function () {
            var request, res;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        request = {
                            auth: this.auth,
                            valueInputOption: "RAW",
                            range: range,
                            spreadsheetId: this.spreadsheetId,
                            resource: {
                                values: [
                                    value
                                ]
                            }
                        };
                        return [4 /*yield*/, sheets.spreadsheets.values.append(request)];
                    case 1:
                        res = (_a.sent()).data;
                        return [2 /*return*/, res];
                }
            });
        });
    };
    GoogleSheet.prototype.read = function (range, raw_data) {
        if (range === void 0) { range = "Sheet1"; }
        if (raw_data === void 0) { raw_data = false; }
        return __awaiter(this, void 0, void 0, function () {
            var request, res, formattedData, err_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        request = {
                            spreadsheetId: this.spreadsheetId,
                            range: range,
                            auth: this.auth
                        };
                        return [4 /*yield*/, sheets.spreadsheets.values.get(request)];
                    case 1:
                        res = (_a.sent()).data;
                        if (!raw_data) {
                            formattedData = utils.formatDataLikeJson(res.values);
                            return [2 /*return*/, formattedData];
                        }
                        return [2 /*return*/, res.values];
                    case 2:
                        err_3 = _a.sent();
                        console.log(err_3);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GoogleSheet.prototype.searchText = function (text, range) {
        if (range === void 0) { range = "Sheet1"; }
        return __awaiter(this, void 0, void 0, function () {
            var data, foundRecords, i, index, cellname;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.read(range, true)];
                    case 1:
                        data = _a.sent();
                        foundRecords = [];
                        for (i = 0; i < data.length; ++i) {
                            if (data[i].includes(text)) {
                                index = data[i].indexOf(text);
                                cellname = this.getCellName([i + 1, index + 1]);
                                foundRecords.push(cellname);
                            }
                        }
                        return [2 /*return*/, foundRecords];
                }
            });
        });
    };
    GoogleSheet.prototype.update = function (data, range) {
        if (data === void 0) { data = []; }
        if (range === void 0) { range = "Sheet1"; }
        return __awaiter(this, void 0, void 0, function () {
            var request, res, err_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        request = {
                            spreadsheetId: this.spreadsheetId,
                            range: range,
                            valueInputOption: this.valueInputOption,
                            resource: {
                                values: [
                                    data
                                ]
                            },
                            auth: this.auth
                        };
                        return [4 /*yield*/, sheets.spreadsheets.values.update(request)];
                    case 1:
                        res = (_a.sent()).data;
                        return [2 /*return*/, res];
                    case 2:
                        err_4 = _a.sent();
                        console.log(err_4);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GoogleSheet.prototype.clear = function (range) {
        if (range === void 0) { range = "Sheet1"; }
        return __awaiter(this, void 0, void 0, function () {
            var request, res, err_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        request = {
                            spreadsheetId: this.spreadsheetId,
                            range: range,
                            resource: {},
                            auth: this.auth
                        };
                        return [4 /*yield*/, sheets.spreadsheets.values.clear(request)];
                    case 1:
                        res = (_a.sent()).data;
                        return [2 /*return*/, res];
                    case 2:
                        err_5 = _a.sent();
                        console.log(err_5);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GoogleSheet.prototype.getSheetId = function (sheetName) {
        var _a, _b;
        if (sheetName === void 0) { sheetName = "Sheet1"; }
        return __awaiter(this, void 0, void 0, function () {
            var request, sheetData, sheetFound, err_6;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _c.trys.push([0, 2, , 3]);
                        request = {
                            spreadsheetId: this.spreadsheetId,
                            auth: this.auth
                        };
                        return [4 /*yield*/, sheets.spreadsheets.get(request)];
                    case 1:
                        sheetData = (_c.sent()).data;
                        if (((_a = sheetData === null || sheetData === void 0 ? void 0 : sheetData.sheets) === null || _a === void 0 ? void 0 : _a.length) > 0) {
                            sheetFound = (_b = sheetData === null || sheetData === void 0 ? void 0 : sheetData.sheets) === null || _b === void 0 ? void 0 : _b.find(function (sheet) { return (sheet.properties.title === sheetName); });
                            if (sheetFound) {
                                return [2 /*return*/, sheetFound.properties.sheetId];
                            }
                        }
                        return [2 /*return*/, -1];
                    case 2:
                        err_6 = _c.sent();
                        console.log(err_6);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GoogleSheet.prototype.deleteRowsOrColumn = function (_a) {
        var _b = _a.dimension, dimension = _b === void 0 ? "ROWS" : _b, _c = _a.sheetName, sheetName = _c === void 0 ? "" : _c, _d = _a.sheetId, sheetId = _d === void 0 ? undefined : _d, indexesToBeDelete = _a.indexesToBeDelete;
        return __awaiter(this, void 0, void 0, function () {
            var request, deleteRes, err_7;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _e.trys.push([0, 4, , 5]);
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
                        if (!(!sheetId && sheetName)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.getSheetId(sheetName)];
                    case 1:
                        sheetId = _e.sent();
                        if (sheetId < 0) {
                            throw new Error("No sheet id found with sheetname: " + sheetName);
                        }
                        _e.label = 2;
                    case 2:
                        request = {
                            spreadsheetId: this.spreadsheetId,
                            resource: {
                                requests: [
                                    {
                                        deleteDimension: {
                                            range: __assign(__assign({}, indexesToBeDelete), { sheetId: sheetId, dimension: dimension })
                                        }
                                    }
                                ]
                            },
                            auth: this.auth
                        };
                        return [4 /*yield*/, sheets.spreadsheets.batchUpdate(request)];
                    case 3:
                        deleteRes = (_e.sent()).data;
                        return [2 /*return*/, deleteRes];
                    case 4:
                        err_7 = _e.sent();
                        console.log(err_7);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    return GoogleSheet;
}());
module.exports = {
    GoogleSheet: GoogleSheet
};
