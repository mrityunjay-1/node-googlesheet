"use strict";
var Utils = /** @class */ (function () {
    function Utils() {
    }
    Utils.prototype.formatDataLikeJson = function (data) {
        var res = [], obj = {}, headers = data[0];
        for (var i = 1; i < data.length; ++i) {
            for (var j = 0; j < headers.length; ++j) {
                obj["".concat(headers[j])] = data[i][j];
            }
            res.push(obj);
            obj = {};
        }
        return res;
    };
    return Utils;
}());
module.exports = Utils;
