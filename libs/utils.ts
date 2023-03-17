class Utils {

    constructor() { }

    formatDataLikeJson(data: [[]]) {

        let res = [], obj = {}, headers: [] = data[0];

        for (let i = 1; i < data.length; ++i) {

            for (let j = 0; j < headers.length; ++j) {
                obj[`${headers[j]}`] = data[i][j];
            }

            res.push(obj);

            obj = {};
        }

        return res;
    }

}

module.exports = Utils;