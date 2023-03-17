# node-googlesheet
Accelerate your productivity with this easy-to-use library for Node js

### Installation
```
npm install node-googlesheet
```
The Most Stable Version is "v1.1.1" so kindly upgrade to same.

### Sample Code to read data

```js
const { GoogleSheet } = require("node-googlesheet");

// update below variables value with your credential details...
const client_id = "";
const client_secret = "";
const refresh_token = "";
const redirect_uri = "";

const gs = new GoogleSheet("client_id_client_secret", {
    client_id, client_secret, refresh_token, redirect_uri
});

gs.setSpreadSheetToWorkWith(""); // For Ex: "1yFKedpLbKZ5MCHJaoXIz_9oc2_qoXZkUEyxTQQEfXtE"

(
    async () => {
        try {

            const data = await gs.read();

            console.log("Data : ", data);

        } catch (err) {
            console.log("err: ", err);
        }
    }
)();

```

### Add Data
```js
const res = await gs.append([2, "MK", 20]); 
// append method takes two parameter having an array (array of values repesenting each cell) and second is string (range where you want to append the data)

console.log("res : ", res);
```

### Update Data
```js
const res = await gs.update([2, "MK", 20], "Sheet!A1"); 
// update method takes two parameter having an array (array of values repesenting each cell) and second is string (range where you want to append the data)

console.log("res : ", res);
```

### Delete Rows / Column

```js
const res = await gs.deleteRowsOrColumn({
    sheetName: "Sheet1",
    indexesToBeDelete: {
        startIndex: 1, endIndex: 2
    }
});

console.log("res : ", res);

```

[Submit Your Issue / Bug / Feedback](https://docs.google.com/forms/d/e/1FAIpQLSc4IcMovlwocjoUnVR3tY6aEC9UDPmpgWvsk0tfGUwTLNhdFw/viewform)