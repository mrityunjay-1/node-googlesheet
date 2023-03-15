# node-googlesheet

### Installation
```
npm install node-googlesheet
```

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

[Submit Your Issue / Bug / Feedback](https://docs.google.com/forms/d/e/1FAIpQLSc4IcMovlwocjoUnVR3tY6aEC9UDPmpgWvsk0tfGUwTLNhdFw/viewform)