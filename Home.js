import { writeJsonToExcel } from './excelHelper.js';

Office.onReady(() => {
    document.getElementById("callApiBtn").onclick = async () => {
        const method = document.getElementById("httpMethod").value;
        let url = document.getElementById("baseUrl").value.trim();
        const pathParamsText = document.getElementById("pathParams").value.trim();
        const queryParamsText = document.getElementById("queryParams").value.trim();
        const bodyText = document.getElementById("bodyContent").value.trim();

        try {
            if (pathParamsText) {
                const pathParts = pathParamsText.split(",").map(p => p.trim()).filter(p => p.length > 0);
                if (pathParts.length > 0) {
                    if (!url.endsWith("/")) url += "/";
                    url += pathParts.join("/");
                }
            }

            if (queryParamsText) {
                const queryParams = JSON.parse(queryParamsText);
                const queryString = new URLSearchParams(queryParams).toString();
                url += (url.includes("?") ? "&" : "?") + queryString;
            }
        } catch (err) {
            return console.log("Parametreler geçersiz JSON formatında olmalı.");
        }

        const fetchOptions = { method };
        if (method === "POST" || method === "PUT") {
            try {
                fetchOptions.body = JSON.stringify(JSON.parse(bodyText));
                fetchOptions.headers = { "Content-Type": "application/json" };
            } catch {
                return console.log("Body JSON formatında olmalı.");
            }
        }

        try {
            const response = await fetch(url, fetchOptions);
            const json = await response.json();

            await writeJsonToExcel(json);
        } catch (err) {
            console.log(err)
        }
    };
});
