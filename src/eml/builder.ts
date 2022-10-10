import { randomUUID } from "crypto";

import type { EmlStruct, InternalEmlHeaders } from "./type";
import { getBoundary, toEmailAddress, wrap } from "./utils";

export interface EmlStringifyOptions {
    /**
     * Line Feed, used for end-of-line. Usualy `require("os").EOL` on NodeJS.
     *
     * @default \r\n
     */
    lf?: string;
}

/**
 * Convert EML struct to EML string.
 */
export const stringify = (
    data: EmlStruct,
    options?: EmlStringifyOptions
): string => {
    let eml = "";
    const EOL = options?.lf ?? "\r\n"; //End-of-line

    const internalHeaders = {} as InternalEmlHeaders;
    if (!data.headers) {
        data.headers = {};
    }

    if (typeof data.subject === "string") {
        internalHeaders.Subject = data.subject;
    }

    if (typeof data.from !== "undefined") {
        internalHeaders.From =
            typeof data.from === "string"
                ? data.from
                : toEmailAddress(data.from);
    }

    internalHeaders.To =
        typeof data.to === "string" ? data.to : toEmailAddress(data.to);

    if (typeof data.cc !== "undefined") {
        internalHeaders.Cc =
            typeof data.cc === "string" ? data.cc : toEmailAddress(data.cc);
    }

    let boundary = `----=${randomUUID()}`;
    if (typeof internalHeaders["Content-Type"] === "undefined") {
        internalHeaders[
            "Content-Type"
        ] = `multipart/mixed;${EOL}boundary="${boundary}"`;
    } else {
        const name = getBoundary(internalHeaders["Content-Type"]);
        if (name) {
            boundary = name;
        }
    }

    data.headers = { ...data.headers, ...internalHeaders };

    //Build headers
    for (const [key, value] of Object.entries(data.headers)) {
        if (typeof value === "undefined") {
            continue; //Skip missing headers
        } else if (typeof value === "string") {
            eml += `${key}: ${value.replace(/\r?\n/g, `${EOL}  `)}${EOL}`;
        } else {
            //Array
            for (const v of value) {
                eml += `${key}: ${v.replace(/\r?\n/g, `${EOL}  `)}${EOL}`;
            }
        }
    }

    //Start the body
    eml += EOL;

    //Plain text content
    if (data.text) {
        eml += `--${boundary}${EOL}`;
        eml += `Content-Type: text/plain; charset=utf-8${EOL}`;
        eml += EOL;
        eml += data.text;
        eml += EOL + EOL;
    }

    //HTML content
    if (data.html) {
        eml += `--${boundary}${EOL}`;
        eml += `Content-Type: text/html; charset=utf-8${EOL}`;
        eml += EOL;
        eml += data.html;
        eml += EOL + EOL;
    }

    //Append attachments
    if (data.attachments) {
        for (let i = 0; i < data.attachments.length; i++) {
            const attachment = data.attachments[i];
            eml += `--${boundary}${EOL}`;
            eml += `Content-Type: ${
                attachment.contentType || "application/octet-stream" // eslint-disable-line @typescript-eslint/prefer-nullish-coalescing -- prevent empty string
            }${EOL}`;
            eml += `Content-Transfer-Encoding: base64${EOL}`;
            eml += `Content-Disposition: ${
                attachment.inline ? "inline" : "attachment"
            }; filename="${
                attachment.filename || // eslint-disable-line @typescript-eslint/prefer-nullish-coalescing -- prevent empty string
                attachment.name || // eslint-disable-line @typescript-eslint/prefer-nullish-coalescing -- prevent empty string
                `attachment_${i + 1}`
            }"${EOL}`;
            if (attachment.cid) {
                eml += `Content-ID: <${attachment.cid}>${EOL}`;
            }
            eml += EOL;
            if (typeof attachment.data === "string") {
                const content = Buffer.from(attachment.data).toString("base64");
                eml += wrap(content, 76) + EOL;
            } else {
                //Buffer
                const content = attachment.data.toString("base64");
                eml += wrap(content, 76) + EOL;
            }
            eml += EOL;
        }
    }

    //Finish the boundary
    eml += `--${boundary}--${EOL}`;

    return eml;
};
