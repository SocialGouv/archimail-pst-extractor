import iconv from "iconv-lite";

import type { NameAndAddress } from "./type";

//Default character set
const defaultCharset: BufferEncoding = "utf-8"; //to use if charset=... is missing

//Gets the character encoding name for iconv, e.g. 'iso-8859-2' -> 'iso88592'
function getCharsetName(charset: BufferEncoding) {
    return charset.toLowerCase().replace(/[^0-9a-z]/g, "");
}

//Word-wrap the string 's' to 'i' chars per row
export function wrap(s: string, i: number): string {
    const a = [];
    do {
        a.push(s.substring(0, i));
    } while ((s = s.substring(i, s.length)) != "");
    return a.join("\r\n");
}

//Gets the boundary name
export const getBoundary = (contentType: string): string | undefined => {
    const match = /boundary="?(.+?)"?(\s*;[\s\S]*)?$/g.exec(contentType);
    return match?.[1];
};

//Gets character set name, e.g. contentType='.....charset="iso-8859-2"....'
export const getCharset = (contentType: string): BufferEncoding | undefined => {
    const match = /charset\s*=\W*([\w-]+)/g.exec(contentType);
    return match?.[1] as BufferEncoding | undefined;
};

//Gets name and e-mail address from a string, e.g. "PayPal" <noreply@paypal.com> => { name: "PayPal", email: "noreply@paypal.com" }
export const getEmailAddress = (raw: string): NameAndAddress[] => {
    const list: NameAndAddress[] = [];

    //Split around ',' char
    //var parts = raw.split(/,/g); //Will also split ',' inside the quotes
    //var parts = raw.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g); //Ignore ',' within the double quotes
    const parts = raw.match(/("[^"]*")|[^,]+/g); //Ignore ',' within the double quotes

    if (!parts) {
        return list;
    }

    for (let i = 0; i < parts.length; i++) {
        const address = {} as NameAndAddress;

        //Quoted name but without the e-mail address
        if (/^".*"$/g.test(parts[i])) {
            address.name = unquoteString(parts[i]).replace(/"/g, "").trim();
            i++; //Shift to another part to capture e-mail address
        }

        const regex = /^(.*?)(\s*<(.*?)>)$/g;
        const match = regex.exec(parts[i]);
        if (match) {
            const name = unquoteString(match[1]).replace(/"/g, "").trim();
            if (name.length) {
                address.name = name;
            }
            address.email = match[3].trim();
            list.push(address);
        } else {
            //E-mail address only (without the name)
            address.email = parts[i].trim();
            list.push(address);
        }
    }

    return list; //Multiple e-mail addresses as array
};

//Builds e-mail address string, e.g. { name: "PayPal", email: "noreply@paypal.com" } => "PayPal" <noreply@paypal.com>
export const toEmailAddress = (
    data?: NameAndAddress | NameAndAddress[] | string
): string => {
    let ret = "";
    if (typeof data === "string") {
        ret = data;
    } else if (typeof data === "object") {
        if (Array.isArray(data)) {
            for (const { name, email } of data) {
                ret += ret.length ? ", " : "";
                if (name) {
                    ret += `"${name}"`;
                }
                if (email) {
                    ret += `${ret.length ? " " : ""}<${email}>`;
                }
            }
        } else {
            if (data.name) {
                ret += `"${data.name}"`;
            }
            if (data.email) {
                ret += `${ret.length ? " " : ""}<${data.email}>`;
            }
        }
    }
    return ret;
};

//Decodes "quoted-printable"
export const unquotePrintable = (
    str: string,
    iconvCharset?: string
): string => {
    //Convert =0D to '\r', =20 to ' ', etc.
    if (!iconvCharset || iconvCharset === "utf8" || iconvCharset === "utf-8") {
        return str
            .replace(
                /=([\w\d]{2})=([\w\d]{2})=([\w\d]{2})/gi,
                (_, p1: string, p2: string, p3: string) =>
                    Buffer.from([
                        parseInt(p1, 16),
                        parseInt(p2, 16),
                        parseInt(p3, 16),
                    ]).toString("utf8")
            )
            .replace(
                /=([\w\d]{2})=([\w\d]{2})/gi,
                (_, p1: string, p2: string) =>
                    Buffer.from([parseInt(p1, 16), parseInt(p2, 16)]).toString(
                        "utf8"
                    )
            )
            .replace(/=([\w\d]{2})/gi, (_, p1: string) =>
                String.fromCharCode(parseInt(p1, 16))
            )
            .replace(/=\r?\n/gi, ""); //Join line
    } else {
        return str
            .replace(
                /=([\w\d]{2})=([\w\d]{2})/gi,
                (_, p1: string, p2: string) =>
                    iconv.decode(
                        Buffer.from([parseInt(p1, 16), parseInt(p2, 16)]),
                        iconvCharset
                    )
            )
            .replace(/=([\w\d]{2})/gi, (_, p1: string): string =>
                iconv.decode(Buffer.from([parseInt(p1, 16)]), iconvCharset)
            )
            .replace(/=\r?\n/gi, ""); //Join line
    }
};

//Decodes string by detecting the charset
export const unquoteString = (str: string): string => {
    const regex = /=\?([^?]+)\?(B|Q)\?(.+?)(\?=)/gi;
    const match = regex.exec(str);
    if (!match) return str;

    const charset = getCharsetName(
        (match[1] as BufferEncoding | undefined) ?? defaultCharset
    ); //eq. match[1] = 'iso-8859-2'; charset = 'iso88592'
    const type = match[2].toUpperCase();
    const value = match[3];
    if (type === "B") {
        //Base64
        if (charset === "utf8") {
            return Buffer.from(value.replace(/\r?\n/g, ""), "base64").toString(
                "utf8"
            );
        } else {
            return iconv.decode(
                Buffer.from(value.replace(/\r?\n/g, ""), "base64"),
                charset
            );
        }
    } else if (type === "Q") {
        //Quoted printable
        return unquotePrintable(value, charset);
    }

    return str;
};

//Decodes string like =?UTF-8?B?V2hhdOKAmXMgeW91ciBvbmxpbmUgc2hvcHBpbmcgc3R5bGU/?= or =?UTF-8?Q?...?=
export const unquoteUTF8 = (str: string): string => {
    const regex = /=\?UTF-8\?(B|Q)\?(.+?)(\?=)/gi;
    const match = regex.exec(str);
    if (match) {
        const type = match[1].toUpperCase();
        const value = match[2];
        if (type === "B") {
            //Base64
            return Buffer.from(value.replace(/\r?\n/g, ""), "base64").toString(
                "utf8"
            );
        } else if (type === "Q") {
            //Quoted printable
            return unquotePrintable(value);
        }
    }
    return str;
};
