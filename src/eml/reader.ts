/* eslint-disable */
// @ts-ignore
// @ts-nocheck
// TODO

import { parse } from "./parser";

/******************************************************************************************
 * Parses EML file content and return user-friendly object.
 * @params eml         EML file content or object from 'parse'
 * @params options     EML parse options
 * @params callback    Callback function(error, data)
 ******************************************************************************************/
export const read = (eml, options, callback) => {
    //Shift arguments
    if (typeof options == "function" && typeof callback == "undefined") {
        callback = options;
        options = null;
    }

    if (typeof callback != "function") {
        callback = function (error, result) {};
    }

    function _read(data) {
        try {
            const result = {};
            if (data.headers.Date) {
                result.date = new Date(data.headers.Date);
            }
            if (data.headers.Subject) {
                result.subject = emlformat.unquoteString(data.headers.Subject);
            }
            if (data.headers.From) {
                result.from = emlformat.getEmailAddress(data.headers.From);
            }
            if (data.headers.To) {
                result.to = emlformat.getEmailAddress(data.headers.To);
            }
            if (data.headers.CC) {
                result.cc = emlformat.getEmailAddress(data.headers.CC);
            }
            if (data.headers.Cc) {
                result.cc = emlformat.getEmailAddress(data.headers.Cc);
            }
            result.headers = data.headers;

            //Appends the boundary to the result
            function _append(headers, content) {
                const contentType = headers["Content-Type"];
                const charset = getCharsetName(
                    emlformat.getCharset(contentType) || defaultCharset
                );
                let encoding =
                    headers["Content-Transfer-Encoding"] ||
                    result.headers["Content-Transfer-Encoding"];
                if (typeof encoding == "string") {
                    encoding = encoding.toLowerCase();
                }
                if (encoding == "base64") {
                    if (contentType.indexOf("gbk") >= 0) {
                        content = new Buffer(
                            iconv.decode(
                                new Buffer(content, "base64"),
                                "gb2312"
                            ),
                            "utf8"
                        );
                    } else {
                        content = Buffer.from(
                            content.replace(/\r?\n/g, ""),
                            "base64"
                        );
                    }
                } else if (encoding == "quoted-printable") {
                    content = emlformat.unquotePrintable(content, charset);
                } else if (
                    charset != "utf8" &&
                    encoding &&
                    (encoding.startsWith("binary") ||
                        encoding.startsWith("8bit"))
                ) {
                    //"8bit", "binary", "8bitmime", "binarymime"
                    content = iconv.decode(
                        Buffer.from(content, "binary"),
                        charset
                    );
                }
                if (
                    !result.html &&
                    contentType &&
                    contentType.indexOf("text/html") >= 0
                ) {
                    if (typeof content != "string") {
                        //content = content.toString("utf8");
                        content = iconv.decode(Buffer.from(content), charset);
                    }
                    //Message in HTML format
                    result.html = content;
                } else if (
                    !result.text &&
                    contentType &&
                    contentType.indexOf("text/plain") >= 0
                ) {
                    if (typeof content != "string") {
                        //content = content.toString("utf8");
                        content = iconv.decode(Buffer.from(content), charset);
                    }
                    //Plain text message
                    result.text = content;
                } else if (
                    !result.text &&
                    contentType &&
                    contentType.indexOf("multipart") >= 0
                ) {
                    if (Array.isArray(content)) {
                        for (let i = 0; i < content.length; i++) {
                            _append(
                                content[i].part.headers,
                                content[i].part.body
                            );
                        }
                    }
                } else {
                    //Get the attachment
                    if (!result.attachments) {
                        result.attachments = [];
                    }

                    const attachment = {};

                    const id = headers["Content-ID"];
                    if (id) {
                        attachment.id = id;
                    }

                    let name =
                        headers["Content-Disposition"] ||
                        headers["Content-Type"];
                    if (name) {
                        const match = /name="?(.+?)"?$/gi.exec(name);
                        if (match) {
                            name = match[1];
                        } else {
                            name = null;
                        }
                    }
                    if (name) {
                        attachment.name = name;
                    }

                    const ct = headers["Content-Type"];
                    if (ct) {
                        attachment.contentType = ct;
                    }

                    const cd = headers["Content-Disposition"];
                    if (cd) {
                        attachment.inline = /^\s*inline/g.test(cd);
                    }

                    attachment.data = content;
                    result.attachments.push(attachment);
                }
            }

            //Content mime type
            let boundary = null;
            const ct = data.headers["Content-Type"];
            if (ct && ct.startsWith("multipart/")) {
                var b = emlformat.getBoundary(ct);
                if (b && b.length) {
                    boundary = b;
                }
            }

            if (boundary) {
                for (let i = 0; i < data.body.length; i++) {
                    var b = data.body[i];

                    //Get the message content
                    if (typeof b.part == "undefined") {
                        console.warn("Warning: undefined b.part");
                    } else if (typeof b.part == "string") {
                        result.data = b.part;
                    } else {
                        if (typeof b.part.body == "undefined") {
                            console.warn("Warning: undefined b.part.body");
                        } else if (typeof b.part.body == "string") {
                            b.part.body;

                            var headers = b.part.headers;
                            var content = b.part.body;

                            _append(headers, content);
                        } else {
                            for (let j = 0; j < b.part.body.length; j++) {
                                if (typeof b.part.body[j] == "string") {
                                    result.data = b.part.body[j];
                                    continue;
                                }

                                var headers = b.part.body[j].part.headers;
                                var content = b.part.body[j].part.body;

                                _append(headers, content);
                            }
                        }
                    }
                }
            } else if (typeof data.body == "string") {
                _append(data.headers, data.body);
            }

            callback(null, result);
        } catch (e) {
            callback(e);
        }
    }

    if (typeof eml == "string") {
        parse(eml, options, function (error, data) {
            if (error) return callback(error);
            if (!data) return callback(new Error("Cannot parse EML content!"));
            _read(data);
        });
    } else if (typeof eml == "object") {
        _read(eml);
    } else {
        callback(new Error("Missing EML file content!"));
    }
};
