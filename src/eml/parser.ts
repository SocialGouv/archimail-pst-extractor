/* eslint-disable */
// @ts-ignore
// @ts-nocheck
// TODO

/******************************************************************************************
 * Parses EML file content and returns object-oriented representation of the content.
 * @params eml         EML file content
 * @params options     EML parse options
 * @params callback    Callback function(error, data)
 ******************************************************************************************/
export const parse = (eml, options, callback) => {
    //Shift arguments
    if (typeof options == "function" && typeof callback == "undefined") {
        callback = options;
        options = null;
    }

    if (typeof callback != "function") {
        callback = function (error, result) {};
    }

    try {
        if (typeof eml != "string") {
            throw new Error("Argument 'eml' expected to be string!");
        }

        const lines = eml.split(/\r?\n/);
        const result = {};
        parseRecursive(lines, 0, result, options);
        callback(null, result);
    } catch (e) {
        callback(e);
    }
};

/******************************************************************************************
 * Parses EML file content.
 ******************************************************************************************/
function parseRecursive(lines, start, parent, options) {
    let boundary = null;
    let lastHeaderName = "";
    let findBoundary = "";
    let insideBody = false;
    let insideBoundary = false;
    let isMultiHeader = false;
    let isMultipart = false;

    parent.headers = {};
    //parent.body = null;

    function complete(boundary) {
        //boundary.part = boundary.lines.join("\r\n");
        boundary.part = {};
        parseRecursive(boundary.lines, 0, boundary.part, options);
        delete boundary.lines;
    }

    //Read line by line
    for (let i = start; i < lines.length; i++) {
        const line = lines[i];

        //Header
        if (!insideBody) {
            //Search for empty line
            if (line == "") {
                insideBody = true;

                if (options && options.headersOnly) {
                    break;
                }

                //Expected boundary
                const ct = parent.headers["Content-Type"];
                if (ct && ct.startsWith("multipart/")) {
                    const b = emlformat.getBoundary(ct);
                    if (b && b.length) {
                        findBoundary = b;
                        isMultipart = true;
                        parent.body = [];
                    } else {
                        if (emlformat.verbose) {
                            console.warn(
                                `Multipart without boundary! ${ct.replace(
                                    /\r?\n/g,
                                    " "
                                )}`
                            );
                        }
                    }
                }

                continue;
            }

            //Header value with new line
            var match = /^\s+([^\r\n]+)/g.exec(line);
            if (match) {
                if (isMultiHeader) {
                    parent.headers[lastHeaderName][
                        parent.headers[lastHeaderName].length - 1
                    ] += `\r\n${match[1]}`;
                } else {
                    parent.headers[lastHeaderName] += `\r\n${match[1]}`;
                }
                continue;
            }

            //Header name and value
            var match = /^([\w\d\-]+):\s+([^\r\n]+)/gi.exec(line);
            if (match) {
                lastHeaderName = match[1];
                if (parent.headers[lastHeaderName]) {
                    //Multiple headers with the same name
                    isMultiHeader = true;
                    if (typeof parent.headers[lastHeaderName] == "string") {
                        parent.headers[lastHeaderName] = [
                            parent.headers[lastHeaderName],
                        ];
                    }
                    parent.headers[lastHeaderName].push(match[2]);
                } else {
                    //Header first appeared here
                    isMultiHeader = false;
                    parent.headers[lastHeaderName] = match[2];
                }
                continue;
            }
        }
        //Body
        else {
            //Multipart body
            if (isMultipart) {
                //Search for boundary start

                //Updated on 2019-10-12: A line before the boundary marker is not required to be an empty line
                //if (lines[i - 1] == "" && line.indexOf("--" + findBoundary) == 0 && !/\-\-(\r?\n)?$/g.test(line)) {
                if (
                    line.indexOf(`--${findBoundary}`) == 0 &&
                    !/\-\-(\r?\n)?$/g.test(line)
                ) {
                    insideBoundary = true;

                    //Complete the previous boundary
                    if (boundary && boundary.lines) {
                        complete(boundary);
                    }

                    //Start a new boundary
                    var match = /^\-\-([^\r\n]+)(\r?\n)?$/g.exec(line);
                    boundary = { boundary: match[1], lines: [] };
                    parent.body.push(boundary);

                    if (emlformat.verbose) {
                        console.log(`Found boundary: ${boundary.boundary}`);
                    }

                    continue;
                }

                if (insideBoundary) {
                    //Search for boundary end
                    if (
                        boundary.boundary &&
                        lines[i - 1] == "" &&
                        line.indexOf(`--${findBoundary}--`) == 0
                    ) {
                        insideBoundary = false;
                        complete(boundary);
                        continue;
                    }
                    boundary.lines.push(line);
                }
            } else {
                //Solid string body
                parent.body = lines.splice(i).join("\r\n");
                break;
            }
        }
    }

    //Complete the last boundary
    if (parent.body?.[parent.body.length - 1].lines) {
        complete(parent.body[parent.body.length - 1]);
    }
}
