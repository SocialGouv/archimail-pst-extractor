/* eslint-disable */
// @ts-ignore
// @ts-nocheck
// TODO

/******************************************************************************************
 * Unpacks EML message and attachments to a directory.
 * @params eml         EML file content or object from 'parse'
 * @params directory   Folder name or directory path where to unpack
 * @params options     Optional parameters: { parsedJsonFile, readJsonFile, simulate }
 * @params callback    Callback function(error)
 ******************************************************************************************/
export const unpack = (eml, directory, options, callback) => {
    //Shift arguments
    if (typeof options == "function" && typeof callback == "undefined") {
        callback = options;
        options = null;
    }

    if (typeof callback != "function") {
        callback = function (error, result) {};
    }

    const result = { files: [] };

    function _unpack(data) {
        try {
            //Create the target directory
            if (!fs.existsSync(directory)) {
                fs.mkdirSync(directory);
            }

            //Plain text file
            if (typeof data.text == "string") {
                result.files.push("index.txt");
                if (options && options.simulate) {
                    //Skip writing to file
                } else {
                    fs.writeFileSync(
                        path.join(directory, "index.txt"),
                        data.text
                    );
                }
            }

            //Message in HTML format
            if (typeof data.html == "string") {
                result.files.push("index.html");
                if (options && options.simulate) {
                    //Skip writing to file
                } else {
                    fs.writeFileSync(
                        path.join(directory, "index.html"),
                        data.html
                    );
                }
            }

            //Attachments
            if (data.attachments && data.attachments.length > 0) {
                for (let i = 0; i < data.attachments.length; i++) {
                    const attachment = data.attachments[i];
                    let filename = attachment.name;
                    if (!filename) {
                        filename = `attachment_${
                            i + 1
                        }${emlformat.getFileExtension(attachment.mimeType)}`;
                    }
                    result.files.push(filename);
                    if (options && options.simulate) continue; //Skip writing to file
                    fs.writeFileSync(
                        path.join(directory, filename),
                        attachment.data
                    );
                }
            }

            callback(null, result);
        } catch (e) {
            callback(e);
        }
    }

    //Check the directory argument
    if (typeof directory != "string" || directory.length == 0) {
        return callback(new Error("Directory argument is missing!"));
    }

    //Argument as EML file content or "parsed" version of object
    if (
        typeof eml == "string" ||
        (typeof eml == "object" && eml.headers && eml.body)
    ) {
        emlformat.parse(eml, function (error, parsed) {
            if (error) return callback(error);

            //Save parsed EML as JSON file
            if (options && options.parsedJsonFile) {
                const file = path.resolve(directory, options.parsedJsonFile);
                const dir = path.dirname(file);
                if (!fs.existsSync(dir)) {
                    fs.mkdirSync(dir);
                }
                result.files.push(options.parsedJsonFile);
                fs.writeFileSync(file, JSON.stringify(parsed, " ", 2));
            }

            //Convert parsed EML object to a friendly object with text, html and attachments
            emlformat.read(parsed, function (error, data) {
                if (error) return callback(error);

                //Save read structure as JSON file
                if (options && options.readJsonFile) {
                    const file = path.resolve(directory, options.readJsonFile);
                    const dir = path.dirname(file);
                    if (!fs.existsSync(dir)) {
                        fs.mkdirSync(dir);
                    }
                    result.files.push(options.readJsonFile);
                    const json = data.attachments
                        ? JSON.stringify(data)
                        : JSON.stringify(data, " ", 2); //Attachments may be large, so make a compact JSON string
                    fs.writeFileSync(file, json);
                }

                //Extract files from the EML file
                _unpack(data);
            });
        });
    } else if (typeof eml != "object") {
        return callback(new Error("Expected string or object as argument!"));
    } else {
        _unpack(eml);
    }
};
