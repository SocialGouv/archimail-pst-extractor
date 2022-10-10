import fs from "fs";
import path from "path";

import type { PSTFolder } from "../src";
import { emlStringify, PSTFile } from "../src";

jest.mock("crypto", () => ({
    randomUUID() {
        return "78b1fa86-9134-4b2a-a2b5-499db50b9a99";
    },
}));

it("should build .eml file", () => {
    //E-mail data
    const data = {
        attachments: [
            {
                contentType: "text/plain; charset=utf-8",
                data: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Morbi eget elit turpis. Aliquam lorem nunc, dignissim in risus at, tempus aliquet justo. In in libero pharetra, tristique est sed, semper diam. Phasellus faucibus eleifend neque. Etiam vitae dolor non turpis finibus condimentum id vitae dolor. Pellentesque vulputate nisi erat, porttitor iaculis ligula euismod nec. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Sed laoreet, turpis at blandit consequat, mauris enim volutpat augue, vel congue nisi quam quis nibh. Pellentesque ultrices tellus eget ullamcorper accumsan. Suspendisse mattis sit amet enim eu congue.",
                name: "sample.txt",
            },
            {
                contentType: "image/jpeg",
                data: fs.readFileSync(
                    path.resolve(__dirname, "./testdata/rickroll.jpg")
                ),
                inline: true,
                name: "rickroll.jpg",
            },
        ],
        from: "no-reply@bar.com",
        html: '<html><head></head><body>Lorem ipsum...<br /><img src="rickroll.jpg" alt="" /></body></html>',
        subject: "Winter promotions",
        text: "Lorem ipsum...",
        to: {
            email: "foo@bar.com",
            name: "Foo Bar",
        },
    };

    const eml = emlStringify(data);
    expect(typeof eml).toEqual("string");
    expect(eml.length).toBeGreaterThan(0);
    expect(eml).toMatchSnapshot();
});

it("should build .eml file from a pst message", () => {
    const pstFile = new PSTFile(
        path.resolve(__dirname, "./testdata/enron.pst")
    );

    let childFolders: PSTFolder[] = pstFile.getRootFolder().getSubFolders();
    expect(childFolders.length).toEqual(3);
    let folder = childFolders[0];
    expect(folder.subFolderCount).toEqual(2);
    expect(folder.displayName).toEqual("Top of Personal Folders");
    childFolders = folder.getSubFolders();
    folder = childFolders[0];
    expect(folder.displayName).toEqual("Deleted Items");
    folder = childFolders[1];
    expect(folder.displayName).toEqual("lokay-m");
    childFolders = folder.getSubFolders();
    folder = childFolders[0];
    expect(folder.displayName).toEqual("MLOKAY (Non-Privileged)");
    childFolders = folder.getSubFolders();
    expect(childFolders[0].displayName).toEqual("TW-Commercial Group");
    const comGroupFolder = childFolders[0];

    const msg = comGroupFolder.getNextChild()!;

    expect(msg.toEML()).toMatchSnapshot();
});
