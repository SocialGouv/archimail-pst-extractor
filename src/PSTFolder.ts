import long from "long";

import { PSTUtil } from ".";
import type { DescriptorIndexNode } from "./DescriptorIndexNode";
import { OutlookProperties } from "./OutlookProperties";
import type { PSTDescriptorItem } from "./PSTDescriptorItem";
import type { PSTFile } from "./PSTFile";
import type { PSTMessage } from "./PSTMessage";
import { PSTNodeInputStream } from "./PSTNodeInputStream";
import { PSTObject } from "./PSTObject";
import { PSTTable7C } from "./PSTTable7C";
import type { PSTTableBC } from "./PSTTableBC";
import type { PSTTableItem } from "./PSTTableItem";

/**
 * Represents a folder in the PST File.  Allows you to access child folders or items.
 * Items are accessed through a sort of cursor arrangement.  This allows for
 * incremental reading of a folder which may have _lots_ of emails.
 * @export
 * @class PSTFolder
 * @extends {PSTObject}
 */
export class PSTFolder extends PSTObject {
    private currentEmailIndex = 0;

    private subfoldersTable: PSTTable7C | null = null;

    private emailsTable: PSTTable7C | null = null;

    private fallbackEmailsTable: DescriptorIndexNode[] | null = null;

    /**
     * Creates an instance of PSTFolder.
     * Represents a folder in the PST File.  Allows you to access child folders or items.
     * Items are accessed through a sort of cursor arrangement.  This allows for
     * incremental reading of a folder which may have _lots_ of emails.
     * @param {PSTFile} pstFile
     * @param {DescriptorIndexNode} descriptorIndexNode
     * @param {PSTTableBC} [table]
     * @param {Map<number, PSTDescriptorItem>} [localDescriptorItems]
     * @memberof PSTFolder
     */
    constructor(
        pstFile: PSTFile,
        descriptorIndexNode: DescriptorIndexNode,
        table?: PSTTableBC,
        localDescriptorItems?: Map<number, PSTDescriptorItem>
    ) {
        super(pstFile, descriptorIndexNode);
        if (table) {
            // pre-populate folder object with values
            this.prePopulate(descriptorIndexNode, table, localDescriptorItems);
        }
    }

    /**
     * Get folders in one fell swoop, since there's not usually thousands of them.
     */
    public getSubFolders(): PSTFolder[] {
        const output: PSTFolder[] = [];
        try {
            this.initSubfoldersTable();
            if (this.subfoldersTable) {
                const itemMapSet = this.subfoldersTable.getItems();
                for (const itemMap of itemMapSet) {
                    const item = itemMap.get(26610);
                    if (item) {
                        output.push(
                            PSTUtil.detectAndLoadPSTObject(
                                this.pstFile,
                                long.fromNumber(item.entryValueReference)
                            ) as PSTFolder
                        );
                    }
                }
            }
        } catch (err: unknown) {
            console.error(
                `PSTFolder::getSubFolders Can't get child folders for folder ${this.displayName}\n${err}`
            );
            throw err;
        }
        return output;
    }

    /**
     * Get the next child of this folder. As there could be thousands of emails, we have these
     * kind of cursor operations.
     */
    public getNextChild(): PSTMessage | null {
        this.initEmailsTable();

        if (this.emailsTable) {
            if (this.currentEmailIndex === this.contentCount) {
                // no more!
                return null;
            }

            // get the emails from the rows in the main email table
            const rows: Map<number, PSTTableItem>[] = this.emailsTable.getItems(
                this.currentEmailIndex,
                1
            );
            const emailRow = rows[0].get(0x67f2);
            if ((emailRow && emailRow.itemIndex === -1) || !emailRow) {
                // no more!
                return null;
            }

            const childDescriptor = this.pstFile.getDescriptorIndexNode(
                long.fromNumber(emailRow.entryValueReference)
            );
            const child = PSTUtil.detectAndLoadPSTObject(
                this.pstFile,
                childDescriptor
            ) as PSTMessage;
            this.currentEmailIndex++;
            return child;
        } else if (this.fallbackEmailsTable) {
            if (
                this.currentEmailIndex >= this.contentCount ||
                this.currentEmailIndex >= this.fallbackEmailsTable.length
            ) {
                // no more!
                return null;
            }

            const childDescriptor =
                this.fallbackEmailsTable[this.currentEmailIndex];
            const child = PSTUtil.detectAndLoadPSTObject(
                this.pstFile,
                childDescriptor
            ) as PSTMessage;
            this.currentEmailIndex++;
            return child;
        }
        return null;
    }

    /**
     * Iterate over children in this folder.
     */
    public *childrenIterator(): Generator<PSTMessage, void> {
        if (this.contentCount) {
            let child = this.getNextChild();
            while (child) {
                yield child;
                child = this.getNextChild();
            }
        }
    }

    /**
     *  Move the internal folder cursor to the desired position position 0 is before the first record.
     */
    public moveChildCursorTo(newIndex: number): void {
        this.initEmailsTable();

        if (newIndex < 1) {
            this.currentEmailIndex = 0;
            return;
        }
        if (newIndex > this.contentCount) {
            newIndex = this.contentCount;
        }
        this.currentEmailIndex = newIndex;
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                containerClass: this.containerClass,
                containerFlags: this.containerFlags,
                contentCount: this.contentCount,
                emailCount: this.emailCount,
                folderType: this.folderType,
                hasSubfolders: this.hasSubfolders,
                subFolderCount: this.subFolderCount,
                unreadCount: this.unreadCount,
            },
            this
        );
        return clone;
    }

    /**
     * Load subfolders table.
     */
    private initSubfoldersTable(): void {
        if (this.subfoldersTable) {
            return;
        }
        if (!this.descriptorIndexNode) {
            throw new Error(
                "PSTFolder::initSubfoldersTable descriptorIndexNode is null"
            );
        }

        const folderDescriptorIndex: long = long.fromValue(
            this.descriptorIndexNode.descriptorIdentifier + 11
        );
        try {
            const folderDescriptor: DescriptorIndexNode =
                this.pstFile.getDescriptorIndexNode(folderDescriptorIndex);
            let tmp = undefined;
            if (
                folderDescriptor.localDescriptorsOffsetIndexIdentifier.greaterThan(
                    0
                )
            ) {
                tmp = this.pstFile.getPSTDescriptorItems(
                    folderDescriptor.localDescriptorsOffsetIndexIdentifier
                );
            }
            const offsetIndexItem = this.pstFile.getOffsetIndexNode(
                folderDescriptor.dataOffsetIndexIdentifier
            );
            const pstNodeInputStream = new PSTNodeInputStream(
                this.pstFile,
                offsetIndexItem
            );
            this.subfoldersTable = new PSTTable7C(pstNodeInputStream, tmp);
        } catch (err: unknown) {
            console.error(
                `PSTFolder::initSubfoldersTable Can't get child folders for folder ${this.displayName}\n${err}`
            );
            throw err;
        }
    }

    // get all of the children
    private initEmailsTable(): void {
        if (this.emailsTable || this.fallbackEmailsTable) {
            return;
        }

        // some folder types don't have children:
        if (this.getNodeType() === PSTUtil.NID_TYPE_SEARCH_FOLDER) {
            return;
        }
        if (!this.descriptorIndexNode) {
            throw new Error(
                "PSTFolder::initEmailsTable descriptorIndexNode is null"
            );
        }

        try {
            const folderDescriptorIndex =
                this.descriptorIndexNode.descriptorIdentifier + 12;
            const folderDescriptor: DescriptorIndexNode =
                this.pstFile.getDescriptorIndexNode(
                    long.fromNumber(folderDescriptorIndex)
                );
            let tmp = undefined;
            if (
                folderDescriptor.localDescriptorsOffsetIndexIdentifier.greaterThan(
                    0
                )
            ) {
                tmp = this.pstFile.getPSTDescriptorItems(
                    folderDescriptor.localDescriptorsOffsetIndexIdentifier
                );
            }
            const offsetIndexItem = this.pstFile.getOffsetIndexNode(
                folderDescriptor.dataOffsetIndexIdentifier
            );
            const pstNodeInputStream = new PSTNodeInputStream(
                this.pstFile,
                offsetIndexItem
            );
            this.emailsTable = new PSTTable7C(pstNodeInputStream, tmp, 0x67f2);
        } catch (err: unknown) {
            // fallback to children as listed in the descriptor b-tree
            // console.log(`PSTFolder::initEmailsTable Can't get child folders for folder {this.displayName}, resorting to using alternate tree`);
            const tree = this.pstFile.getChildDescriptorTree();
            this.fallbackEmailsTable = [];
            const allChildren = tree.get(
                this.descriptorIndexNode.descriptorIdentifier
            );
            if (allChildren) {
                // remove items that aren't messages
                for (const node of allChildren) {
                    if (
                        this.getNodeType(node.descriptorIdentifier) ==
                        PSTUtil.NID_TYPE_NORMAL_MESSAGE
                    ) {
                        this.fallbackEmailsTable.push(node);
                    }
                }
            }
        }
    }

    /**
     * The number of child folders in this folder
     */
    public get subFolderCount(): number {
        this.initSubfoldersTable();
        if (this.subfoldersTable !== null) {
            return this.subfoldersTable.rowCount;
        } else {
            return 0;
        }
    }

    /**
     * Number of emails in this folder
     */
    public get emailCount(): number {
        this.initEmailsTable();
        if (this.emailsTable === null) {
            return -1;
        }
        return this.emailsTable.rowCount;
    }

    /**
     * Contains a constant that indicates the folder type.
     * https://msdn.microsoft.com/en-us/library/office/cc815373.aspx
     */
    public get folderType(): number {
        return this.getIntItem(OutlookProperties.PR_FOLDER_TYPE);
    }

    /**
     * Contains the number of messages in a folder, as computed by the message store.
     * For a number calculated by the library use getEmailCount
     */
    public get contentCount(): number {
        return this.getIntItem(OutlookProperties.PR_CONTENT_COUNT);
    }

    /**
     * Contains the number of unread messages in a folder, as computed by the message store.
     * https://msdn.microsoft.com/en-us/library/office/cc841964.aspx
     */
    public get unreadCount(): number {
        return this.getIntItem(OutlookProperties.PR_CONTENT_UNREAD);
    }

    /**
     * Contains TRUE if a folder contains subfolders.
     * once again, read from the PST, use getSubFolderCount if you want to know
     * @readonly
     * @type {boolean}
     * @memberof PSTFolder
     */
    public get hasSubfolders(): boolean {
        return this.getIntItem(OutlookProperties.PR_SUBFOLDERS) !== 0;
    }

    /**
     * Contains a text string describing the type of a folder. Although this property is
     * generally ignored, versions of Microsoft® Exchange Server prior to Exchange Server
     * 2003 Mailbox Manager expect this property to be present.
     * https://msdn.microsoft.com/en-us/library/office/cc839839.aspx
     */
    public get containerClass(): string {
        return this.getStringItem(OutlookProperties.PR_CONTAINER_CLASS);
    }

    /**
     * Contains a bitmask of flags describing capabilities of an address book container.
     * https://msdn.microsoft.com/en-us/library/office/cc839610.aspx
     */
    public get containerFlags(): number {
        return this.getIntItem(OutlookProperties.PR_CONTAINER_FLAGS);
    }
}
