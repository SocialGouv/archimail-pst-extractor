/* eslint-disable no-fallthrough */
import long from "long";

import { PSTUtil } from ".";
import type { PSTDescriptorItem } from "./PSTDescriptorItem";
import type { PSTNodeInputStream } from "./PSTNodeInputStream";
import { PSTTable } from "./PSTTable";
import { PSTTableItem } from "./PSTTableItem";

export class PSTTableBC extends PSTTable {
    private readonly items: Map<number, PSTTableItem> = new Map();

    // private readonly isDescNotYetInitiated = false;

    /**
     * Creates an instance of PSTTableBC ("Property Context").
     */
    constructor(pstNodeInputStream: PSTNodeInputStream) {
        super(pstNodeInputStream, new Map<number, PSTDescriptorItem>());

        if (this.tableTypeByte !== 188) {
            throw new Error(
                "PSTTableBC::constructor unable to create PSTTableBC, table does not appear to be a bc!"
            );
        }

        // go through each of the entries
        const keyTableInfoNodeInfo = this.getNodeInfo(this.hidRoot);
        if (!keyTableInfoNodeInfo) {
            throw new Error(
                "PSTTableBC::constructor keyTableInfoNodeInfo is null"
            );
        }

        const keyTableInfo: Buffer = Buffer.alloc(
            keyTableInfoNodeInfo.length()
        );
        keyTableInfoNodeInfo.pstNodeInputStream.seek(
            long.fromValue(keyTableInfoNodeInfo.startOffset)
        );
        keyTableInfoNodeInfo.pstNodeInputStream.readCompletely(keyTableInfo);
        this.numberOfKeys = Math.trunc(
            keyTableInfo.length / (this.sizeOfItemKey + this.sizeOfItemValue)
        );
        if (this.numberOfKeys === 0) {
            // debugger
        }

        // Read the key table
        let offset = 0;
        for (let x = 0; x < this.numberOfKeys; x++) {
            const item = new PSTTableItem();
            item.itemIndex = x;
            item.entryType = PSTUtil.convertLittleEndianBytesToLong(
                keyTableInfo,
                offset + 0,
                offset + 2
            );
            item.entryValueType = PSTUtil.convertLittleEndianBytesToLong(
                keyTableInfo,
                offset + 2,
                offset + 4
            ).toNumber();
            item.entryValueReference = PSTUtil.convertLittleEndianBytesToLong(
                keyTableInfo,
                offset + 4,
                offset + 8
            ).toNumber();

            // Data is in entryValueReference for all types <= 4 bytes long
            switch (item.entryValueType) {
                case 0x0002: // 16bit integer
                    item.entryValueReference &= 0xffff;
                case 0x0003: // 32bit integer
                case 0x000a: // 32bit error code
                case 0x0001: // Place-holder
                case 0x0004: // 32bit floating
                    item.isExternalValueReference = true;
                    break;

                case 0x000b: // Boolean - a single byte
                    item.entryValueReference &= 0xff;
                    item.isExternalValueReference = true;
                    break;

                case 0x000d:
                default:
                    // Is it in the local heap?
                    // Assume not
                    item.isExternalValueReference = true;
                    // eslint-disable-next-line no-case-declarations
                    const nodeInfoNodeInfo = this.getNodeInfo(
                        item.entryValueReference
                    );
                    if (nodeInfoNodeInfo === null) {
                        // It's an external reference that we don't deal with here.
                    } else {
                        // Make a copy of the data
                        const nodeInfo = Buffer.alloc(
                            nodeInfoNodeInfo.length()
                        );
                        nodeInfoNodeInfo.pstNodeInputStream.seek(
                            long.fromValue(nodeInfoNodeInfo.startOffset)
                        );
                        nodeInfoNodeInfo.pstNodeInputStream.readCompletely(
                            nodeInfo
                        );
                        item.data = nodeInfo; // should be new array, so just use it
                        item.isExternalValueReference = false;
                    }
                    break;
            }
            offset = offset + 8;
            this.items.set(item.entryType.toNumber(), item);
        }
        this.releaseRawData();
    }

    /**
     * Get the items parsed out of this table.
     */
    public getItems(): Map<number, PSTTableItem> {
        return this.items;
    }

    /**
     * JSON stringify the items list.
     */
    public itemsJSON(): string {
        let s = "";
        this.items.forEach((item) => {
            s = s + JSON.stringify(item);
        });
        return s;
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        return this;
    }
}
