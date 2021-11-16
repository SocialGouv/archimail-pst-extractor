import type long from "long";

import { PSTUtil } from ".";
import { PSTFile } from "./PSTFile";

// DescriptorIndexNode is a leaf item from the Descriptor index b-tree
// It is like a pointer to an element in the PST file, everything has one...
export class DescriptorIndexNode {
    public itemType = 0;

    private readonly _descriptorIdentifier: number;

    public get descriptorIdentifier(): number {
        return this._descriptorIdentifier;
    }

    private readonly _parentDescriptorIndexIdentifier: number;

    public get parentDescriptorIndexIdentifier(): number {
        return this._parentDescriptorIndexIdentifier;
    }

    private readonly _localDescriptorsOffsetIndexIdentifier: long;

    public get localDescriptorsOffsetIndexIdentifier(): long {
        return this._localDescriptorsOffsetIndexIdentifier;
    }

    private readonly _dataOffsetIndexIdentifier: long;

    public get dataOffsetIndexIdentifier(): long {
        return this._dataOffsetIndexIdentifier;
    }

    /**
     * Creates an instance of DescriptorIndexNode, a component of the internal descriptor b-tree.
     */
    constructor(buffer: Buffer, pstFileType: number) {
        if (pstFileType === PSTFile.PST_TYPE_ANSI) {
            this._descriptorIdentifier = PSTUtil.convertLittleEndianBytesToLong(
                buffer,
                0,
                4
            ).toNumber();
            this._dataOffsetIndexIdentifier =
                PSTUtil.convertLittleEndianBytesToLong(buffer, 4, 8);
            this._localDescriptorsOffsetIndexIdentifier =
                PSTUtil.convertLittleEndianBytesToLong(buffer, 8, 12);
            this._parentDescriptorIndexIdentifier =
                PSTUtil.convertLittleEndianBytesToLong(
                    buffer,
                    12,
                    16
                ).toNumber();
        } else {
            this._descriptorIdentifier = PSTUtil.convertLittleEndianBytesToLong(
                buffer,
                0,
                4
            ).toNumber();
            this._dataOffsetIndexIdentifier =
                PSTUtil.convertLittleEndianBytesToLong(buffer, 8, 16);
            this._localDescriptorsOffsetIndexIdentifier =
                PSTUtil.convertLittleEndianBytesToLong(buffer, 16, 24);
            this._parentDescriptorIndexIdentifier =
                PSTUtil.convertLittleEndianBytesToLong(
                    buffer,
                    24,
                    28
                ).toNumber();
            this.itemType = PSTUtil.convertLittleEndianBytesToLong(
                buffer,
                28,
                32
            ).toNumber();
        }
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        return this;
    }
}
