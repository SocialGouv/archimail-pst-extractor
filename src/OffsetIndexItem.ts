import type long from "long";

import { PSTFile, PSTUtil } from ".";

export class OffsetIndexItem {
    private readonly _indexIdentifier: long;

    public get indexIdentifier(): long {
        return this._indexIdentifier;
    }

    private readonly _fileOffset: long;

    public get fileOffset(): long {
        return this._fileOffset;
    }

    private readonly _size: number;

    public get size(): number {
        return this._size;
    }

    private readonly cRef: long;

    /**
     * Creates an instance of OffsetIndexItem, part of the node table.
     */
    constructor(data: Buffer, pstFileType: number) {
        if (pstFileType === PSTFile.PST_TYPE_ANSI) {
            this._indexIdentifier = PSTUtil.convertLittleEndianBytesToLong(
                data,
                0,
                4
            );
            this._fileOffset = PSTUtil.convertLittleEndianBytesToLong(
                data,
                4,
                8
            );
            this._size = PSTUtil.convertLittleEndianBytesToLong(
                data,
                8,
                10
            ).toNumber();
            this.cRef = PSTUtil.convertLittleEndianBytesToLong(data, 10, 12);
        } else {
            this._indexIdentifier = PSTUtil.convertLittleEndianBytesToLong(
                data,
                0,
                8
            );
            this._fileOffset = PSTUtil.convertLittleEndianBytesToLong(
                data,
                8,
                16
            );
            this._size = PSTUtil.convertLittleEndianBytesToLong(
                data,
                16,
                18
            ).toNumber();
            this.cRef = PSTUtil.convertLittleEndianBytesToLong(data, 16, 18);
        }
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        return this;
    }
}
