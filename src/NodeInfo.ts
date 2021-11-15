import type long from "long";

import type { PSTNodeInputStream } from "./PSTNodeInputStream";

export class NodeInfo {
    private readonly _startOffset: number;

    get startOffset(): number {
        return this._startOffset;
    }

    private readonly _pstNodeInputStream: PSTNodeInputStream;

    private readonly _endOffset: number;

    get endOffset(): number {
        return this._endOffset;
    }

    get pstNodeInputStream(): PSTNodeInputStream {
        return this._pstNodeInputStream;
    }

    /**
     * Creates an instance of NodeInfo, part of the node table.
     */
    constructor(
        start: number,
        end: number,
        pstNodeInputStream: PSTNodeInputStream
    ) {
        if (start > end) {
            throw new Error(
                `NodeInfo:: constructor Invalid NodeInfo parameters: start ${start} is greater than end ${end}`
            );
        }
        this._startOffset = start;
        this._endOffset = end;
        this._pstNodeInputStream = pstNodeInputStream;
    }

    public length(): number {
        return this.endOffset - this.startOffset;
    }

    /**
     * Seek to position in node input stream and read a long
     */
    public seekAndReadLong(offset: long, length: number): long {
        return this.pstNodeInputStream.seekAndReadLong(
            offset.add(this.startOffset),
            length
        );
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        return this;
    }
}
