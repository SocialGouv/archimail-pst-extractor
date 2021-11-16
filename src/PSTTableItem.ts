import long from "long";

import { PSTUtil } from ".";

/**
 * Generic table item.
 * Provides some basic string functions
 */
export class PSTTableItem {
    public static VALUE_TYPE_PT_UNICODE = 0x1f;

    public static VALUE_TYPE_PT_STRING8 = 0x1e;

    public static VALUE_TYPE_PT_BIN = 0x102;

    private _itemIndex = 0;

    public set itemIndex(val: number) {
        this._itemIndex = val;
    }

    public get itemIndex(): number {
        return this._itemIndex;
    }

    private _entryType: long = long.ZERO;

    public set entryType(val: long) {
        this._entryType = val;
    }

    public get entryType(): long {
        return this._entryType;
    }

    private _isExternalValueReference = false;

    public set isExternalValueReference(val: boolean) {
        this._isExternalValueReference = val;
    }

    public get isExternalValueReference(): boolean {
        return this._isExternalValueReference;
    }

    private _entryValueReference = 0;

    public set entryValueReference(val: number) {
        this._entryValueReference = val;
    }

    public get entryValueReference(): number {
        return this._entryValueReference;
    }

    private _entryValueType = 0;

    public set entryValueType(val: number) {
        this._entryValueType = val;
    }

    public get entryValueType(): number {
        return this._entryValueType;
    }

    private _data: Buffer = Buffer.alloc(0);

    public set data(val: Buffer) {
        this._data = val;
    }

    public get data(): Buffer {
        return this._data;
    }

    /**
     * Get long value from table item.
     */
    public getLongValue(): long {
        if (this.data.length > 0) {
            return PSTUtil.convertLittleEndianBytesToLong(this.data);
        }
        return long.fromNumber(-1);
    }

    /**
     * Get string value of data.
     */
    public getStringValue(stringType = this.entryValueType): string {
        if (stringType === PSTTableItem.VALUE_TYPE_PT_UNICODE) {
            // little-endian unicode string
            try {
                if (this.isExternalValueReference) {
                    return "External string reference!";
                }
                return this.data.toString("utf16le").replace(/\0/g, "");
            } catch (err: unknown) {
                console.error(
                    `Error decoding string: ${this.data
                        .toString("utf16le")
                        .replace(/\0/g, "")}\n${err}`
                );
                return "";
            }
        }

        if (stringType === PSTTableItem.VALUE_TYPE_PT_STRING8) {
            return this.data.toString();
        }

        return "hex string";
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                data: this.data,
                entryType: this.entryType,
                entryValueReference: this.entryValueReference,
                entryValueType: this.entryValueType,
                isExternalValueReference: this.isExternalValueReference,
                itemIndex: this.itemIndex,
            },
            this
        );
        return clone;
    }
}
