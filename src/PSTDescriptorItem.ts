/* eslint-disable @typescript-eslint/no-explicit-any */
import * as long from "long";

import { PSTFile } from "./PSTFile";
import type { PSTNodeInputStream } from "./PSTNodeInputStream";
import { PSTUtil } from "./PSTUtil";

export class PSTDescriptorItem {
  private dataBlockData: Buffer | null = null;

  private readonly dataBlockOffsets: number[] = [];

  private readonly _pstFile: PSTFile;

  private readonly _subNodeOffsetIndexIdentifier: number;

  public get subNodeOffsetIndexIdentifier(): number {
    return this._subNodeOffsetIndexIdentifier;
  }

  private readonly _descriptorIdentifier: number;

  public get descriptorIdentifier(): number {
    return this._descriptorIdentifier;
  }

  private readonly _offsetIndexIdentifier: number;

  public get offsetIndexIdentifier(): number {
    return this._offsetIndexIdentifier;
  }

  /**
   * Creates an instance of PSTDescriptorItem.
   * @param {Buffer} data
   * @param {number} offset
   * @param {PSTFile} pstFile
   * @memberof PSTDescriptorItem
   */
  constructor(data: Buffer, offset: number, pstFile: PSTFile) {
    this._pstFile = pstFile;

    if (pstFile.pstFileType == PSTFile.PST_TYPE_ANSI) {
      this._descriptorIdentifier = PSTUtil.convertLittleEndianBytesToLong(
        data,
        offset,
        offset + 4
      ).toNumber();
      this._offsetIndexIdentifier =
        PSTUtil.convertLittleEndianBytesToLong(
          data,
          offset + 4,
          offset + 8
        ).toNumber() & 0xfffffffe;
      this._subNodeOffsetIndexIdentifier =
        PSTUtil.convertLittleEndianBytesToLong(
          data,
          offset + 8,
          offset + 12
        ).toNumber() & 0xfffffffe;
    } else {
      this._descriptorIdentifier = PSTUtil.convertLittleEndianBytesToLong(
        data,
        offset,
        offset + 4
      ).toNumber();
      this._offsetIndexIdentifier =
        PSTUtil.convertLittleEndianBytesToLong(
          data,
          offset + 8,
          offset + 16
        ).toNumber() & 0xfffffffe;
      this._subNodeOffsetIndexIdentifier =
        PSTUtil.convertLittleEndianBytesToLong(
          data,
          offset + 16,
          offset + 24
        ).toNumber() & 0xfffffffe;
    }
  }

  /**
   * Get a node input stream from the offset index and read into a buffer.
   * @returns {Buffer}
   * @memberof PSTDescriptorItem
   */
  public getData(): Buffer {
    if (this.dataBlockData != null) {
      return this.dataBlockData;
    }

    const pstNodeInputStream: PSTNodeInputStream = this._pstFile.readLeaf(
      long.fromValue(this.offsetIndexIdentifier)
    );
    const out = Buffer.alloc(pstNodeInputStream.length.toNumber());
    pstNodeInputStream.readCompletely(out);
    this.dataBlockData = out;
    return this.dataBlockData;
  }

  /**
   * Get block offsets within current file.
   * @returns {number[]}
   * @memberof PSTDescriptorItem
   */
  public getBlockOffsets(): number[] {
    debugger;

    if (this.dataBlockOffsets != null) {
      return this.dataBlockOffsets;
    }
    const offsets: long[] = this._pstFile
      .readLeaf(long.fromNumber(this.offsetIndexIdentifier))
      .getBlockOffsets();
    const offsetsOut: number[] = [];
    for (let x = 0; x < offsets.length; x++) {
      offsetsOut[x] = offsets[x].toNumber();
    }
    return offsetsOut;
  }

  /**
   * Get the size of this this leaf of the b-tree.
   * @readonly
   * @type {number}
   * @memberof PSTDescriptorItem
   */
  public get dataSize(): number {
    return this._pstFile.getLeafSize(
      long.fromNumber(this.offsetIndexIdentifier)
    );
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   * @memberof PSTDescriptorItem
   */
  public toJSON(): any {
    return this;
  }
}
