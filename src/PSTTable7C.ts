import long from "long";

import { ColumnDescriptor } from "./ColumnDescriptor";
import type { NodeInfo } from "./NodeInfo";
import type { PSTDescriptorItem } from "./PSTDescriptorItem";
import { PSTFile } from "./PSTFile";
import type { PSTNodeInputStream } from "./PSTNodeInputStream";
import { PSTTable } from "./PSTTable";
import { PSTTableItem } from "./PSTTableItem";

export class PSTTable7C extends PSTTable {
    private items: Map<number, PSTTableItem>[] = [];

    private numberOfDataSets = 0;

    private readonly BLOCK_SIZE = 8176;

    private readonly numColumns: number = 0;

    // eslint-disable-next-line @typescript-eslint/naming-convention
    private readonly TCI_bm: number = 0;

    // eslint-disable-next-line @typescript-eslint/naming-convention
    private readonly TCI_1b: number = 0;

    private readonly columnDescriptors: ColumnDescriptor[] = [];

    private readonly overrideCol: number = -1;

    private readonly rowNodeInfo: NodeInfo;

    private readonly keyMap: Map<number, number>;

    /**
     * Creates an instance of PSTTable7C ("Table Context").
     */
    constructor(
        pstNodeInputStream: PSTNodeInputStream,
        subNodeDescriptorItems?: Map<number, PSTDescriptorItem>,
        entityToExtract?: number
    ) {
        super(pstNodeInputStream, subNodeDescriptorItems);

        if (this.tableTypeByte !== 0x7c) {
            throw new Error(
                "PSTTable7C::constructor unable to create PSTTable7C, table does not appear to be a 7c!"
            );
        }

        // TCINFO header is in the hidUserRoot node
        const tcHeaderNode = this.getNodeInfo(this.hidUserRoot);
        if (!tcHeaderNode) {
            throw new Error("PSTTable7C::constructor tcHeaderNode is null");
        }

        // get the TCINFO header information
        let offset = 0;
        this.numColumns = tcHeaderNode
            .seekAndReadLong(long.fromNumber(offset + 1), 1)
            .toNumber();
        // const TCI_4b: number = tcHeaderNode
        //     .seekAndReadLong(long.fromNumber(offset + 2), 2)
        //     .toNumber();
        // const TCI_2b: number = tcHeaderNode
        //     .seekAndReadLong(long.fromNumber(offset + 4), 2)
        //     .toNumber();
        this.TCI_1b = tcHeaderNode
            .seekAndReadLong(long.fromNumber(offset + 6), 2)
            .toNumber();
        this.TCI_bm = tcHeaderNode
            .seekAndReadLong(long.fromNumber(offset + 8), 2)
            .toNumber();
        // const hidRowIndex: number = tcHeaderNode
        //     .seekAndReadLong(long.fromNumber(offset + 10), 4)
        //     .toNumber();
        const hnidRows: number = tcHeaderNode
            .seekAndReadLong(long.fromNumber(offset + 14), 4)
            .toNumber();

        // 22... column descriptors
        offset += 22;
        if (this.numColumns !== 0) {
            for (let col = 0; col < this.numColumns; ++col) {
                this.columnDescriptors[col] = new ColumnDescriptor(
                    tcHeaderNode,
                    offset
                );
                if (this.columnDescriptors[col].id === entityToExtract) {
                    this.overrideCol = col;
                }
                offset += 8;
            }
        }

        // if we are asking for a specific column, only get that!
        if (this.overrideCol > -1) {
            this.numColumns = this.overrideCol + 1;
        }

        // Read the key table
        this.keyMap = new Map();
        const keyTableInfo = this.getNodeInfo(this.hidRoot);
        if (!keyTableInfo) {
            throw new Error("PSTTable7C::constructor keyTableInfo is null");
        }
        this.numberOfKeys = Math.trunc(
            keyTableInfo.length() / (this.sizeOfItemKey + this.sizeOfItemValue)
        );
        offset = 0;
        for (let x = 0; x < this.numberOfKeys; x++) {
            const context = keyTableInfo
                .seekAndReadLong(long.fromNumber(offset), this.sizeOfItemKey)
                .toNumber();
            offset += this.sizeOfItemKey;
            const rowIndex = keyTableInfo
                .seekAndReadLong(long.fromNumber(offset), this.sizeOfItemValue)
                .toNumber();
            offset += this.sizeOfItemValue;
            this.keyMap.set(context, rowIndex);
        }

        // Read the Row Matrix
        const rowNodeInfo = this.getNodeInfo(hnidRows);
        if (!rowNodeInfo) {
            throw new Error("PSTTable7C::constructor rowNodeInfo is null");
        }
        this.rowNodeInfo = rowNodeInfo;
        const numberOfBlocks: number = Math.trunc(
            this.rowNodeInfo.length() / this.BLOCK_SIZE
        );
        const numberOfRowsPerBlock: number = Math.trunc(
            this.BLOCK_SIZE / this.TCI_bm
        );
        // const blockPadding =
        //     this.BLOCK_SIZE - numberOfRowsPerBlock * this.TCI_bm;
        this.numberOfDataSets = Math.trunc(
            numberOfBlocks * numberOfRowsPerBlock +
                (this.rowNodeInfo.length() % this.BLOCK_SIZE) / this.TCI_bm
        );
    }

    /**
     * Get items from the table.
     */
    public getItems(
        startAtRecord?: number,
        numberOfRecordsToReturn?: number
    ): Map<number, PSTTableItem>[] {
        const itemList: Map<number, PSTTableItem>[] = [];
        let setLocalList = false;

        // okay, work out the number of records we have
        const numberOfBlocks = Math.trunc(
            this.rowNodeInfo.length() / this.BLOCK_SIZE
        );
        const numberOfRowsPerBlock = Math.trunc(this.BLOCK_SIZE / this.TCI_bm);
        const blockPadding =
            this.BLOCK_SIZE - numberOfRowsPerBlock * this.TCI_bm;
        this.numberOfDataSets = Math.trunc(
            numberOfBlocks * numberOfRowsPerBlock +
                (this.rowNodeInfo.length() % this.BLOCK_SIZE) / this.TCI_bm
        );

        if (startAtRecord === undefined) {
            numberOfRecordsToReturn = this.numberOfDataSets;
            startAtRecord = 0;
            setLocalList = true;
        }

        if (numberOfRecordsToReturn === undefined) {
            numberOfRecordsToReturn = 0;
        }

        // repeat the reading process for every dataset
        let currentValueArrayStart =
            Math.trunc(startAtRecord / numberOfRowsPerBlock) * this.BLOCK_SIZE +
            (startAtRecord % numberOfRowsPerBlock) * this.TCI_bm;
        if (numberOfRecordsToReturn > this.rowCount - startAtRecord) {
            numberOfRecordsToReturn = this.rowCount - startAtRecord;
        }

        let dataSetNumber = 0;
        for (
            let rowCounter = 0;
            rowCounter < numberOfRecordsToReturn;
            rowCounter++
        ) {
            const currentItem: Map<number, PSTTableItem> = new Map();
            // add on some padding for block boundries?
            if (
                this.rowNodeInfo.pstNodeInputStream.pstFile.pstFileType ==
                PSTFile.PST_TYPE_ANSI
            ) {
                if (currentValueArrayStart >= this.BLOCK_SIZE) {
                    currentValueArrayStart =
                        currentValueArrayStart +
                        4 * (currentValueArrayStart / this.BLOCK_SIZE);
                }
                if (
                    this.rowNodeInfo.startOffset +
                        currentValueArrayStart +
                        this.TCI_1b >
                    this.rowNodeInfo.pstNodeInputStream.length.toNumber()
                ) {
                    continue;
                }
            } else {
                if (
                    currentValueArrayStart % this.BLOCK_SIZE >
                    this.BLOCK_SIZE - this.TCI_bm
                ) {
                    // adjust!
                    currentValueArrayStart += blockPadding;
                    if (
                        currentValueArrayStart + this.TCI_bm >
                        this.rowNodeInfo.length()
                    ) {
                        continue;
                    }
                }
            }
            const bitmap = Buffer.alloc((this.numColumns + 7) / 8);
            this.rowNodeInfo.pstNodeInputStream.seek(
                long.fromNumber(
                    this.rowNodeInfo.startOffset +
                        currentValueArrayStart +
                        this.TCI_1b
                )
            );
            this.rowNodeInfo.pstNodeInputStream.readCompletely(bitmap);
            const id = this.rowNodeInfo.seekAndReadLong(
                long.fromNumber(currentValueArrayStart),
                4
            );

            // Put into the item map as PidTagLtpRowId (0x67F2)
            let item: PSTTableItem = new PSTTableItem();
            item.itemIndex = -1;
            item.entryValueType = 3;
            item.entryType = long.fromNumber(0x67f2);
            item.entryValueReference = id.toNumber();
            item.isExternalValueReference = true;
            currentItem.set(item.entryType.toNumber(), item);

            let col = 0;
            if (this.overrideCol > -1) {
                col = this.overrideCol - 1;
            }
            while (col < this.numColumns - 1) {
                col++;

                // Does this column exist for this row?
                const bitIndex = Math.trunc(
                    this.columnDescriptors[col].iBit / 8
                );
                const bit = this.columnDescriptors[col].iBit % 8;
                if (
                    bitIndex >= bitmap.length ||
                    (bitmap[bitIndex] & (1 << bit)) === 0
                ) {
                    // Column doesn't exist
                    continue;
                }

                item = new PSTTableItem();
                item.itemIndex = col;
                item.entryValueType = this.columnDescriptors[col].type;
                item.entryType = long.fromNumber(
                    this.columnDescriptors[col].id
                );
                item.entryValueReference = 0;

                switch (this.columnDescriptors[col].cbData) {
                    case 1: // Single byte data
                        item.entryValueReference =
                            this.rowNodeInfo
                                .seekAndReadLong(
                                    long.fromNumber(
                                        currentValueArrayStart +
                                            this.columnDescriptors[col].ibData
                                    ),
                                    1
                                )
                                .toNumber() & 0xff;
                        item.isExternalValueReference = true;
                        break;

                    case 2: // Two byte data
                        item.entryValueReference =
                            this.rowNodeInfo
                                .seekAndReadLong(
                                    long.fromNumber(
                                        currentValueArrayStart +
                                            this.columnDescriptors[col].ibData
                                    ),
                                    2
                                )
                                .toNumber() & 0xffff;
                        item.isExternalValueReference = true;
                        break;

                    case 8: // 8 byte data
                        item.data = Buffer.alloc(8);
                        this.rowNodeInfo.pstNodeInputStream.seek(
                            long.fromNumber(
                                this.rowNodeInfo.startOffset +
                                    currentValueArrayStart +
                                    this.columnDescriptors[col].ibData
                            )
                        );
                        this.rowNodeInfo.pstNodeInputStream.readCompletely(
                            item.data
                        );
                        break;

                    default:
                        // Four byte data
                        item.entryValueReference = this.rowNodeInfo
                            .seekAndReadLong(
                                long.fromNumber(
                                    currentValueArrayStart +
                                        this.columnDescriptors[col].ibData
                                ),
                                4
                            )
                            .toNumber();
                        if (
                            this.columnDescriptors[col].type === 0x0003 ||
                            this.columnDescriptors[col].type === 0x0004 ||
                            this.columnDescriptors[col].type === 0x000a
                        ) {
                            // True 32bit data
                            item.isExternalValueReference = true;
                            break;
                        }

                        // Variable length data so it's an hnid
                        if ((item.entryValueReference & 0x1f) !== 0) {
                            // Some kind of external reference...
                            item.isExternalValueReference = true;
                            break;
                        }

                        if (item.entryValueReference === 0) {
                            item.data = Buffer.alloc(0);
                            break;
                        } else {
                            const entryInfo = this.getNodeInfo(
                                item.entryValueReference
                            );
                            if (entryInfo) {
                                item.data = Buffer.alloc(entryInfo.length());
                                entryInfo.pstNodeInputStream.seek(
                                    long.fromNumber(entryInfo.startOffset)
                                );
                                entryInfo.pstNodeInputStream.readCompletely(
                                    item.data
                                );
                            }
                        }
                        break;
                }
                currentItem.set(item.entryType.toNumber(), item);
            }
            itemList[dataSetNumber] = currentItem;
            dataSetNumber++;
            currentValueArrayStart += this.TCI_bm;
        }
        if (setLocalList) {
            this.items = itemList;
        }
        return itemList;
    }

    /**
     * Get the number of rows.
     */
    public get rowCount(): number {
        return this.numberOfDataSets;
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        return this;
    }
}
