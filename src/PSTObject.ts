/* eslint-disable @typescript-eslint/no-explicit-any */
import * as long from "long";

import type { DescriptorIndexNode } from "./DescriptorIndexNode";
import type { OffsetIndexItem } from "./OffsetIndexItem";
import { OutlookProperties } from "./OutlookProperties";
import type { PSTDescriptorItem } from "./PSTDescriptorItem";
import type { PSTFile } from "./PSTFile";
import { PSTNodeInputStream } from "./PSTNodeInputStream";
import { PSTTableBC } from "./PSTTableBC";
import type { PSTTableItem } from "./PSTTableItem";
import { PSTUtil } from "./PSTUtil";

export abstract class PSTObject {
  protected pstFile: PSTFile;

  protected descriptorIndexNode: DescriptorIndexNode | null = null;

  protected localDescriptorItems: Map<number, PSTDescriptorItem> | null = null;

  private pstTableBC: PSTTableBC | null = null;

  protected pstTableItems: Map<number, PSTTableItem> | null = null;

  /**
   * Creates an instance of PSTObject, the root class of most PST Items.
   */
  constructor(
    pstFile: PSTFile,
    descriptorIndexNode?: DescriptorIndexNode,
    pstTableItems?: Map<number, PSTTableItem>
  ) {
    this.pstFile = pstFile;

    if (descriptorIndexNode) {
      this.loadDescriptor(descriptorIndexNode);
    }
    if (pstTableItems) {
      this.pstTableItems = pstTableItems;
    }
  }

  /**
   * Load a descriptor from the PST.
   */
  private loadDescriptor(descriptorIndexNode: DescriptorIndexNode): void {
    this.descriptorIndexNode = descriptorIndexNode;

    // get the table items for this descriptor
    const offsetIndexItem: OffsetIndexItem = this.pstFile.getOffsetIndexNode(
      descriptorIndexNode.dataOffsetIndexIdentifier
    );
    const pstNodeInputStream: PSTNodeInputStream = new PSTNodeInputStream(
      this.pstFile,
      offsetIndexItem
    );
    this.pstTableBC = new PSTTableBC(pstNodeInputStream);
    this.pstTableItems = this.pstTableBC.getItems();

    if (
      descriptorIndexNode.localDescriptorsOffsetIndexIdentifier.notEquals(
        long.ZERO
      )
    ) {
      this.localDescriptorItems = this.pstFile.getPSTDescriptorItems(
        descriptorIndexNode.localDescriptorsOffsetIndexIdentifier
      );
    }
  }

  /**
   * Get table items.
   */
  protected prePopulate(
    folderIndexNode: DescriptorIndexNode | null,
    pstTableBC: PSTTableBC,
    localDescriptorItems?: Map<number, PSTDescriptorItem>
  ): void {
    this.descriptorIndexNode = folderIndexNode;
    this.pstTableItems = pstTableBC.getItems();
    this.pstTableBC = pstTableBC;
    this.localDescriptorItems = localDescriptorItems
      ? localDescriptorItems
      : null;
  }

  /**
   * Get the descriptor identifier for this item which can be used for loading objects
   * through detectAndLoadPSTObject(PSTFile theFile, long descriptorIndex)
   */
  public get descriptorNodeId(): long {
    // Prevent null pointer exceptions for embedded messages
    if (this.descriptorIndexNode != null) {
      return long.fromNumber(this.descriptorIndexNode.descriptorIdentifier);
    } else {
      return long.ZERO;
    }
  }

  /**
   * Get the node type for the descriptor id.
   */
  public getNodeType(descriptorIdentifier?: number): number {
    if (descriptorIdentifier) {
      return descriptorIdentifier & 0x1f;
    } else if (this.descriptorIndexNode) {
      return this.descriptorIndexNode.descriptorIdentifier & 0x1f;
    } else {
      return -1;
    }
  }

  /**
   * Get a number.
   */
  public getIntItem(identifier: number, defaultValue?: number): number {
    if (!defaultValue) {
      defaultValue = 0;
    }
    if (this.pstTableItems && this.pstTableItems.has(identifier)) {
      const item = this.pstTableItems.get(identifier);
      if (item) {
        return item.entryValueReference;
      }
    }
    return defaultValue;
  }

  /**
   * Get a boolean.
   */
  public getBooleanItem(identifier: number, defaultValue?: boolean): boolean {
    if (defaultValue === undefined) {
      defaultValue = false;
    }
    if (this.pstTableItems && this.pstTableItems.has(identifier)) {
      const item = this.pstTableItems.get(identifier);
      if (item) {
        return item.entryValueReference != 0;
      }
    }
    return defaultValue;
  }

  /**
   * Get a double.
   */
  public getDoubleItem(identifier: number, defaultValue?: number): number {
    if (defaultValue === undefined) {
      defaultValue = 0;
    }
    if (this.pstTableItems && this.pstTableItems.has(identifier)) {
      const item = this.pstTableItems.get(identifier);
      if (item) {
        const longVersion: long = PSTUtil.convertLittleEndianBytesToLong(
          item.data
        );

        // interpret {low, high} signed 32 bit integers as double
        return new Float64Array(
          new Int32Array([longVersion.low, longVersion.high]).buffer
        )[0];
      }
    }
    return defaultValue;
  }

  /**
   * Get a long.
   */
  public getLongItem(identifier: number, defaultValue?: long): long {
    if (defaultValue === undefined) {
      defaultValue = long.ZERO;
    }
    if (this.pstTableItems && this.pstTableItems.has(identifier)) {
      const item = this.pstTableItems.get(identifier);
      if (item && item.entryValueType == 0x0003) {
        // we are really just an int
        return long.fromNumber(item.entryValueReference);
      } else if (item && item.entryValueType == 0x0014) {
        // we are a long
        if (item.data != null && item.data.length == 8) {
          return PSTUtil.convertLittleEndianBytesToLong(item.data, 0, 8);
        } else {
          console.error(
            `PSTObject::getLongItem Invalid data length for long id ${identifier}`
          );
          // Return the default value for now...
        }
      }
    }
    return defaultValue;
  }

  /**
   * Get a string.
   */
  public getStringItem(
    identifier: number,
    stringType?: number,
    codepage?: string
  ): string {
    if (!stringType) {
      stringType = 0;
    }
    const item = this.pstTableItems
      ? this.pstTableItems.get(identifier)
      : undefined;
    if (item) {
      if (!codepage) {
        codepage = this.stringCodepage;
      }

      // get the string type from the item if not explicitly set
      if (!stringType) {
        stringType = item.entryValueType;
      }

      // see if there is a descriptor entry
      if (!item.isExternalValueReference) {
        return PSTUtil.createJavascriptString(item.data, stringType, codepage);
      }

      if (
        this.localDescriptorItems != null &&
        this.localDescriptorItems.has(item.entryValueReference)
      ) {
        // we have a hit!
        const descItem = this.localDescriptorItems.get(
          item.entryValueReference
        );

        try {
          const data = descItem ? descItem.getData() : null;
          if (data == null) {
            return "";
          }

          return PSTUtil.createJavascriptString(data, stringType, codepage);
        } catch (err) {
          console.error(
            `PSTObject::getStringItem error decoding string\n${err}`
          );
          return "";
        }
      }
    }
    return "";
  }

  /**
   * Get a codepage.
   */
  public get stringCodepage(): string | undefined {
    // try and get the codepage
    let cpItem = this.pstTableItems ? this.pstTableItems.get(0x3ffd) : null; // PidTagMessageCodepage
    if (cpItem == null) {
      cpItem = this.pstTableItems ? this.pstTableItems.get(0x3fde) : null; // PidTagInternetCodepage
    }
    if (cpItem != null) {
      return PSTUtil.getInternetCodePageCharset(cpItem.entryValueReference);
    }
    return "";
  }

  /**
   * Get a date.
   */
  public getDateItem(identifier: number): Date | null {
    if (this.pstTableItems && this.pstTableItems.has(identifier)) {
      const item = this.pstTableItems.get(identifier);
      if (item && item.data.length == 0) {
        return new Date(0);
      }
      if (item) {
        const hi = PSTUtil.convertLittleEndianBytesToLong(item.data, 4, 8);
        const low = PSTUtil.convertLittleEndianBytesToLong(item.data, 0, 4);
        return PSTUtil.filetimeToDate(hi, low);
      }
    }
    return null;
  }

  /**
   * Get a blob.
   */
  public getBinaryItem(identifier: number): Buffer | null {
    if (this.pstTableItems && this.pstTableItems.has(identifier)) {
      const item = this.pstTableItems.get(identifier);
      if (item && item.entryValueType == 0x0102) {
        if (!item.isExternalValueReference) {
          return item.data;
        }
        if (
          this.localDescriptorItems != null &&
          this.localDescriptorItems.has(item.entryValueReference)
        ) {
          // we have a hit!
          const descItem = this.localDescriptorItems.get(
            item.entryValueReference
          );
          try {
            return descItem ? descItem.getData() : null;
          } catch (err) {
            console.error(`PSTObject::Exception reading binary item\n${err}`);
            throw err;
          }
        }
      }
    }
    return null;
  }

  /**
   * Get the display name of this object.
   * https://msdn.microsoft.com/en-us/library/office/cc842383.aspx
   */
  public get displayName(): string {
    return this.getStringItem(OutlookProperties.PR_DISPLAY_NAME);
  }

  /**
   * JSON the object.
   */
  public toJSON(): any {
    return this;
  }
}
