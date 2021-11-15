/* eslint-disable @typescript-eslint/no-explicit-any */
import type { DescriptorIndexNode } from "./DescriptorIndexNode";
import { OutlookProperties } from "./OutlookProperties";
import type { PSTDescriptorItem } from "./PSTDescriptorItem";
import type { PSTFile } from "./PSTFile";
import { PSTMessage } from "./PSTMessage";
import type { PSTTableBC } from "./PSTTableBC";

export class PSTActivity extends PSTMessage {
  /**
   * Creates an instance of PSTActivity.  Represents Journal entries, class IPM.Activity.
   * https://msdn.microsoft.com/en-us/library/office/aa204771(v=office.11).aspx
   * @param {PSTFile} pstFile
   * @param {DescriptorIndexNode} descriptorIndexNode
   * @param {PSTTableBC} [table]
   * @param {Map<number, PSTDescriptorItem>} [localDescriptorItems]
   * @memberof PSTActivity
   */
  constructor(
    pstFile: PSTFile,
    descriptorIndexNode: DescriptorIndexNode,
    table?: PSTTableBC,
    localDescriptorItems?: Map<number, PSTDescriptorItem>
  ) {
    super(pstFile, descriptorIndexNode, table, localDescriptorItems);
  }

  /**
   * Contains the display name of the journaling application (for example, "MSWord"), and is typically a free-form attribute of a journal message, usually a string.
   * https://msdn.microsoft.com/en-us/library/office/cc839662.aspx
   * @readonly
   * @type {string}
   * @memberof PSTActivity
   */
  public get logType(): string {
    return this.getStringItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogType,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Represents the start date and time for the journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc842339.aspx
   * @readonly
   * @type {Date}
   * @memberof PSTActivity
   */
  public get logStart(): Date | null {
    return this.getDateItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogStart,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Represents the duration, in minutes, of a journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc765536.aspx
   * @readonly
   * @type {number}
   * @memberof PSTActivity
   */
  public get logDuration(): number {
    return this.getIntItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogDuration,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Represents the end date and time for the journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc839572.aspx
   * @readonly
   * @type {Date}
   * @memberof PSTActivity
   */
  public get logEnd(): Date | null {
    return this.getDateItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogEnd,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Contains metadata about the journal.
   * https://msdn.microsoft.com/en-us/library/office/cc815433.aspx
   * @readonly
   * @type {number}
   * @memberof PSTActivity
   */
  public get logFlags(): number {
    return this.getIntItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogFlags,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Indicates whether the document was printed during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc839873.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentPrinted(): boolean {
    return this.getBooleanItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentPrinted,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Indicates whether the document was saved during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc815488.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentSaved(): boolean {
    return this.getBooleanItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentSaved,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Indicates whether the document was sent to a routing recipient during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc839558.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentRouted(): boolean {
    return this.getBooleanItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentRouted,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Indicates whether the document was sent by e-mail or posted to a server folder during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc815353.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentPosted(): boolean {
    return this.getBooleanItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentPosted,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * Describes the activity that is being recorded.
   * https://msdn.microsoft.com/en-us/library/office/cc815500.aspx
   * @readonly
   * @type {string}
   * @memberof PSTActivity
   */
  public get logTypeDesc(): string {
    return this.getStringItem(
      this.pstFile.getNameToIdMapItem(
        OutlookProperties.PidLidLogTypeDesc,
        OutlookProperties.PSETID_Log
      )
    );
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   * @memberof PSTActivity
   */
  public toJSON(): any {
    const clone = Object.assign(
      {
        importance: this.importance,
        isDocumentPrinted: this.isDocumentPrinted,
        isDocumentRouted: this.isDocumentRouted,
        logDuration: this.logDuration,
        isDocumentPosted: this.isDocumentPosted,
        logEnd: this.logEnd,
        isDocumentSaved: this.isDocumentSaved,
        logType: this.logType,
        logFlags: this.logFlags,
        messageClass: this.messageClass,
        logStart: this.logStart,
        subject: this.subject,
        logTypeDesc: this.logTypeDesc,
        transportMessageHeaders: this.transportMessageHeaders,
      },
      this
    );
    return clone;
  }
}
