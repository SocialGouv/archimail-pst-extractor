import type { DescriptorIndexNode } from "./DescriptorIndexNode";
import type { PSTFile } from "./PSTFile";
import { PSTObject } from "./PSTObject";

export class PSTMessageStore extends PSTObject {
  /**
   * Creates an instance of PSTMessageStore.
   * Not much use other than to get the "name" of the PST file.
   * @param {PSTFile} pstFile
   * @param {DescriptorIndexNode} descriptorIndexNode
   * @memberof PSTMessageStore
   */
  constructor(pstFile: PSTFile, descriptorIndexNode: DescriptorIndexNode) {
    super(pstFile, descriptorIndexNode);
  }
}
