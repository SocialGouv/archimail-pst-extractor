import long from "long";

import { PSTUtil } from ".";
import type { DescriptorIndexNode } from "./DescriptorIndexNode";
import { LZFu } from "./LZFu";
import { OutlookProperties } from "./OutlookProperties";
import { PSTAttachment } from "./PSTAttachment";
import type { PSTDescriptorItem } from "./PSTDescriptorItem";
import { PSTFile } from "./PSTFile";
import { PSTNodeInputStream } from "./PSTNodeInputStream";
import { PSTObject } from "./PSTObject";
import { PSTRecipient } from "./PSTRecipient";
import { PSTTable7C } from "./PSTTable7C";
import { PSTTableBC } from "./PSTTableBC";
import type { MessageClass } from "./PSTUtil";

/* eslint-disable @typescript-eslint/naming-convention */
enum PidTagMessageFlags {
    MSGFLAG_READ = 0x01,
    MSGFLAG_UNMODIFIED = 0x02,
    MSGFLAG_SUBMIT = 0x04,
    MSGFLAG_UNSENT = 0x08,
    MSGFLAG_HASATTACH = 0x10,
    MSGFLAG_FROMME = 0x20,
    MSGFLAG_ASSOCIATED = 0x40,
    MSGFLAG_RESEND = 0x80,
}
/* eslint-enable @typescript-eslint/naming-convention */

export class PSTMessage extends PSTObject {
    public static IMPORTANCE_LOW = 0;

    public static IMPORTANCE_NORMAL = 1;

    public static IMPORTANCE_HIGH = 2;

    public static RECIPIENT_TYPE_TO = 1;

    public static RECIPIENT_TYPE_CC = 2;

    private recipientTable: PSTTable7C | null = null;

    private attachmentTable: PSTTable7C | null = null;

    private recipients?: PSTRecipient[];

    /**
     * Creates an instance of PSTMessage. PST Message contains functions that are common across most MAPI objects.
     * Note that many of these functions may not be applicable for the item in question,
     * however there seems to be no hard and fast outline for what properties apply to which
     * objects. For properties where no value is set, a blank value is returned (rather than
     * an exception being raised).
     */
    constructor(
        pstFile: PSTFile,
        descriptorIndexNode: DescriptorIndexNode,
        pstTableBC?: PSTTableBC,
        localDescriptorItems?: Map<number, PSTDescriptorItem>
    ) {
        super(pstFile, descriptorIndexNode);
        if (pstTableBC) {
            // pre-populate folder object with values
            this.prePopulate(
                descriptorIndexNode,
                pstTableBC,
                localDescriptorItems
            );
        }
    }

    /**
     * Get specific recipient.
     * @deprecated
     */
    public getRecipient(recipientNumber: number): PSTRecipient {
        return this.getRecipients()[recipientNumber];
    }

    public getRecipients(): PSTRecipient[] {
        this.processRecipients();
        if (!this.recipientTable) {
            return [];
            // throw new Error("PSTMessage::getRecipient recipientTable is null");
        }
        return (this.recipients =
            this.recipients ??
            this.recipientTable
                .getItems()
                .map(
                    (recipientDetails) =>
                        new PSTRecipient(this.pstFile, recipientDetails)
                ));
    }

    /**
     * Get specific attachment from table using index.
     */
    public getAttachment(attachmentNumber: number): PSTAttachment {
        this.processAttachments();

        let attachmentCount = 0;
        if (this.attachmentTable !== null) {
            attachmentCount = this.attachmentTable.rowCount;
        }
        if (!this.attachmentTable) {
            throw new Error(
                "PSTMessage::getAttachment attachmentTable is null"
            );
        }
        if (!this.localDescriptorItems) {
            throw new Error(
                "PSTMessage::getAttachment localDescriptorItems is null"
            );
        }
        if (attachmentNumber >= attachmentCount) {
            throw new Error(
                `PSTMessage::getAttachment unable to fetch attachment number ${attachmentNumber}`
            );
        }

        // we process the C7 table here, basically we just want the attachment local descriptor...
        const attachmentDetails =
            this.attachmentTable.getItems()[attachmentNumber];
        const attachmentTableItem = attachmentDetails.get(0x67f2);
        if (!attachmentTableItem) {
            throw new Error(
                "PSTMessage::getAttachment attachmentTableItem is null"
            );
        }
        const descriptorItemId = attachmentTableItem.entryValueReference;
        if (!descriptorItemId) {
            throw new Error(
                "PSTMessage::getAttachment descriptorItemId is null"
            );
        }

        // get the local descriptor for the attachmentDetails table.
        const descriptorItem = this.localDescriptorItems.get(descriptorItemId);
        if (!descriptorItem) {
            throw new Error("PSTMessage::getAttachment descriptorItem is null");
        }

        // try and decode it
        const attachmentData: Buffer = descriptorItem.getData();
        if (attachmentData.length > 0) {
            const attachmentDetailsTable: PSTTableBC = new PSTTableBC(
                new PSTNodeInputStream(this.pstFile, descriptorItem)
            );

            // create our all-precious attachment object.
            // note that all the information that was in the c7 table is
            // repeated in the eb table in attachment data.
            // so no need to pass it...
            let attachmentDescriptorItems: Map<number, PSTDescriptorItem> =
                new Map();
            if (descriptorItem.subNodeOffsetIndexIdentifier > 0) {
                attachmentDescriptorItems = this.pstFile.getPSTDescriptorItems(
                    long.fromNumber(descriptorItem.subNodeOffsetIndexIdentifier)
                );
            }
            return new PSTAttachment(
                this.pstFile,
                attachmentDetailsTable,
                attachmentDescriptorItems
            );
        }

        throw new Error(
            `PSTMessage::getAttachment unable to fetch attachment number ${attachmentNumber}, unable to read attachment details table`
        );
    }

    /**
     * JSON stringify the object properties.  Large fields (like body) aren't included.
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                acknowledgementMode: this.acknowledgementMode,
                addrType: this.addrType,
                attrHidden: this.attrHidden,
                bloriginalSubjectah: this.originalSubject,
                bodyPrefix: this.bodyPrefix,
                clientSubmitTime: this.clientSubmitTime,
                colorCategories: this.colorCategories,
                comment: this.comment,
                conversationId: this.conversationId,
                conversationTopic: this.conversationTopic,
                creationTime: this.creationTime,
                deleteAfterSubmit: this.deleteAfterSubmit,
                displayBCC: this.displayBCC,
                displayCC: this.displayCC,
                displayTo: this.displayTo,
                emailAddress: this.emailAddress,
                hasAttachments: this.hasAttachments,
                iconIndex: this.iconIndex,
                importance: this.importance,
                inReplyToId: this.inReplyToId,
                internetArticleNumber: this.internetArticleNumber,
                internetMessageId: this.internetMessageId,
                isAssociated: this.isAssociated,
                isConversationIndexTracking: this.isConversationIndexTracking,
                isFromMe: this.isFromMe,
                isNonReceiptNotificationRequested:
                    this.isNonReceiptNotificationRequested,
                isOriginatorNonDeliveryReportRequested:
                    this.isOriginatorNonDeliveryReportRequested,
                isRTFInSync: this.isRTFInSync,
                isRead: this.isRead,
                isReplyRequested: this.isReplyRequested,
                isResent: this.isResent,
                isSubmitted: this.isSubmitted,
                isUnmodified: this.isUnmodified,
                isUnsent: this.isUnsent,
                lastVerbExecutionTime: this.lastVerbExecutionTime,
                messageCcMe: this.messageCcMe,
                messageClass: this.messageClass,
                messageDeliveryTime: this.messageDeliveryTime,
                messageRecipMe: this.messageRecipMe,
                messageSize: this.messageSize,
                messageToMe: this.messageToMe,
                modificationTime: this.modificationTime,
                nextSendAcct: this.nextSendAcct,
                numberOfAttachments: this.numberOfAttachments,
                numberOfRecipients: this.numberOfRecipients,
                objectType: this.objectType,
                originalDisplayBcc: this.originalDisplayBcc,
                originalDisplayCc: this.originalDisplayCc,
                originalDisplayTo: this.originalDisplayTo,
                originalSensitivity: this.originalSensitivity,
                originatorDeliveryReportRequested:
                    this.originatorDeliveryReportRequested,
                primarySendAccount: this.primarySendAccount,
                rcvdRepresentingAddrtype: this.rcvdRepresentingAddrtype,
                rcvdRepresentingEmailAddress: this.rcvdRepresentingEmailAddress,
                rcvdRepresentingName: this.rcvdRepresentingName,
                readReceiptRequested: this.readReceiptRequested,
                receivedByAddress: this.receivedByAddress,
                receivedByAddressType: this.receivedByAddressType,
                receivedByName: this.receivedByName,
                recipientReassignmentProhibited:
                    this.recipientReassignmentProhibited,
                recipientTable: this.recipientTable,
                recipientType: this.recipientType,
                reminderDelta: this.reminderDelta,
                reminderSet: this.reminderSet,
                replyRecipientNames: this.replyRecipientNames,
                responseRequested: this.responseRequested,
                responsibility: this.responsibility,
                returnPath: this.returnPath,
                rtfSyncBodyCRC: this.rtfSyncBodyCRC,
                rtfSyncBodyCount: this.rtfSyncBodyCount,
                rtfSyncBodyTag: this.rtfSyncBodyTag,
                rtfSyncPrefixCount: this.rtfSyncPrefixCount,
                rtfSyncTrailingCount: this.rtfSyncTrailingCount,
                senderAddrtype: this.senderAddrtype,
                senderEmailAddress: this.senderEmailAddress,
                senderName: this.senderName,
                senderSmtpEmailAddress: this.senderSmtpEmailAddress,
                sensitivity: this.sensitivity,
                sentRepresentingAddressType: this.sentRepresentingAddressType,
                sentRepresentingEmailAddress: this.sentRepresentingEmailAddress,
                sentRepresentingName: this.sentRepresentingName,
                subject: this.subject,
                taskDueDate: this.taskDueDate,
                taskStartDate: this.taskStartDate,
                transportMessageHeaders: this.transportMessageHeaders,
                urlCompName: this.urlCompName,
            },
            this
        );
        return clone;
    }

    /**
     * The message is marked as having been read.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isRead(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_READ) !=
            0
        );
    }

    /**
     * The outgoing message has not been modified since the first time that it was saved; the incoming message has not been modified since it was delivered.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isUnmodified(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_UNMODIFIED) !=
            0
        );
    }

    /**
     * The message is marked for sending as a result of a call to the RopSubmitMessage ROP
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isSubmitted(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_SUBMIT) !=
            0
        );
    }

    /**
     * The message is still being composed. It is saved, but has not been sent.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isUnsent(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_UNSENT) !=
            0
        );
    }

    /**
     * The message has at least one attachment.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get hasAttachments(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_HASATTACH) !=
            0
        );
    }

    /**
     * The user receiving the message was also the user who sent the message.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isFromMe(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_FROMME) !=
            0
        );
    }

    /**
     * The message is an FAI message.  An FAI Message object is used to store a variety of settings and
     * auxiliary data, including forms, views, calendar options, favorites, and category lists.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isAssociated(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_ASSOCIATED) !=
            0
        );
    }

    /**
     * The message includes a request for a resend operation with a nondelivery report.
     * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
     */
    public get isResent(): boolean {
        return (
            (this.getIntItem(OutlookProperties.PR_MESSAGE_FLAGS) &
                PidTagMessageFlags.MSGFLAG_RESEND) !=
            0
        );
    }

    /**
     * Find, extract and load up all of the attachments in this email
     */
    private processRecipients(): void {
        if (this.recipientTable) {
            return;
        }

        try {
            const recipientTableKey = 0x0692;
            if (this.localDescriptorItems?.has(recipientTableKey)) {
                const item = this.localDescriptorItems.get(recipientTableKey);
                let descriptorItems = new Map<number, PSTDescriptorItem>();
                if (item && item.subNodeOffsetIndexIdentifier > 0) {
                    descriptorItems = this.pstFile.getPSTDescriptorItems(
                        long.fromNumber(item.subNodeOffsetIndexIdentifier)
                    );
                }
                this.recipientTable = new PSTTable7C(
                    new PSTNodeInputStream(this.pstFile, item),
                    descriptorItems
                );
            }
        } catch (err: unknown) {
            console.error(`PSTMessage::processRecipients\n${err}`);
            this.recipientTable = null;
        }
    }

    /**
     * Get the recipients table.
     */
    public get numberOfRecipients(): number {
        this.processRecipients();
        return this.recipientTable?.rowCount ?? 0;
    }

    /**
     * Contains TRUE if a message sender wants notification of non-receipt for a specified recipient.
     * https://msdn.microsoft.com/en-us/library/office/cc979208.aspx
     */
    public get isNonReceiptNotificationRequested(): boolean {
        return (
            this.getIntItem(
                OutlookProperties.PR_NON_RECEIPT_NOTIFICATION_REQUESTED
            ) !== 0
        );
    }

    /**
     * Contains TRUE if a message sender wants notification of non-deliver for a specified recipient.
     * https://msdn.microsoft.com/en-us/library/ms987568(v=exchg.65).aspx
     */
    public get isOriginatorNonDeliveryReportRequested(): boolean {
        return (
            this.getIntItem(
                OutlookProperties.PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED
            ) !== 0
        );
    }

    /**
     * Contains the recipient type for a message recipient.
     * https://msdn.microsoft.com/en-us/library/office/cc839620.aspx
     */
    public get recipientType(): number {
        return this.getIntItem(OutlookProperties.PR_RECIPIENT_TYPE);
    }

    /*
        Body (plain text, RTF, HTML)
    */
    /**
     * Plain text message body.
     * https://msdn.microsoft.com/en-us/library/office/cc765874.aspx
     */
    public get body(): string {
        const codepage = this.getCodepage();
        if (codepage) {
            return this.getStringItem(OutlookProperties.PR_BODY, 0, codepage);
        } else {
            return this.getStringItem(OutlookProperties.PR_BODY);
        }
    }

    /**
     * Plain text body prefix.
     */
    public get bodyPrefix(): string {
        return this.getStringItem(0x6619);
    }

    /**
     * Contains the Rich Text Format (RTF) version of the message text, usually in compressed form.
     * https://technet.microsoft.com/en-us/library/cc815911
     */
    public get bodyRTF(): string {
        // do we have an entry for it?
        if (this.pstTableItems?.has(0x1009)) {
            // is it a reference?
            const item = this.pstTableItems.get(0x1009);
            if (item && item.data.length > 0) {
                return LZFu.decode(item.data);
            }
            const ref = item ? item.entryValueReference : null;
            if (ref) {
                const descItem = this.localDescriptorItems
                    ? this.localDescriptorItems.get(ref)
                    : null;
                if (descItem) {
                    return LZFu.decode(descItem.getData());
                }
            }
        }
        return "";
    }

    /**
     * Contains the cyclical redundancy check (CRC) computed for the message text.
     * https://technet.microsoft.com/en-us/library/cc815532(v=office.15).aspx
     */
    public get rtfSyncBodyCRC(): number {
        return this.getIntItem(OutlookProperties.PR_RTF_SYNC_BODY_CRC);
    }

    /**
     * Contains a count of the significant characters of the message text.
     * https://msdn.microsoft.com/en-us/library/windows/desktop/cc842324.aspx
     */
    public get rtfSyncBodyCount(): number {
        return this.getIntItem(OutlookProperties.PR_RTF_SYNC_BODY_COUNT);
    }

    /**
     * Contains significant characters that appear at the beginning of the message text.
     * https://technet.microsoft.com/en-us/library/cc815400(v=office.15).aspx
     */
    public get rtfSyncBodyTag(): string {
        return this.getStringItem(OutlookProperties.PR_RTF_SYNC_BODY_TAG);
    }

    /**
     * Contains a count of the ignorable characters that appear before the significant characters of the message.
     * https://msdn.microsoft.com/en-us/magazine/cc842437.aspx
     */
    public get rtfSyncPrefixCount(): number {
        return this.getIntItem(OutlookProperties.PR_RTF_SYNC_PREFIX_COUNT);
    }

    /**
     * Contains a count of the ignorable characters that appear after the significant characters of the message.
     * https://msdn.microsoft.com/en-us/magazine/cc765795.aspx
     */
    public get rtfSyncTrailingCount(): number {
        return this.getIntItem(OutlookProperties.PR_RTF_SYNC_TRAILING_COUNT);
    }

    /**
     * Gets codepage to use.
     * TODO - does this work?
     */
    private getCodepage(): string | null | undefined {
        let cpItem = this.pstTableItems
            ? this.pstTableItems.get(OutlookProperties.PR_INTERNET_CPID)
            : null;
        if (cpItem === null) {
            cpItem = this.pstTableItems
                ? this.pstTableItems.get(OutlookProperties.PR_MESSAGE_CODEPAGE)
                : null;
        }
        if (cpItem) {
            return PSTUtil.getInternetCodePageCharset(
                cpItem.entryValueReference
            );
        }
        return null;
    }

    /**
     * Contains the HTML version of the message text.
     */
    public get bodyHTML(): string {
        const codepage = this.getCodepage();
        if (codepage) {
            return this.getStringItem(
                OutlookProperties.PR_BODY_HTML,
                0,
                codepage
            );
        } else {
            return this.getStringItem(OutlookProperties.PR_BODY_HTML);
        }
    }

    /*
        Attachments
    */
    /**
     * Processes table which holds attachments.
     */
    private processAttachments(): void {
        const attachmentTableKey = 0x0671;
        if (
            this.attachmentTable === null &&
            this.localDescriptorItems !== null &&
            this.localDescriptorItems.has(attachmentTableKey)
        ) {
            const item = this.localDescriptorItems.get(attachmentTableKey);
            let descriptorItems: Map<number, PSTDescriptorItem> = new Map();
            if (item && item.subNodeOffsetIndexIdentifier > 0) {
                descriptorItems = this.pstFile.getPSTDescriptorItems(
                    long.fromValue(item.subNodeOffsetIndexIdentifier)
                );
            }
            this.attachmentTable = new PSTTable7C(
                new PSTNodeInputStream(this.pstFile, item),
                descriptorItems
            );
        }
    }

    /**
     * Number of attachments by counting rows in attachment table.
     */
    public get numberOfAttachments(): number {
        try {
            this.processAttachments();
        } catch (err: unknown) {
            console.error(`PSTMessage::numberOfAttachments\n${err}`);
            return 0;
        }
        return this.attachmentTable ? this.attachmentTable.rowCount : 0;
    }

    /*
        Miscellaneous properties
    */
    /**
     * Importance of email (sender determined)
     * https://msdn.microsoft.com/en-us/library/cc815346(v=office.12).aspx
     */
    public get importance(): number {
        return this.getIntItem(
            OutlookProperties.PR_IMPORTANCE,
            PSTMessage.IMPORTANCE_NORMAL
        );
    }

    /**
     * Contains a text string that identifies the sender-defined message class, such as IPM.Note.
     * https://msdn.microsoft.com/en-us/library/office/cc765765.aspx
     */
    public get messageClass(): MessageClass {
        return this.getStringItem(
            OutlookProperties.PR_MESSAGE_CLASS
        ) as MessageClass;
    }

    /**
     * Contains the full subject of a message.
     * https://technet.microsoft.com/en-us/library/cc815720
     */
    public get subject(): string {
        let subject = this.getStringItem(OutlookProperties.PR_SUBJECT);
        if (subject.length >= 2 && subject.charCodeAt(0) === 0x01) {
            if (subject.length === 2) {
                subject = "";
            } else {
                subject = subject.substring(2, subject.length);
            }
        }
        return subject;
    }

    /**
     * Contains the date and time the message sender submitted a message.
     * https://technet.microsoft.com/en-us/library/cc839781
     */
    public get clientSubmitTime(): Date | null {
        return this.getDateItem(OutlookProperties.PR_CLIENT_SUBMIT_TIME);
    }

    /**
     * Contains the display name of the messaging user who receives the message.
     * https://msdn.microsoft.com/en-us/library/office/cc840015.aspx
     */
    public get receivedByName(): string {
        return this.getStringItem(OutlookProperties.PR_RECEIVED_BY_NAME);
    }

    /**
     * Contains the display name for the messaging user represented by the sender.
     * https://msdn.microsoft.com/en-us/library/office/cc842405.aspx
     */
    public get sentRepresentingName(): string {
        return this.getStringItem(OutlookProperties.PR_SENT_REPRESENTING_NAME);
    }

    /**
     * Contains the address type for the messaging user who is represented by the sender.
     * https://msdn.microsoft.com/en-us/library/office/cc839677.aspx
     */
    public get sentRepresentingAddressType(): string {
        return this.getStringItem(
            OutlookProperties.PR_SENT_REPRESENTING_ADDRTYPE
        );
    }

    /**
     * Contains the e-mail address for the messaging user who is represented by the sender.
     * https://msdn.microsoft.com/en-us/library/office/cc839552.aspx
     */
    public get sentRepresentingEmailAddress(): string {
        return this.getStringItem(
            OutlookProperties.PR_SENT_REPRESENTING_EMAIL_ADDRESS
        );
    }

    /**
     * Contains the Simple Mail Transport Protocol (SMTP) email address for the messaging user who is represented by the sender.
     * https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property
     */
    public get sentRepresentingSmtpAddress(): string {
        return this.getStringItem(
            OutlookProperties.PR_SENT_REPRESENTING_SMTP_ADDRESS
        );
    }

    /**
     * Contains the topic of the first message in a conversation thread.
     * https://technet.microsoft.com/en-us/windows/cc839841
     */
    public get conversationTopic(): string {
        return this.getStringItem(OutlookProperties.PR_CONVERSATION_TOPIC);
    }

    /**
     * Contains the e-mail address type, such as SMTP, for the messaging user who actually receives the message.
     * https://technet.microsoft.com/en-us/library/cc765641(v=office.14)
     */
    public get receivedByAddressType(): string {
        return this.getStringItem(OutlookProperties.PR_RECEIVED_BY_ADDRTYPE);
    }

    /**
     * Contains the e-mail address for the messaging user who receives the message.
     * https://technet.microsoft.com/en-us/library/cc839550(v=office.14)
     */
    public get receivedByAddress(): string {
        return this.getStringItem(
            OutlookProperties.PR_RECEIVED_BY_EMAIL_ADDRESS
        );
    }

    /**
     * Contains the SMTP address for the messaging user who receives the message.
     */
    public get receivedBySmtpAddress(): string {
        return this.getStringItem(
            OutlookProperties.PR_RECEIVED_BY_SMTP_ADDRESS
        );
    }

    /**
     * Contains transport-specific message envelope information.
     * https://technet.microsoft.com/en-us/library/cc815628
     */
    public get transportMessageHeaders(): string {
        return this.getStringItem(
            OutlookProperties.PR_TRANSPORT_MESSAGE_HEADERS
        );
    }

    /**
     * Acknowledgment mode Integer 32-bit signed.
     */
    public get acknowledgementMode(): number {
        return this.getIntItem(0x0001);
    }

    /**
     * Contains TRUE if a message sender requests a delivery report for a particular recipient from the messaging system before the message is placed in the message store.
     * https://msdn.microsoft.com/en-us/library/office/cc765845.aspx
     */
    public get originatorDeliveryReportRequested(): boolean {
        return (
            this.getIntItem(
                OutlookProperties.PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED
            ) !== 0
        );
    }

    /**
     * Contains the relative priority of a message.
     * https://msdn.microsoft.com/en-us/library/office/cc765646.aspx
     */
    public get priority(): number {
        return this.getIntItem(0x0026);
    }

    /**
     * Contains TRUE if a message sender wants the messaging system to generate a read report when the recipient has read a message.
     * https://msdn.microsoft.com/en-us/library/office/cc842094.aspx
     */
    public get readReceiptRequested(): boolean {
        return (
            this.getIntItem(OutlookProperties.PR_READ_RECEIPT_REQUESTED) !== 0
        );
    }

    /**
     * Specifies whether adding additional recipients, when forwarding the message, is prohibited for the e-mail message.
     * https://msdn.microsoft.com/en-us/library/office/cc979216.aspx
     */
    public get recipientReassignmentProhibited(): boolean {
        return (
            this.getIntItem(
                OutlookProperties.PR_RECIPIENT_REASSIGNMENT_PROHIBITED
            ) !== 0
        );
    }

    /**
     * Contains the sensitivity value assigned by the sender of the first version of a message that is, the message before being forwarded or replied to.
     * https://msdn.microsoft.com/en-us/library/cc839694(office.12).aspx
     */
    public get originalSensitivity(): number {
        return this.getIntItem(OutlookProperties.PR_ORIGINAL_SENSITIVITY);
    }

    /**
     * Contains a value that indicates the message sender's opinion of the sensitivity of a message.
     * https://msdn.microsoft.com/en-us/library/office/cc839518.aspx
     */
    public get sensitivity(): number {
        return this.getIntItem(OutlookProperties.PR_SENSITIVITY);
    }

    /**
     * Contains the search key for the messaging user represented by the sender.
     * https://msdn.microsoft.com/en-us/magazine/cc842068.aspx
     */
    public get pidTagSentRepresentingSearchKey(): Buffer | null {
        return this.getBinaryItem(
            OutlookProperties.PR_SENT_REPRESENTING_SEARCH_KEY
        );
    }

    /**
     * Contains the display name for the messaging user who is represented by the receiving user.
     * https://technet.microsoft.com/en-us/library/cc842260.aspx
     */
    public get rcvdRepresentingName(): string {
        return this.getStringItem(OutlookProperties.PR_RCVD_REPRESENTING_NAME);
    }

    /**
     * Contains the subject of an original message for use in a report about the message.
     * https://msdn.microsoft.com/en-us/library/office/cc842182.aspx
     */
    public get originalSubject(): string {
        return this.getStringItem(OutlookProperties.PR_ORIGINAL_SUBJECT);
    }

    /**
     * Contains a list of display names for recipients that are to get a reply.
     * https://msdn.microsoft.com/en-us/library/windows/desktop/cc815850.aspx
     */
    public get replyRecipientNames(): string {
        return this.getStringItem(OutlookProperties.PR_REPLY_RECIPIENT_NAMES);
    }

    /**
     * Contains TRUE if this messaging user is specifically named as a primary (To) recipient of this message and is not part of a distribution list.
     * https://technet.microsoft.com/en-us/library/cc815755
     */
    public get messageToMe(): boolean {
        return this.getIntItem(OutlookProperties.PR_MESSAGE_TO_ME) !== 0;
    }

    /**
     * Contains TRUE if this messaging user is specifically named as a carbon copy (CC) recipient of this message and is not part of a distribution list.
     * https://msdn.microsoft.com/en-us/library/office/cc839713.aspx
     */
    public get messageCcMe(): boolean {
        return this.getIntItem(OutlookProperties.PR_MESSAGE_CC_ME) !== 0;
    }

    /**
     * Contains TRUE if this messaging user is specifically named as a primary (To), carbon copy (CC), or blind carbon copy (BCC) recipient of this message and is not part of a distribution list.
     * https://msdn.microsoft.com/en-us/library/office/cc842268.aspx
     */
    public get messageRecipMe(): boolean {
        return this.getIntItem(OutlookProperties.PR_MESSAGE_RECIP_ME) !== 0;
    }

    /**
     * Contains TRUE if the message sender wants a response to a meeting request.
     * https://msdn.microsoft.com/en-us/library/office/cc839921.aspx
     */
    public get responseRequested(): boolean {
        return this.getBooleanItem(OutlookProperties.PR_RESPONSE_REQUESTED);
    }

    /**
     * Contains the display names of any carbon copy (CC) recipients of the original message.
     * https://msdn.microsoft.com/en-us/magazine/cc815841(v=office.14).aspx
     */
    public get originalDisplayBcc(): string {
        return this.getStringItem(OutlookProperties.PR_ORIGINAL_DISPLAY_BCC);
    }

    /**
     * Contains the display names of any carbon copy (CC) recipients of the original message.
     * https://msdn.microsoft.com/en-us/magazine/cc815841(v=office.14).aspx
     */
    public get originalDisplayCc(): string {
        return this.getStringItem(OutlookProperties.PR_ORIGINAL_DISPLAY_CC);
    }

    /**
     * Contains the display names of the primary (To) recipients of the original message.
     * https://msdn.microsoft.com/en-us/magazine/cc842235(v=office.14).aspx
     */
    public get originalDisplayTo(): string {
        return this.getStringItem(OutlookProperties.PR_ORIGINAL_DISPLAY_TO);
    }

    /**
     * Contains the address type for the messaging user who is represented by the user actually receiving the message.
     * https://msdn.microsoft.com/en-us/library/office/cc842447.aspx
     */
    public get rcvdRepresentingAddrtype(): string {
        return this.getStringItem(
            OutlookProperties.PR_RCVD_REPRESENTING_ADDRTYPE
        );
    }

    /**
     * Contains the e-mail address for the messaging user who is represented by the receiving user.
     * https://msdn.microsoft.com/en-us/library/office/cc815875.aspx
     */
    public get rcvdRepresentingEmailAddress(): string {
        return this.getStringItem(
            OutlookProperties.PR_RCVD_REPRESENTING_EMAIL_ADDRESS
        );
    }

    /**
     * Received representing SMTP address.
     */
    public get rcvdRepresentingSmtpAddress(): string {
        return this.getStringItem(
            OutlookProperties.PR_RCVD_REPRESENTING_SMTP_ADDRESS
        );
    }

    /**
     * Contains TRUE if a message sender requests a reply from a recipient.
     * https://msdn.microsoft.com/en-us/library/office/cc815286.aspx
     */
    public get isReplyRequested(): boolean {
        return this.getIntItem(OutlookProperties.PR_REPLY_REQUESTED) !== 0;
    }

    /**
     * Contains the message sender's entry identifier.
     * https://msdn.microsoft.com/en-us/library/office/cc815625.aspx
     */
    public get senderEntryId(): Buffer | null {
        return this.getBinaryItem(OutlookProperties.PR_SENDER_ENTRYID);
    }

    /**
     * Contains the message sender's display name.
     * https://msdn.microsoft.com/en-us/library/office/cc815457.aspx
     */
    public get senderName(): string {
        return this.getStringItem(OutlookProperties.PR_SENDER_NAME);
    }

    /**
     * Contains the message sender's e-mail address type.
     * https://msdn.microsoft.com/en-us/library/office/cc815748.aspx
     */
    public get senderAddrtype(): string {
        return this.getStringItem(OutlookProperties.PR_SENDER_ADDRTYPE);
    }

    /**
     * Contains the message sender's e-mail address.
     * https://msdn.microsoft.com/en-us/library/office/cc839670.aspx
     */
    public get senderEmailAddress(): string {
        return this.getStringItem(OutlookProperties.PR_SENDER_EMAIL_ADDRESS);
    }

    /**
     * Contains the message sender's SMTP e-mail address.
     * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property
     */
    public get senderSmtpEmailAddress(): string {
        return this.getStringItem(OutlookProperties.PR_SENDER_SMTP_ADDRESS);
    }

    /**
     * Contains the sum, in bytes, of the sizes of all properties on a message object
     * https://technet.microsoft.com/en-us/library/cc842471
     */
    public get messageSize(): long {
        return this.getLongItem(OutlookProperties.PR_MESSAGE_SIZE);
    }

    /**
     * A number associated with an item in a message store.
     * https://msdn.microsoft.com/en-us/library/office/cc815718.aspx
     */
    public get internetArticleNumber(): number {
        return this.getIntItem(OutlookProperties.PR_INTERNET_ARTICLE_NUMBER);
    }

    /**
     * Contains a string that names the first server that is used to send the message.
     * https://msdn.microsoft.com/en-us/library/office/cc815413.aspx
     */
    public get primarySendAccount(): string {
        return this.getStringItem(OutlookProperties.PR_PRIMARY_SEND_ACCOUNT);
    }

    /**
     * Specifies the server that a client is currently attempting to use to send e-mail.
     * https://technet.microsoft.com/en-us/library/cc842327(v=office.14)
     */
    public get nextSendAcct(): string {
        return this.getStringItem(OutlookProperties.PR_NEXT_SEND_ACCT);
    }

    /**
     * Contains the type of an object.
     * https://msdn.microsoft.com/en-us/library/office/cc815487.aspx
     */
    public get objectType(): number {
        return this.getIntItem(OutlookProperties.PR_OBJECT_TYPE);
    }

    /**
     * Contains TRUE if a client application wants MAPI to delete the associated message after submission.
     * https://msdn.microsoft.com/en-us/library/office/cc842353.aspx
     */
    public get deleteAfterSubmit(): boolean {
        return this.getIntItem(OutlookProperties.PR_DELETE_AFTER_SUBMIT) !== 0;
    }

    /**
     * Contains TRUE if some transport provider has already accepted responsibility for delivering the message to this recipient, and FALSE if the MAPI spooler considers that this transport provider should accept responsibility.
     * https://msdn.microsoft.com/en-us/library/office/cc765767.aspx
     */
    public get responsibility(): boolean {
        return this.getIntItem(OutlookProperties.PR_RESPONSIBILITY) !== 0;
    }

    /**
     * Contains TRUE if the PR_RTF_COMPRESSED (PidTagRtfCompressed) property has the same text content as the PR_BODY (PidTagBody) property for this message.
     * https://msdn.microsoft.com/en-us/library/office/cc765844.aspx
     */
    public get isRTFInSync(): boolean {
        return this.getIntItem(OutlookProperties.PR_RTF_IN_SYNC) !== 0;
    }

    /**
     * Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;).
     * https://msdn.microsoft.com/en-us/library/office/cc815730.aspx
     */
    public get displayBCC(): string {
        return this.getStringItem(OutlookProperties.PR_DISPLAY_BCC);
    }

    /**
     * Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;).
     * https://msdn.microsoft.com/en-us/library/office/cc765528.aspx
     */
    public get displayCC(): string {
        return this.getStringItem(OutlookProperties.PR_DISPLAY_CC);
    }

    /**
     * Contains a list of the display names of the primary (To) message recipients, separated by semicolons (;).
     * https://msdn.microsoft.com/en-us/library/office/cc839687.aspx
     */
    public get displayTo(): string {
        return this.getStringItem(OutlookProperties.PR_DISPLAY_TO);
    }

    /**
     * Contains the date and time when a message was delivered.
     * https://msdn.microsoft.com/en-us/library/office/cc841961.aspx
     */
    public get messageDeliveryTime(): Date | null {
        return this.getDateItem(OutlookProperties.PR_MESSAGE_DELIVERY_TIME);
    }

    /**
     * Corresponds to the message ID field as specified in [RFC2822].
     * https://msdn.microsoft.com/en-us/library/office/cc839521.aspx
     */
    public get internetMessageId(): string {
        return this.getStringItem(OutlookProperties.PR_INTERNET_MESSAGE_ID);
    }

    /**
     * Contains the original message's PR_INTERNET_MESSAGE_ID (PidTagInternetMessageId) property value.
     * https://msdn.microsoft.com/en-us/library/office/cc839776.aspx
     */
    public get inReplyToId(): string {
        return this.getStringItem(OutlookProperties.PR_IN_REPLY_TO_ID);
    }

    /**
     * Contains the value of a Multipurpose Internet Mail Extensions (MIME) message's Return-Path header field. The e-mail address of the message's sender.
     * https://msdn.microsoft.com/en-us/library/office/cc765856.aspx
     */
    public get returnPath(): string {
        return this.getStringItem(OutlookProperties.PR_INTERNET_RETURN_PATH);
    }

    /**
     * Contains a number that indicates which icon to use when you display a group of e-mail objects.
     * https://msdn.microsoft.com/en-us/library/office/cc815472.aspx
     */
    public get iconIndex(): number {
        return this.getIntItem(OutlookProperties.PR_ICON_INDEX);
    }

    /**
     * Contains the last verb executed.
     * Todo: Helper methods for each flag.
     * https://msdn.microsoft.com/en-us/library/office/cc841968.aspx
     */
    public get lastVerbExecuted(): number {
        return this.getIntItem(OutlookProperties.PR_LAST_VERB_EXECUTED);
    }

    /**
     * Contains the time when the last verb was executed.
     * https://msdn.microsoft.com/en-us/library/office/cc839918.aspx
     */
    public get lastVerbExecutionTime(): Date | null {
        return this.getDateItem(OutlookProperties.PR_LAST_VERB_EXECUTION_TIME);
    }

    /**
     * The URL component name for a message.
     * https://msdn.microsoft.com/en-us/library/office/cc815653.aspx
     */
    public get urlCompName(): string {
        return this.getStringItem(OutlookProperties.PR_URL_COMP_NAME);
    }

    /**
     * Specifies the hide or show status of a folder.
     * https://msdn.microsoft.com/en-us/library/ee159038(v=exchg.80).aspx
     */
    public get attrHidden(): boolean {
        return this.getIntItem(OutlookProperties.PR_ATTR_HIDDEN) !== 0;
    }

    /**
     * Specifies the date on which the user expects work on the task to begin.
     * https://technet.microsoft.com/en-us/library/cc815922(v=office.12).aspx
     */
    public get taskStartDate(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskStartDate,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Represents the date when the user expects to complete the task.
     * https://technet.microsoft.com/en-us/library/cc839641(v=office.12).aspx
     */
    public get taskDueDate(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskDueDate,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Specifies whether a reminder is set on the object.
     * https://msdn.microsoft.com/en-us/library/office/cc765589.aspx
     */
    public get reminderSet(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidReminderSet,
                OutlookProperties.PSETID_Common
            )
        );
    }

    /**
     * Specifies the interval, in minutes, between the time when the reminder first becomes overdue and the start time of the calendar object.
     * https://msdn.microsoft.com/en-us/library/office/cc765535.aspx
     */
    public get reminderDelta(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidReminderDelta,
                OutlookProperties.PSETID_Common
            )
        );
    }

    /**
     * Color categories
     */
    public get colorCategories(): string[] {
        const keywordCategory: number =
            PSTFile.getPublicStringToIdMapItem("Keywords");

        const categories: string[] = [];
        if (this.pstTableItems?.has(keywordCategory)) {
            try {
                const item = this.pstTableItems.get(keywordCategory);
                if (item && item.data.length === 0) {
                    return [];
                }
                if (item) {
                    const categoryCount: number = item.data[0];
                    if (categoryCount > 0) {
                        const offsets: number[] = [];
                        for (let x = 0; x < categoryCount; x++) {
                            offsets[x] = PSTUtil.convertBigEndianBytesToLong(
                                item.data,
                                x * 4 + 1,
                                (x + 1) * 4 + 1
                            ).toNumber();
                        }
                        for (let x = 0; x < offsets.length - 1; x++) {
                            const start = offsets[x];
                            const end = offsets[x + 1];
                            const length = end - start;
                            const buf: Buffer = Buffer.alloc(length);
                            PSTUtil.arraycopy(item.data, start, buf, 0, length);
                            const name: string = Buffer.from(buf).toString();
                            categories[x] = name;
                        }
                        const start = offsets[offsets.length - 1];
                        const end = item.data.length;
                        const length = end - start;
                        const buf: Buffer = Buffer.alloc(length);
                        PSTUtil.arraycopy(item.data, start, buf, 0, length);
                        const name: string = Buffer.from(buf).toString();
                        categories[categories.length - 1] = name;
                    }
                }
            } catch (err: unknown) {
                console.error(
                    `PSTMessage::colorCategories Unable to decode category data\n${err}`
                );
                throw err;
            }
        }
        return categories;
    }

    /**
     * Contains a computed value derived from other conversation-related properties.
     * https://msdn.microsoft.com/en-us/library/ee204279(v=exchg.80).aspx
     */
    public get conversationId(): Buffer | null {
        return this.getBinaryItem(OutlookProperties.PidTagConversationId);
    }

    /**
     * Indicates whether the GUID portion of the PidTagConversationIndex property (section 2.641) is to be used to compute the PidTagConversationId property (section 2.640).
     * https://msdn.microsoft.com/en-us/library/ee218393(v=exchg.80).aspx
     */
    public get isConversationIndexTracking(): boolean {
        return this.getBooleanItem(
            OutlookProperties.PidTagConversationIndexTracking,
            false
        );
    }

    /**
     * Contains the messaging user's e-mail address.
     * https://msdn.microsoft.com/en-us/library/office/cc842372.aspx
     */
    public get emailAddress(): string {
        return this.getStringItem(OutlookProperties.PR_EMAIL_ADDRESS);
    }

    /**
     * Contains the messaging user's e-mail address type, such as SMTP.
     * https://msdn.microsoft.com/en-us/library/office/cc815548.aspx
     */
    public get addrType(): string {
        return this.getStringItem(OutlookProperties.PR_ADDRTYPE);
    }

    /**
     * Contains a comment about the purpose or content of an object.
     * https://msdn.microsoft.com/en-us/library/office/cc842022.aspx
     */
    public get comment(): string {
        return this.getStringItem(OutlookProperties.PR_COMMENT);
    }

    /**
     * Contains the creation date and time of a message.
     * https://msdn.microsoft.com/en-us/library/office/cc765677.aspx
     */
    public get creationTime(): Date | null {
        return this.getDateItem(OutlookProperties.PR_CREATION_TIME);
    }

    /**
     * Contains the date and time when the object or subobject was last modified.
     * https://msdn.microsoft.com/en-us/library/office/cc815689.aspx
     */
    public get modificationTime(): Date | null {
        return this.getDateItem(OutlookProperties.PR_LAST_MODIFICATION_TIME);
    }
}

export const enum AcknowledgementMode {
    manual = 0,
    automatic = 1,
}

export const enum Importance {
    low = 0,
    normal = 1,
    high = 2,
}
