import { OutlookProperties } from "./OutlookProperties";
import type { PSTFile } from "./PSTFile";
import { PSTObject } from "./PSTObject";
import type { PSTTableItem } from "./PSTTableItem";

// Class containing recipient information
export class PSTRecipient extends PSTObject {
    /**
     * Creates an instance of PSTRecipient.
     */
    constructor(pstFile: PSTFile, recipientDetails: Map<number, PSTTableItem>) {
        super(pstFile, undefined, recipientDetails);
    }

    /**
     * Contains the recipient type for a message recipient.
     * https://msdn.microsoft.com/en-us/library/office/cc839620.aspx
     */
    public get recipientType(): number {
        return this.getIntItem(OutlookProperties.PR_RECIPIENT_TYPE);
    }

    /**
     * Contains the messaging user's e-mail address type, such as SMTP.
     * https://msdn.microsoft.com/en-us/library/office/cc815548.aspx
     */
    public get addrType(): string {
        return this.getStringItem(OutlookProperties.PR_ADDRTYPE);
    }

    /**
     * Contains the messaging user's e-mail address.
     * https://msdn.microsoft.com/en-us/library/office/cc842372.aspx
     */
    public get emailAddress(): string {
        return this.getStringItem(OutlookProperties.PR_EMAIL_ADDRESS);
    }

    /**
     * Specifies a bit field that describes the recipient status.
     * https://msdn.microsoft.com/en-us/library/office/cc815629.aspx
     */
    public get recipientFlags(): number {
        return this.getIntItem(OutlookProperties.PR_RECIPIENT_FLAGS);
    }

    /**
     * Specifies the location of the current recipient in the recipient table.
     * https://msdn.microsoft.com/en-us/library/ee201359(v=exchg.80).aspx
     */
    public get recipientOrder(): number {
        return this.getIntItem(OutlookProperties.PidTagRecipientOrder);
    }

    /**
     * Contains recipient display name (for recipient, not message).
     * https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/7a299df8-2a9e-4a5c-8ebf-0afb8630ff54
     */
    public get recipientDisplayName(): string {
        return this.getStringItem(OutlookProperties.PidTagRecipientDisplayName);
    }

    /**
     * Contains the SMTP address for the address book object.
     * https://msdn.microsoft.com/en-us/library/office/cc842421.aspx
     */
    public get smtpAddress(): string {
        // If the recipient address type is SMTP, we can simply return the recipient address.
        const addressType = this.addrType;
        if (addressType.toLowerCase() === "smtp") {
            const addr = this.emailAddress;
            if (addr.length) {
                return addr;
            }
        }
        // Otherwise, we have to hope the SMTP address is present as the PidTagPrimarySmtpAddress property.
        return this.getStringItem(OutlookProperties.PR_SMTP_ADDRESS);
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                addrType: this.addrType,
                emailAddress: this.emailAddress,
                recipientFlags: this.recipientFlags,
                recipientOrder: this.recipientOrder,
                recipientType: this.recipientType,
                smtpAddress: this.smtpAddress,
            },
            this
        );
        return clone;
    }
}

export const enum RecipientType {
    originator = 0x00000000,
    primary = 0x00000001,
    cc = 0x00000002,
    bcc = 0x00000003,
}

export const enum RecipientTags {
    sendable = 0x00000001,
}
