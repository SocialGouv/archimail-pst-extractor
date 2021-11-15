import { OutlookProperties } from "./OutlookProperties";
import { PSTMessage } from "./PSTMessage";

export class PSTActivity extends PSTMessage {
    /**
     * Contains the display name of the journaling application (for example, "MSWord"), and is typically a free-form attribute of a journal message, usually a string.
     * https://msdn.microsoft.com/en-us/library/office/cc839662.aspx
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
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                importance: this.importance,
                isDocumentPosted: this.isDocumentPosted,
                isDocumentPrinted: this.isDocumentPrinted,
                isDocumentRouted: this.isDocumentRouted,
                isDocumentSaved: this.isDocumentSaved,
                logDuration: this.logDuration,
                logEnd: this.logEnd,
                logFlags: this.logFlags,
                logStart: this.logStart,
                logType: this.logType,
                logTypeDesc: this.logTypeDesc,
                messageClass: this.messageClass,
                subject: this.subject,
                transportMessageHeaders: this.transportMessageHeaders,
            },
            this
        );
        return clone;
    }
}
