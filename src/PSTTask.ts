import { OutlookProperties } from "./OutlookProperties";
import { PSTFile } from "./PSTFile";
import { PSTMessage } from "./PSTMessage";
import { RecurrencePattern } from "./RecurrencePattern";

export class PSTTask extends PSTMessage {
    /**
     * Specifies the status of the user's progress on the task.
     * https://msdn.microsoft.com/en-us/library/office/cc842120.aspx
     */
    public get taskStatus(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskStatus,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates the progress the user has made on a task.
     * https://msdn.microsoft.com/en-us/library/office/cc839932.aspx
     */
    public get percentComplete(): number {
        return this.getDoubleItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidPercentComplete,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Specifies the date when the user completes the task.
     * https://msdn.microsoft.com/en-us/library/office/cc815753.aspx
     */
    public get taskDateCompleted(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskDateCompleted,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates the number of minutes that the user performed a task.
     * https://msdn.microsoft.com/en-us/library/office/cc842253.aspx
     */
    public get taskActualEffort(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskActualEffort,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates the amount of time, in minutes, that the user expects to perform a task.
     * https://msdn.microsoft.com/en-us/library/office/cc842485.aspx
     */
    public get taskEstimatedEffort(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskEstimatedEffort,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates which copy is the latest update of a task.
     * https://msdn.microsoft.com/en-us/library/office/cc815510.aspx
     */
    public get taskVersion(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskVersion,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates the task is complete.
     * https://msdn.microsoft.com/en-us/library/office/cc839514.aspx
     */
    public get isTaskComplete(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskComplete,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Contains the name of the task owner.
     * https://msdn.microsoft.com/en-us/library/office/cc842363.aspx
     */
    public get taskOwner(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskOwner,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Names the user who was last assigned the task.
     * https://msdn.microsoft.com/en-us/library/office/cc815865.aspx
     */
    public get taskAssigner(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskAssigner,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Names the most recent user who was the task owner.
     * https://msdn.microsoft.com/en-us/library/office/cc842278.aspx
     */
    public get taskLastUser(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskLastUser,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Provides an aid to custom sorting tasks.
     * https://msdn.microsoft.com/en-us/library/office/cc765654.aspx
     */
    public get taskOrdinal(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskOrdinal,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates whether the task includes a recurrence pattern.
     * https://msdn.microsoft.com/en-us/library/office/cc765875.aspx
     */
    public get isTaskRecurring(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskFRecurring,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidtaskrecurrence-canonical-property
     */
    public get taskRecurrencePattern(): RecurrencePattern | null {
        const recurrenceBLOB = this.getBinaryItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskRecurrence,
                PSTFile.PSETID_Task
            )
        );
        return recurrenceBLOB && new RecurrencePattern(recurrenceBLOB);
    }

    /**
     * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidtaskdeadoccurrence-canonical-property
     */
    public get taskDeadOccurrence(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskDeadOccurrence,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates the role of the current user relative to the task.
     * https://msdn.microsoft.com/en-us/library/office/cc842113.aspx
     */
    public get taskOwnership(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskOwnership,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * Indicates the acceptance state of the task.
     * https://msdn.microsoft.com/en-us/library/office/cc839689.aspx
     */
    public get acceptanceState(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTaskAcceptanceState,
                PSTFile.PSETID_Task
            )
        );
    }

    /**
     * JSON stringify the object properties.
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                acceptanceState: this.acceptanceState,
                importance: this.importance,
                isTaskComplete: this.isTaskComplete,
                isTaskRecurring: this.isTaskRecurring,
                messageClass: this.messageClass,
                percentComplete: this.percentComplete,
                subject: this.subject,
                taskActualEffort: this.taskActualEffort,
                taskAssigner: this.taskAssigner,
                taskDateCompleted: this.taskDateCompleted,
                taskEstimatedEffort: this.taskEstimatedEffort,
                taskLastUser: this.taskLastUser,
                taskOrdinal: this.taskOrdinal,
                taskOwner: this.taskOwner,
                taskOwnership: this.taskOwnership,
                taskStatus: this.taskStatus,
                taskVersion: this.taskVersion,
                transportMessageHeaders: this.transportMessageHeaders,
            },
            this
        );
        return clone;
    }
}
