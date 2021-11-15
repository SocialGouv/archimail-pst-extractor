import { OutlookProperties } from "./OutlookProperties";
import { PSTMessage } from "./PSTMessage";

/**
 * PSTAppointment is for Calendar items.
 */
export class PSTAppointment extends PSTMessage {
    /**
     * Specifies if a meeting request should be sent as an iCal message.
     * https://msdn.microsoft.com/en-us/library/office/cc839802.aspx
     */
    public get sendAsICAL(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidSendMeetingAsIcal,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Represents the user’s availability for an appointment.
     * https://msdn.microsoft.com/en-us/library/office/cc841972.aspx
     */
    public get busyStatus(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidBusyStatus,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * The user is busy.
     * https://msdn.microsoft.com/en-us/library/office/cc841972.aspx
     */
    public get showAsBusy(): boolean {
        return this.busyStatus === 2;
    }

    /**
     * Represents the location of an appointment.
     * https://msdn.microsoft.com/en-us/library/office/cc842419.aspx
     */
    public get location(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidLocation,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Represents the date and time when an appointment begins.
     * https://msdn.microsoft.com/en-us/library/office/cc839929.aspx
     */
    public get startTime(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentStartWhole,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Represents the date and time that an appointment ends.
     * https://msdn.microsoft.com/en-us/library/office/cc815864.aspx
     */
    public get endTime(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentEndWhole,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Represents the length of time, in minutes, when an appointment is scheduled.
     * https://msdn.microsoft.com/en-us/library/office/cc842287.aspx
     */
    public get duration(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentDuration,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the color to use when displaying the calendar.
     * https://msdn.microsoft.com/en-us/library/office/cc842274.aspx
     */
    public get color(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentColor,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies whether or not the event is all day.
     * https://msdn.microsoft.com/en-us/library/office/cc839901.aspx
     */
    public get subType(): boolean {
        return (
            this.getIntItem(
                this.pstFile.getNameToIdMapItem(
                    OutlookProperties.PidLidAppointmentSubType,
                    OutlookProperties.PSETID_Appointment
                )
            ) !== 0
        );
    }

    /**
     * Specifies a bit field that describes the state of the object.
     * https://msdn.microsoft.com/en-us/library/office/cc765762.aspx
     */
    public get meetingStatus(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentStateFlags,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the response status of an attendee.
     * https://msdn.microsoft.com/en-us/library/office/cc839923.aspx
     */
    public get responseStatus(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidResponseStatus,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies whether an appointment message is recurrent.
     * https://msdn.microsoft.com/en-us/library/office/cc765772.aspx
     */
    public get isRecurring(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidRecurring,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the date and time within the recurrence pattern that the exception will replace.
     * https://msdn.microsoft.com/en-us/library/office/cc842450.aspx
     */
    public get recurrenceBase(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidExceptionReplaceTime,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the recurrence type of the recurring series.
     * https://msdn.microsoft.com/en-us/library/office/cc842135.aspx
     */
    public get recurrenceType(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidRecurrenceType,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies a description of the recurrence pattern of the calendar object.
     * https://msdn.microsoft.com/en-us/library/office/cc815733.aspx
     */
    public get recurrencePattern(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidRecurrencePattern,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the dates and times when a recurring series occurs by using one of the recurrence patterns and ranges that are specified in [MS-OXOCAL].
     * https://msdn.microsoft.com/en-us/library/office/cc842017.aspx
     */
    public get recurrenceStructure(): Buffer | null {
        return this.getBinaryItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentRecur,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Contains a stream that maps to the persisted format of a TZREG structure, which describes the time zone to be used for the start and end time of a recurring appointment or meeting request.
     * https://msdn.microsoft.com/en-us/library/office/cc815376.aspx
     */
    public get timezone(): Buffer | null {
        return this.getBinaryItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidTimeZoneStruct,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies a list of all the attendees except for the organizer, including resources and unsendable attendees.
     * https://msdn.microsoft.com/en-us/library/office/cc815418.aspx
     */
    public get allAttendees(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAllAttendeesString,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Contains a list of all the sendable attendees who are also required attendees.
     * https://msdn.microsoft.com/en-us/library/office/cc842502.aspx
     */
    public get toAttendees(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidToAttendeesString,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Contains a list of all the sendable attendees who are also optional attendees.
     * https://msdn.microsoft.com/en-us/library/office/cc839636.aspx
     */
    public get ccAttendees(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidCcAttendeesString,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the sequence number of a Meeting object.
     * https://msdn.microsoft.com/en-us/library/office/cc765937.aspx
     */
    public get appointmentSequence(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentSequence,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Is a hosted meeting?
     * https://msdn.microsoft.com/en-us/library/ee200872(v=exchg.80).aspx
     */
    public get isOnlineMeeting(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidConferencingCheck,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the type of the meeting.
     * https://msdn.microsoft.com/en-us/library/ee158396(v=exchg.80).aspx
     */
    public get netMeetingType(): number {
        return this.getIntItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidConferencingType,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the directory server to be used.
     * https://msdn.microsoft.com/en-us/library/ee201516(v=exchg.80).aspx
     */
    public get netMeetingServer(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidDirectory,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the email address of the organizer.
     * https://msdn.microsoft.com/en-us/library/ee203317(v=exchg.80).aspx
     */
    public get netMeetingOrganizerAlias(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidOrganizerAlias,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the document to be launched when the user joins the meeting.
     * https://msdn.microsoft.com/en-us/library/ee204395(v=exchg.80).aspx
     */
    public get netMeetingDocumentPathName(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidCollaborateDoc,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * The PidLidNetShowUrl property ([MS-OXPROPS] section 2.175) specifies the URL to be launched when the user joins the meeting
     * https://msdn.microsoft.com/en-us/library/ee179451(v=exchg.80).aspx
     */
    public get netShowURL(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidNetShowUrl,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Specifies the date and time at which the meeting-related object was sent.
     * https://msdn.microsoft.com/en-us/library/ee237112(v=exchg.80).aspx
     */
    public get attendeeCriticalChange(): Date | null {
        return this.getDateItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAttendeeCriticalChange,
                OutlookProperties.PSETID_Meeting
            )
        );
    }

    /**
     * Indicates that this meeting response is a counter proposal.
     * https://msdn.microsoft.com/en-us/magazine/cc815846.aspx
     */
    public get appointmentCounterProposal(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidAppointmentCounterProposal,
                OutlookProperties.PSETID_Appointment
            )
        );
    }

    /**
     * Indicates whether the user did not include any text in the body of the Meeting Response object.
     * https://msdn.microsoft.com/en-us/library/ee159822(v=exchg.80).aspx
     */
    public get isSilent(): boolean {
        return this.getBooleanItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidIsSilent,
                OutlookProperties.PSETID_Meeting
            )
        );
    }

    /**
     * Identifies required attendees for the appointment or meeting.
     * https://msdn.microsoft.com/en-us/library/ee160700(v=exchg.80).aspx
     */
    public get requiredAttendees(): string {
        return this.getStringItem(
            this.pstFile.getNameToIdMapItem(
                OutlookProperties.PidLidRequiredAttendees,
                OutlookProperties.PSETID_Meeting
            )
        );
    }

    /**
     * Contains the Windows Locale ID of the end-user who created this message.
     * https://msdn.microsoft.com/en-us/library/ee201602(v=exchg.80).aspx
     */
    public get localeId(): number {
        return this.getIntItem(OutlookProperties.PidTagMessageLocaleId);
    }

    /**
     * JSON stringify the object properties.  Large fields (like body) aren't included.
     */
    public toJSON(): unknown {
        const clone = Object.assign(
            {
                allAttendees: this.allAttendees,
                appointmentCounterProposal: this.appointmentCounterProposal,
                appointmentSequence: this.appointmentSequence,
                attendeeCriticalChange: this.attendeeCriticalChange,
                busyStatus: this.busyStatus,
                ccAttendees: this.ccAttendees,
                color: this.color,
                duration: this.duration,
                endTime: this.endTime,
                importance: this.importance,
                isOnlineMeeting: this.isOnlineMeeting,
                isRecurring: this.isRecurring,
                isSilent: this.isSilent,
                localeId: this.localeId,
                location: this.location,
                meetingStatus: this.meetingStatus,
                messageClass: this.messageClass,
                netMeetingDocumentPathName: this.netMeetingDocumentPathName,
                netMeetingOrganizerAlias: this.netMeetingOrganizerAlias,
                netMeetingServer: this.netMeetingServer,
                netMeetingType: this.netMeetingType,
                recurrenceBase: this.recurrenceBase,
                recurrencePattern: this.recurrencePattern,
                recurrenceStructure: this.recurrenceStructure,
                recurrenceType: this.recurrenceType,
                requiredAttendees: this.requiredAttendees,
                sendAsICAL: this.sendAsICAL,
                showAsBusy: this.showAsBusy,
                startTime: this.startTime,
                subType: this.subType,
                subject: this.subject,
                timezone: this.timezone,
                toAttendees: this.toAttendees,
                transportMessageHeaders: this.transportMessageHeaders,
            },
            this
        );
        return clone;
    }
}
