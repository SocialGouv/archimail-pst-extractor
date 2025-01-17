/* eslint-disable @typescript-eslint/naming-convention */
const OFFSETS = {
    EndDate: -4,
    EndType: 22,
    FirstDOW: 30,
    FirstDateTime: 10,
    OccurrenceCount: 26,
    PatternType: 6,
    PatternTypeSpecific: 22,
    Period: 14,
    RecurFrequency: 4,
    StartDate: -8,
};

export enum RecurFrequency {
    Daily = 0x200a,
    Weekly = 0x200b,
    Monthly = 0x200c,
    Yearly = 0x200d,
}

export enum PatternType {
    Day = 0x0000,
    Week = 0x0001,
    Month = 0x0002,
    MonthEnd = 0x0004,
    MonthNth = 0x0003,
}

export enum EndType {
    AfterDate = 0x00002021,
    AfterNOccurrences = 0x00002022,
    NeverEnd = 0x00002023,
}

export enum NthOccurrence {
    First = 0x0001,
    Second = 0x0002,
    Third = 0x0003,
    Fourth = 0x0004,
    Last = 0x0005,
}

export type WeekSpecific = boolean[] & { length: 7 };
export interface MonthNthSpecific {
    weekdays: WeekSpecific;
    nth: NthOccurrence;
}

export class RecurrencePattern {
    public recurFrequency: RecurFrequency;

    public patternType: PatternType;

    public firstDateTime: Date;

    public period: number;

    public patternTypeSpecific;

    public endType: EndType;

    public occurrenceCount: number;

    public firstDOW: number;

    public startDate: Date;

    public endDate: Date;

    constructor(private readonly buffer: Buffer) {
        const bufferEnd = buffer.length;
        let patternTypeOffset = 0;

        this.recurFrequency = this.readInt(OFFSETS.RecurFrequency, 1);
        this.patternType = this.readInt(OFFSETS.PatternType, 1);
        this.firstDateTime = winToJsDate(
            this.readInt(OFFSETS.FirstDateTime, 2)
        );
        this.period = this.readInt(OFFSETS.Period, 2);
        this.patternTypeSpecific = this.readPatternTypeSpecific(
            this.patternType
        );

        switch (this.patternType) {
            case PatternType.Week:
            case PatternType.Month:
            case PatternType.MonthEnd:
                patternTypeOffset = 4;
                break;
            case PatternType.MonthNth:
                patternTypeOffset = 8;
                break;
            default:
                break;
        }

        this.endType = this.readInt(OFFSETS.EndType + patternTypeOffset, 2);
        this.occurrenceCount = this.readInt(
            OFFSETS.OccurrenceCount + patternTypeOffset,
            2
        );
        this.firstDOW = this.readInt(OFFSETS.FirstDOW + patternTypeOffset, 2);
        this.startDate = winToJsDate(
            this.readInt(bufferEnd + OFFSETS.StartDate, 2)
        );
        this.endDate = winToJsDate(
            this.readInt(bufferEnd + OFFSETS.EndDate, 2)
        );
    }

    public toJSON(): unknown {
        return {
            endDate: this.endDate,
            endType: EndType[this.endType],
            firstDOW: this.firstDOW,
            firstDateTime: this.firstDateTime,
            occurrenceCount: this.occurrenceCount,
            patternType: PatternType[this.patternType],
            patternTypeSpecific: this.patternTypeSpecific,
            period: this.period,
            recurFrequency: RecurFrequency[this.recurFrequency],
            startDate: this.startDate,
        };
    }

    private readInt(offset: number, size: 1 | 2) {
        switch (size) {
            case 1:
                return this.buffer.readInt16LE(offset);
            case 2:
                return this.buffer.readInt32LE(offset);
        }
    }

    private readPatternTypeSpecific(
        type: PatternType
    ): MonthNthSpecific | WeekSpecific | number | null {
        switch (type) {
            case PatternType.Day:
                return null;
            case PatternType.Week:
                return readWeekByte(
                    this.buffer.readInt8(OFFSETS.PatternTypeSpecific)
                );
            case PatternType.Month:
            case PatternType.MonthEnd:
                return this.readInt(OFFSETS.PatternTypeSpecific, 2);
            case PatternType.MonthNth:
                return {
                    nth: this.readInt(OFFSETS.PatternTypeSpecific + 4, 2),
                    weekdays: readWeekByte(
                        this.buffer.readInt8(OFFSETS.PatternTypeSpecific)
                    ),
                };
        }
    }
}

function winToJsDate(dateInt: number): Date {
    return new Date(dateInt * 60 * 1000 - 1.16444736e13); // subtract milliseconds between 1601-01-01 and 1970-01-01
}

function readWeekByte(byte: number): WeekSpecific {
    const weekArr = [];
    for (let i = 0; i < 7; ++i) {
        weekArr.push(Boolean(byte & (1 << i)));
    }
    return weekArr as WeekSpecific;
}
