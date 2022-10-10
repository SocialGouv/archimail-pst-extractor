export interface NameAndAddress {
    name?: string;
    email?: string;
}

/* eslint-disable @typescript-eslint/naming-convention */
export interface InternalEmlHeaders {
    Bcc: string;
    Cc: string;
    "Content-Type": string;
    From: string;
    Subject: string;
    To: string;
}
/* eslint-enable @typescript-eslint/naming-convention */

export interface EmlAttachment {
    contentType?: string;
    inline?: boolean;
    filename?: string;
    name?: string;
    cid?: string;
    data: Buffer | string;
}

export interface EmlStruct {
    headers?: Record<string, string[] | string>;
    subject?: string;
    from?: NameAndAddress | string;
    to: NameAndAddress | NameAndAddress[] | string;
    cc?: NameAndAddress | NameAndAddress[] | string;
    text?: string;
    html?: string;
    attachments?: EmlAttachment[];
}
