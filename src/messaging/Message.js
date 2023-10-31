import { PropertyContext } from "../ltp/PropertyContext.js";
import { Utf8ArrayToStr } from "../utf8.js";
import { propertiesToObject } from "../util/propertiesToObject.js";
import { arrayBufferFromDataView } from "../util/arrayBufferFromDataView.js";
import * as Tags from "../ltp/Tags.js";
import { spGetSubject } from "../util/spGetSubject.js";
import { MF_HAS_ATTACH } from "./MessageFlags.js";
import { TableContext } from "../ltp/TableContext.js";
import {writeFileSync} from "fs";
import {join} from "path";

export class Message {
    #pstContext;
    #nid;
    #pc;
    /** @type {TableContext?} */
    #recipients;
    /** @type {TableContext?} */
    #attachments;

    get nid () { return this.#nid; }
    // get nidParent () { return this.#node.nidParent; }

    get subject () {
        return spGetSubject(/** @type {string|undefined} */(this.#pc.getValueByKey(Tags.PID_TAG_SUBJECT))||"");
    }
    get body () { return /** @type {string|undefined} */(this.#pc.getValueByKey(Tags.PID_TAG_BODY)); }
    get bodyHTML () {
        const dv = this.#pc.getValueByKey(Tags.PID_TAG_BODY_HTML);
        if (dv instanceof DataView) {
            const buffer = arrayBufferFromDataView(dv);
            return Utf8ArrayToStr(buffer);
        }
    }

    get hasAttachments () {
        const flags = /** @type {number|undefined} */(this.#pc.getValueByKey(Tags.PID_TAG_MESSAGE_FLAGS));
        if (typeof flags === "undefined") return;
        return Boolean(flags & MF_HAS_ATTACH);
    }

    /**
     * @param {import("../file/PSTInternal").PSTContext} pstContext
     * @param {number} nid
     * @param {PropertyContext} pc
     * @param {import('../ltp/TableContext').TableContext?} recipients
     * @param {import('../ltp/TableContext').TableContext?} attachments
     */
    constructor (pstContext, nid, pc, recipients, attachments) {
        this.#pstContext = pstContext;
        this.#nid = nid;
        this.#pc = pc;
        this.#recipients = recipients;
        this.#attachments = attachments;
    }

    getAllProperties () {
        return propertiesToObject(this.#pc.getAllProperties());
    }

    getAllPropertiesWithHeaders() {
        const keys = /** @type {number[]} */(this.#pc.keys);
        keys.splice(0,0, parseInt("0x007D", 16))
        const res = keys.map(key => ({
            tag: key,
            tagHex: "0x" + key.toString(16).padStart(4, "0"),
            tagName: this.#pc.getTagName(key),
            value: (() => {
                try {
                    return this.#pc.getValueByKey(key)
                } catch (e) {
                    return null
                }
            })()
        }));
        return propertiesToObject(res);
    }

    getAllRecipients () {
        const recipients = [];

        if (this.#recipients) {
            for (let i = 0; i < this.#recipients.recordCount; i++) {
                recipients.push(propertiesToObject(this.#recipients.getAllRowProperties(i)));
            }
        }

        return recipients;
    }

    getAttachmentEntries () {
        const attachments = [];

        if (this.#attachments) {
            for (let i = 0; i < this.#attachments.recordCount; i++) {
                attachments.push(propertiesToObject(this.#attachments.getAllRowProperties(i)));
            }
        }

        return attachments;
    }

    /**
     * @param {number} index
     */
    getAttachment (index) {
        if (index < 0) return null;

        if (this.#attachments) {
            const attachmentPCNid = this.#attachments.getCellValueByColumnTag(index, Tags.PID_TAG_LTP_ROW_ID);

            if (typeof attachmentPCNid === "number") {
                const pc = this.#pstContext.getSubPropertyContext(attachmentPCNid);

                if (!pc) {
                    throw Error("Unable to find Attachment PropertyContext");
                }

                return propertiesToObject(pc.getAllProperties());
            }
        }

    }

    async getRFC() {
        const info = this.getAllPropertiesWithHeaders()

        if (info.messageClass === "IPM.Note") {
            const recipients = this.getAllRecipients()
            const to = recipients.filter(addr => addr.recipientType === 1).map(addr => getRFCAddress(addr.displayName, addr.emailAddress))
            const cc = recipients.filter(addr => addr.recipientType === 2).map(addr => getRFCAddress(addr.displayName, addr.emailAddress))
            const bcc = recipients.filter(addr => addr.recipientType === 3).map(addr => getRFCAddress(addr.displayName, addr.emailAddress))
            const attachments = []
            if (this.hasAttachments) {
                try {
                    const attEntries = this.getAttachmentEntries()
                    const attObjects = attEntries.map((att, i) => {
                        try {
                            return this.getAttachment(i)
                        } catch (e) {
                            // console.log(e)
                            console.log('ERROR')
                            return null
                        }
                    }).filter(val => !!val)

                    for (const obj of attObjects) {
                        try {
                            const att = {
                                filename: obj.attachFilename || obj.displayName || null,
                                cid: obj.attachContentId || undefined,
                                content: Buffer.from(obj.attachDataBinary),
                                contentType: obj.attachMimeTag || undefined,
                            }
                            if (!att.cid)
                                delete att.cid
                            attachments.push(att)
                        } catch (e) {
                            console.log(e)
                        }
                    }
                } catch (e) {
                    console.log(e)
                }
            }


            const messageObject = {
                messageId: info.internetMessageId,
                from: getRFCAddress(info.senderName, info.senderEmailAddress),
                to,
                cc,
                bcc,
                subject: info.subject,
                text: info.body || null,
                html: info.html ? arrayBufferToString(info.html) : null,
                headers: headersToArr(info.transportMessageHeaders),
                date: info.clientSubmitTime || info.messageDeliveryTime || null,
                attachments
            }

            const mail = await buildMail(messageObject)
            return mail
        }
        return  null
    }
}

function getRFCAddress(display, address) {
    if (!display && address)
        return address.toLowerCase().trim()
    else if (display && !address)
        return display.trim()
    else if (!address && !display)
        return null
    else if (display.toLowerCase().trim() === address.toLowerCase().trim())
        return address.toLowerCase().trim()
    else
        return `${display} <${address.toLowerCase().trim()}>`
}
function arrayBufferToString(buffer) {
    let decoder = new TextDecoder("utf-8");
    return decoder.decode(buffer);
}

function headersToArr(headersStr) {
    const lines = (headersStr || '').split('\n')
    let actualLines = []
    for (const line of lines) {
        if (line && line.length) {
            if (!actualLines.length)
                actualLines.push(line)
            else if (line.charAt(0) === ' ')
                actualLines.splice(-1, 1, actualLines.at(-1) + '\n' + line)
            else
                actualLines.push(line)

        }
    }
    const headers = []
    for (const line of actualLines) {
        const key = line.match(/^[^:]+/i)?.shift()
        if (key && key.length < 100 && key.toLowerCase() !== 'content-type') {
            const value = line.replace(key + ":", '').trim()
            if (value)
                headers.push({key, value})
        }
    }



    return headers
}

function buildMail(mailObj) {
    return new Promise((resolve, reject) => {
        try {
            import('nodemailer/lib/mail-composer/index.js')
                .then(MailComposerModule => {
                    const MailComposer = MailComposerModule.default;
                    const mail = new MailComposer(mailObj).compile()
                    mail.keepBcc = true
                    mail.build(function(err, message){
                        if (err)
                            reject(err)
                        else
                            resolve(message)
                    });
                });
        } catch (e) {
            reject(e)
        }
    })
}
