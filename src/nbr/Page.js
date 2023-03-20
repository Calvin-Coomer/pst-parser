import { BTEntry } from "./BTEntry.js";
import { BlockEntry } from "./BlockEntry.js";
import { NodeEntry } from "./NodeEntry.js";

export class Page {
    #dv;

    get ptype () { return this.#dv.getUint8(496); }
    get ptypeRepeat () { return this.#dv.getUint8(497); }
    get wSig () { return this.#dv.getUint16(498, true); }
    get dwCRC () { return this.#dv.getUint32(500, true); }
    get bid () { return this.#dv.getBigUint64(504, true); }

    /**
     * @param {DataView} dv
     */
    constructor (dv) {
        this.#dv = dv;
    }

    /**
     * @param {DataView} dv
     */
    static getPage (dv) {
        const p = new Page(dv);

        if (p.ptype === 0x80) return new BTPage(dv);
        if (p.ptype === 0x81) return new BTPage(dv);
        if (p.ptype === 0x82) return new FMapPage(dv);
        if (p.ptype === 0x83) return new PMapPage(dv);
        if (p.ptype === 0x84) return new AMapPage(dv);
        if (p.ptype === 0x85) return new FPMapPage(dv);
        if (p.ptype === 0x86) return new DListPage(dv);

        return p;
    }
}

export class BTPage extends Page {
    #dv;

    /**
     * @param {DataView} dv
     */
    constructor (dv) {
        super(dv);

        this.#dv = dv;
    }

    get cEnt () { return this.#dv.getUint8(488); }
    get cEntMax () { return this.#dv.getUint8(489); }
    get cbEnt () { return this.#dv.getUint8(490); }
    get cLevel () { return this.#dv.getUint8(491); }
    get dwPadding () { return this.#dv.getUint32(492, true); }

    /**
     * @param {number} n
     */
    getEntry (n) {
        const begin = n * this.cbEnt;

        if (this.cLevel === 0) {
            if (this.ptype === 0x80) {
                return new BlockEntry(new DataView(this.#dv.buffer, this.#dv.byteOffset + begin));
            }
            if (this.ptype === 0x81) {
                return new NodeEntry(new DataView(this.#dv.buffer, this.#dv.byteOffset + begin));
            }
            throw Error("Invalid BT Page");
        }

        return new BTEntry(new DataView(this.#dv.buffer, this.#dv.byteOffset + begin));
    }

    get keys () {
        const out = [];
        for (let i = 0; i < this.cEnt; i++) {
            const entry = this.getEntry(i);

            if (entry instanceof BlockEntry) {
                out.push(entry.BREF.bid);
            }
            else if (entry instanceof NodeEntry) {
                out.push(entry.nid);
            }
            else {
                out.push(entry.btkey);
            }
        }
        return out;
    }

    /**
     * @param {number | bigint} key
     * @return {(BlockEntry|NodeEntry)?}
     */
    findEntry (key) {
        const keys = this.keys;

        // If we're level 0 then either *we* have the key or it doesn't exist.

        if (this.cLevel === 0) {
            const k = typeof key === "bigint" ? key : BigInt(key);
            const index = keys.indexOf(k);

            if (index >= 0) {
                return /** @type {BlockEntry|NodeEntry} */(this.getEntry(index));
            }

            // Not in leaf
            return null;
        }

        // If we're above level 0 then find which of our child pages should
        // be the next step.

        for (let i = 0; i < this.cEnt; i++) {
            if (keys[i] > key) {
                // Before the first entry
                if (i === 0) return null;

                const entry = this.getEntry(i - 1);
                if (!(entry instanceof BTEntry)) {
                    throw Error("Expected BT Entry");
                }

                const page = Page.getPage(new DataView(this.#dv.buffer, parseInt(entry.BREF.ib.toString())));
                if (!(page instanceof BTPage)) {
                    throw Error("Expected BT Page");
                }

                return page.findEntry(key);
            }
        }

        // It must be in the our page (or doesn't exist).

        const entry = this.getEntry(this.cEnt - 1);
        if (!(entry instanceof BTEntry)) {
            throw Error("Expected BT Entry");
        }

        const page = Page.getPage(new DataView(this.#dv.buffer, parseInt(entry.BREF.ib.toString())));
        if (!(page instanceof BTPage)) {
            throw Error("Expected BT Page");
        }

        return page.findEntry(key);
    }

    /**
     *
     * @returns {bigint[]}
     */
    getAllKeys () {
        if (this.cLevel === 0) {
            return this.keys;
        }

        const out = [];

        for (let i = 0; i < this.cEnt; i++) {
            const entry = this.getEntry(i);

            if (!(entry instanceof BTEntry)) {
                throw Error("Expected BTEntry. Got: " + (
                    entry instanceof NodeEntry ? "Node" : (
                        entry instanceof BTEntry ? ("BTEntry. bid: " + entry.BREF.bid) : "Other"
                    )
                ));
            }

            const page = Page.getPage(new DataView(this.#dv.buffer, parseInt(entry.BREF.ib.toString())));
            if (!(page instanceof BTPage)) {
                throw Error("Expected BT Page");
            }

            out.push(...page.getAllKeys());
        }

        return out;
    }
}

export class AMapPage extends Page {

}

export class PMapPage extends Page {

}

export class DListPage extends Page {

}

export class FMapPage extends Page {

}

export class FPMapPage extends Page {

}