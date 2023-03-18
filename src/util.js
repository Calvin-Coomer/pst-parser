
/**
 * @param {number|bigint} n
 */
export function h (n) {
    const s = n.toString(16).toUpperCase();
    return s.padStart(s.length + (s.length % 2), "0");
}

/**
 * @typedef {{tag: number;tagHex: string;tagName: string;value: string | number | bigint | boolean | string[] | DataView | number[] | Date | undefined;}} PropertyData
 */

/**
 * @param {PropertyData[]} properties
 */
export function propertiesToObject (properties) {
    const out = {};
    for (const prop of properties) {
        const value = prop.value instanceof DataView ?
            arrayBufferFromDataView(prop.value) :
            prop.value;
        out[prop.tagName||prop.tagHex] = value;
    }
    return out;
}

/**
 * UTF-16 strings from a buffer
 * @param {ArrayBuffer} buffer
 * @param {number} byteOffset
 * @param {number} byteLength
 */
export function stringFromBuffer (buffer, byteOffset, byteLength) {
    if (byteOffset % 2) {
        // Uint16Array *must* start on a multiple of 2
        // We have no choice but to copy the buffer
        const stringBuffer = buffer.slice(byteOffset, byteOffset + byteLength);
        return String.fromCharCode(...new Uint16Array(stringBuffer));
    }

    return String.fromCharCode(...new Uint16Array(buffer, byteOffset, byteLength/2));
}

/**
 * Produce an ArrayBuffer exactly matching the view defined by the DataView
 * @param {DataView} dataView
 */
export function arrayBufferFromDataView (dataView) {
    if (dataView.byteOffset === 0
        && dataView.byteLength === dataView.buffer.byteLength)
    {
        // We don't need to slice anything. Just return the
        return dataView.buffer;
    }

    return dataView.buffer.slice(dataView.byteOffset, dataView.byteOffset + dataView.byteLength);
}