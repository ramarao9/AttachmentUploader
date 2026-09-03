const BASE64_MARKER = ";base64,";

export function stripDataUriPrefix(dataUrl: string): string {
    const markerIndex = dataUrl.indexOf(BASE64_MARKER);
    return markerIndex < 0 ? "" : dataUrl.substring(markerIndex + BASE64_MARKER.length);
}

/**
 * Base64 encodes a blob. Lives in its own module because both the single request upload and the
 * block upload need it, and having either import the other would be a cycle.
 *
 * FileReader rather than btoa(): btoa needs the bytes as a binary string, which means building one
 * from a Uint8Array, and String.fromCharCode(...bytes) blows the argument limit long before the
 * 4 MB blocks this is used for.
 *
 * `label` is only used to name the blob in error messages.
 */
export async function readBlobAsBase64(blob: Blob, label: string): Promise<string> {
    const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onabort = () => reject(new Error(`Reading ${label} was aborted.`));
        reader.onerror = () => reject(reader.error ?? new Error(`Could not read ${label}.`));
        reader.readAsDataURL(blob);
    });

    return stripDataUriPrefix(dataUrl);
}

/**
 * Uuid for block ids and client generated annotation ids.
 *
 * crypto.randomUUID needs a secure context. Every real host is https and localhost counts as
 * secure, so the fallback only exists so a plain http harness does not fail outright; these values
 * have to be unique, not unguessable.
 */
export function newUuid(): string {
    const cryptoApi: Crypto | undefined = window.crypto;
    if (typeof cryptoApi?.randomUUID === "function") {
        return cryptoApi.randomUUID();
    }

    const bytes = new Uint8Array(16);
    if (typeof cryptoApi?.getRandomValues === "function") {
        cryptoApi.getRandomValues(bytes);
    } else {
        for (let i = 0; i < bytes.length; i++) {
            bytes[i] = Math.floor(Math.random() * 256);
        }
    }
    // Version 4, variant 1.
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;

    const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}
