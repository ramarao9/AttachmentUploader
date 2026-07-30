import { IInputs } from "./generated/ManifestTypes";

const BASE64_MARKER = ";base64,";

export interface FileInfo {
    name: string;
    type: string;
    size: number;
    /** File content as base64, without the `data:<mime>;base64,` prefix. */
    content: string;
}

export interface UploadTarget {
    id: string;
    entityName: string;
    entitySetName: string;
    useNoteAttachment: boolean;
    defaultNoteTitle: string | null;
}

export interface UploadResult {
    fileName: string;
    succeeded: boolean;
    error?: string;
}

/** Reads a browser File into base64. */
export async function readFile(file: File): Promise<FileInfo> {
    const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onabort = () => reject(new Error(`Reading ${file.name} was aborted.`));
        reader.onerror = () => reject(reader.error ?? new Error(`Could not read ${file.name}.`));
        reader.readAsDataURL(file);
    });

    return {
        name: file.name,
        type: file.type,
        size: file.size,
        content: stripDataUriPrefix(dataUrl)
    };
}

export function stripDataUriPrefix(dataUrl: string): string {
    const markerIndex = dataUrl.indexOf(BASE64_MARKER);
    return markerIndex < 0 ? "" : dataUrl.substring(markerIndex + BASE64_MARKER.length);
}

/** Splits the AcceptedFileTypes property into individual `.ext` / `mime/type` / `mime/*` tokens. */
export function parseAcceptTokens(acceptedFileTypes: string | null): string[] {
    return (acceptedFileTypes ?? "")
        .split(",")
        .map((token) => token.trim().toLowerCase())
        .filter((token) => token !== "");
}

/** Mirrors the semantics of the html `accept` attribute. An empty token list accepts everything. */
export function matchesAccept(file: { name: string; type: string }, tokens: string[]): boolean {
    if (tokens.length === 0) {
        return true;
    }

    const name = file.name.toLowerCase();
    const type = (file.type ?? "").toLowerCase();

    return tokens.some((token) => {
        if (token.startsWith(".")) {
            return name.endsWith(token);
        }
        if (token.endsWith("/*")) {
            return type.startsWith(token.slice(0, -1));
        }
        return type === token;
    });
}

/**
 * Attachments on emails and appointments go to activitymimeattachment unless the maker opted into
 * notes; everything else becomes an annotation.
 */
export function isActivityMimeAttachment(target: UploadTarget): boolean {
    const entity = target.entityName.toLowerCase();
    return !target.useNoteAttachment && (entity === "email" || entity === "appointment");
}

export function buildAttachmentRecord(target: UploadTarget, fileInfo: FileInfo): ComponentFramework.WebApi.Entity {
    const attachmentRecord: ComponentFramework.WebApi.Entity = {};

    if (isActivityMimeAttachment(target)) {
        attachmentRecord["objectid_activitypointer@odata.bind"] = `/activitypointers(${target.id})`;
        attachmentRecord.body = fileInfo.content;
    } else {
        attachmentRecord[`objectid_${target.entityName}@odata.bind`] = `/${target.entitySetName}(${target.id})`;
        attachmentRecord.documentbody = fileInfo.content;

        if (target.defaultNoteTitle != null && target.defaultNoteTitle !== "") {
            attachmentRecord.subject = target.defaultNoteTitle;
        }
    }

    if (fileInfo.type && fileInfo.type !== "") {
        attachmentRecord.mimetype = fileInfo.type;
    }

    attachmentRecord.filename = fileInfo.name;
    attachmentRecord.objecttypecode = target.entityName;

    return attachmentRecord;
}

/**
 * Uploads the files one after another. Each file is wrapped in its own try/catch so a single
 * failure no longer abandons the rest of the batch; the caller decides what to show.
 */
export async function uploadToDataverse(
    context: ComponentFramework.Context<IInputs>,
    target: UploadTarget,
    files: FileInfo[],
    onProgress?: (completed: number, total: number) => void
): Promise<UploadResult[]> {
    const results: UploadResult[] = [];
    const attachmentEntity = isActivityMimeAttachment(target) ? "activitymimeattachment" : "annotation";

    for (let i = 0; i < files.length; i++) {
        const fileInfo = files[i];
        try {
            await context.webAPI.createRecord(attachmentEntity, buildAttachmentRecord(target, fileInfo));
            results.push({ fileName: fileInfo.name, succeeded: true });
        } catch (e) {
            results.push({ fileName: fileInfo.name, succeeded: false, error: toErrorMessage(e) });
        }
        onProgress?.(i + 1, files.length);
    }

    return results;
}

export function toErrorMessage(error: unknown): string {
    if (error instanceof Error) {
        return error.message;
    }
    const message = (error as { message?: string })?.message;
    return message ?? String(error);
}
