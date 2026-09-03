import { IInputs } from "./generated/ManifestTypes";
import { newUuid, readBlobAsBase64 } from "./base64";
import { BlockUploadEntity, CHUNK_THRESHOLD_BYTES, uploadLargeFile } from "./chunkedUpload";
import { getClientUrl } from "./hostContext";

/**
 * Progress across a batch. `detail` carries whatever the current step can say beyond the file
 * counts, such as which block of a large file is in flight.
 */
export type ProgressCallback = (completed: number, total: number, detail?: string) => void;

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

/**
 * Reads a browser File into base64. Only for files small enough to hold in memory whole - a base64
 * string is UTF-16, so this costs roughly 2.7x the file size. Large files go through
 * uploadLargeFile instead, which never holds more than one block.
 */
export async function readFile(file: File): Promise<FileInfo> {
    return {
        name: file.name,
        type: file.type,
        size: file.size,
        content: await readBlobAsBase64(file, file.name)
    };
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

export function attachmentEntityFor(target: UploadTarget): BlockUploadEntity {
    return isActivityMimeAttachment(target) ? "activitymimeattachment" : "annotation";
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
 * The same record, shaped for the block upload actions rather than createRecord.
 *
 * Differences: the inline file content must not be present (the actions reject a target that
 * carries it), the target needs an explicit `@odata.type`, and Dataverse does not allocate an id
 * for annotations here so the caller has to supply one.
 */
export function buildBlockUploadTarget(target: UploadTarget, file: File): ComponentFramework.WebApi.Entity {
    const entityType = attachmentEntityFor(target);
    const record = buildAttachmentRecord(target, { name: file.name, type: file.type, size: file.size, content: "" });

    delete record.body;
    delete record.documentbody;

    record["@odata.type"] = `Microsoft.Dynamics.CRM.${entityType}`;

    if (entityType === "annotation") {
        record.annotationid = newUuid();
    }

    return record;
}

/**
 * Uploads the files one after another. Each file is wrapped in its own try/catch so a single
 * failure no longer abandons the rest of the batch; the caller decides what to show.
 *
 * Takes File objects rather than pre-read base64 so that large files can be streamed a block at a
 * time. Anything at or below the threshold keeps going out as a single createRecord.
 */
export async function uploadToDataverse(
    context: ComponentFramework.Context<IInputs>,
    target: UploadTarget,
    files: File[],
    onProgress?: ProgressCallback
): Promise<UploadResult[]> {
    const results: UploadResult[] = [];
    const attachmentEntity = attachmentEntityFor(target);

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
            if (file.size > CHUNK_THRESHOLD_BYTES) {
                await uploadLargeFile({
                    clientUrl: getClientUrl(),
                    entityType: attachmentEntity,
                    target: buildBlockUploadTarget(target, file),
                    file,
                    onBlock: (done, total) => onProgress?.(i, files.length, `${file.name} (${done}/${total})`)
                });
            } else {
                await context.webAPI.createRecord(attachmentEntity, buildAttachmentRecord(target, await readFile(file)));
            }
            results.push({ fileName: file.name, succeeded: true });
        } catch (e) {
            results.push({ fileName: file.name, succeeded: false, error: toErrorMessage(e) });
        }
        onProgress?.(i + 1, files.length);
    }

    return results;
}

/**
 * The organization wide cap Dataverse applies to uploads, in bytes. It governs the base64 encoded
 * length rather than the raw file, so a caller comparing it against `File.size` has to scale it.
 *
 * Resolves to 0 when it cannot be read, which callers treat as "unknown" - failing to read a
 * setting should not block uploads that would have worked.
 */
export async function getMaxUploadFileSize(context: ComponentFramework.Context<IInputs>): Promise<number> {
    try {
        const response = await context.webAPI.retrieveMultipleRecords("organization", "?$select=maxuploadfilesize", 1);
        const value = response.entities?.[0]?.maxuploadfilesize as number | undefined;
        return typeof value === "number" && value > 0 ? value : 0;
    } catch (e) {
        console.warn("[AttachmentUpload] Could not read the organization max upload file size.", e);
        return 0;
    }
}

export function toErrorMessage(error: unknown): string {
    if (error instanceof Error) {
        return error.message;
    }
    const message = (error as { message?: string })?.message;
    return message ?? String(error);
}
