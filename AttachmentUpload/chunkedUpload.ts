import { newUuid, readBlobAsBase64 } from "./base64";

/**
 * Uploads a file to Dataverse in blocks, for files too large to send as a single base64 string.
 *
 * PCF's `context.webAPI` only exposes createRecord/retrieveRecord/retrieveMultipleRecords/
 * updateRecord/deleteRecord - there is no `execute` - so the InitializeXBlocksUpload / UploadBlock /
 * CommitXBlocksUpload actions cannot be called through it at all. In a model driven app the control
 * runs same origin with the form, so a plain fetch against the Web API carries the session cookie
 * and works. This is the same step outside the sanctioned PCF surface that findXrm() in
 * hostContext.ts already takes for refreshFormControl.
 *
 * https://learn.microsoft.com/en-us/power-apps/developer/data-platform/attachment-annotation-files
 */

/** Files at or below this go through the single request createRecord path instead. */
export const CHUNK_THRESHOLD_BYTES = 10 * 1024 * 1024;

/**
 * Microsoft's SDK sample uses exactly this, while the Web API docs say "less than 4 MB". If a block
 * ever comes back 413, this is the thing to drop - 4_000_000 satisfies both readings.
 */
export const BLOCK_SIZE_BYTES = 4 * 1024 * 1024;

const API_VERSION = "v9.2";

export type BlockUploadEntity = "annotation" | "activitymimeattachment";

export interface LargeFileUpload {
    clientUrl: string;
    entityType: BlockUploadEntity;
    /** The attachment record, without the inline body, carrying its `@odata.type`. */
    target: ComponentFramework.WebApi.Entity;
    file: File;
    onBlock?: (completed: number, total: number) => void;
}

export async function uploadLargeFile(upload: LargeFileUpload): Promise<void> {
    const { clientUrl, entityType, target, file, onBlock } = upload;
    // annotation -> Initialize/CommitAnnotationBlocksUpload, attachment -> ...AttachmentBlocksUpload.
    const suffix = entityType === "annotation" ? "AnnotationBlocksUpload" : "AttachmentBlocksUpload";

    const token = await initialize(clientUrl, `Initialize${suffix}`, target);

    const blockIds: string[] = [];
    const totalBlocks = Math.ceil(file.size / BLOCK_SIZE_BYTES);

    for (let offset = 0; offset < file.size; offset += BLOCK_SIZE_BYTES) {
        // Slice the file, not the base64 string: every BlockData has to be a self contained
        // encoding of its own byte range, because the server decodes each block and concatenates
        // the resulting bytes. Only one block is ever held in memory.
        const blockId = newBlockId();
        const blockData = await readBlobAsBase64(file.slice(offset, offset + BLOCK_SIZE_BYTES), file.name);

        const response = await postWithRetry(clientUrl, "UploadBlock", {
            BlockId: blockId,
            BlockData: blockData,
            FileContinuationToken: token
        });
        if (!response.ok) {
            await throwForResponse(response, "UploadBlock");
        }

        blockIds.push(blockId);
        onBlock?.(blockIds.length, totalBlocks);
    }

    // Nothing exists in Dataverse until this succeeds, so a failure part way through leaves no
    // half written record behind.
    const response = await postAction(clientUrl, `Commit${suffix}`, {
        Target: target,
        BlockList: blockIds,
        FileContinuationToken: token
    });
    if (!response.ok) {
        await throwForResponse(response, `Commit${suffix}`);
    }
}

async function initialize(clientUrl: string, action: string, target: ComponentFramework.WebApi.Entity): Promise<string> {
    const response = await postAction(clientUrl, action, { Target: target });
    if (!response.ok) {
        await throwForResponse(response, action);
    }

    const payload = (await response.json()) as { FileContinuationToken?: string };
    if (!payload.FileContinuationToken) {
        throw new Error(`${action} did not return a file continuation token.`);
    }
    return payload.FileContinuationToken;
}

function postAction(clientUrl: string, action: string, body: unknown): Promise<Response> {
    return fetch(`${clientUrl}/api/data/${API_VERSION}/${action}`, {
        method: "POST",
        credentials: "include",
        headers: {
            "Content-Type": "application/json; charset=utf-8",
            Accept: "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
        },
        body: JSON.stringify(body)
    });
}

/**
 * A large file is a lot of sequential requests, so one blip should not lose the whole upload.
 * Only retries a rejected fetch, which means a network level failure - http errors resolve
 * normally and are handled by the caller. Re-sending the same BlockId is safe: block staging is
 * keyed by id, so a block that did land is overwritten rather than duplicated.
 */
async function postWithRetry(clientUrl: string, action: string, body: unknown): Promise<Response> {
    try {
        return await postAction(clientUrl, action, body);
    } catch {
        return await postAction(clientUrl, action, body);
    }
}

/** Surfaces the Dataverse error message so it reaches the failure list instead of a bare status. */
async function throwForResponse(response: Response, action: string): Promise<never> {
    let message = `${action} failed with status ${response.status}.`;
    try {
        const payload = (await response.json()) as { error?: { message?: string } };
        if (payload?.error?.message) {
            message = payload.error.message;
        }
    } catch {
        // Not a json body, keep the status based message.
    }
    throw new Error(message);
}

/**
 * Every block of a file has to use a BlockId of the same length. A 36 character uuid always base64
 * encodes to 48 characters, so deriving the id from one satisfies that for free.
 */
function newBlockId(): string {
    return btoa(newUuid());
}
