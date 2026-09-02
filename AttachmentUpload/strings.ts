import { IInputs } from "./generated/ManifestTypes";

export type Translate = (key: string) => string;

const ENGLISH_LCID = 1033;

/**
 * English strings, kept in sync with AttachmentUpload.1033.resx.
 *
 * These are used directly for English users rather than going through `context.resources`. The PCF
 * test harness has no language selection: it merges every resx listed in the manifest into a single
 * string map, so whichever file happens to load last wins and the control renders in Danish or
 * Norwegian locally. Real apps pick the resx matching the user's language, so only non-English
 * users go through `getString`.
 *
 * They also serve as the fallback whenever a resx entry is missing, in which case `getString`
 * returns an empty string or echoes the key back.
 */
const ENGLISH: Record<string, string> = {
    error_while_uploading_attachment: "An error has occurred while trying to upload the attachment.",
    error_while_uploading_attachments: "One or more errors occured when trying to upload the attachments.",
    save_record_to_enable_content: "To enable the content create the record.",
    uploading: "Uploading",
    drop_files_here: "Drop the files here ...",
    drop_files_here_or_click_to_upload: "Drop files here or click to upload.",
    file_too_large: "is larger than the allowed size.",
    file_exceeds_org_limit: "exceeds the maximum upload size configured for this environment.",
    file_type_not_allowed: "is not an accepted file type.",
    files_uploaded: "file(s) uploaded."
};

export function makeTranslator(context: ComponentFramework.Context<IInputs> | undefined): Translate {
    const languageId = context?.userSettings?.languageId;
    const preferBundledEnglish = languageId == null || languageId === ENGLISH_LCID;

    return (key: string): string => {
        if (!preferBundledEnglish) {
            const resolved = context?.resources?.getString(key);
            if (resolved && resolved !== key) {
                return resolved;
            }
        }
        return ENGLISH[key] ?? key;
    };
}
