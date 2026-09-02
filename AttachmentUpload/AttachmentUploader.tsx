import { useCallback, useMemo, useState, type ReactElement } from "react";
import { useDropzone } from "react-dropzone";
import {
    FluentProvider,
    MessageBar,
    MessageBarBody,
    MessageBarTitle,
    Spinner,
    Text,
    makeStyles,
    shorthands,
    tokens,
    type Theme
} from "@fluentui/react-components";
import { CloudArrowUp48Regular } from "@fluentui/react-icons";
import {
    ProgressCallback,
    UploadResult,
    matchesAccept,
    parseAcceptTokens,
    toErrorMessage
} from "./attachmentService";
import { HostMode } from "./hostContext";
import { Translate } from "./strings";

/**
 * Canvas has no Dataverse API: the files come back as base64 in an output property, which will not
 * survive a large file. Without a maker set limit, cap it here so a big drop fails with a message
 * instead of locking up the browser.
 */
const CANVAS_DEFAULT_MAX_BYTES = 10 * 1024 * 1024;

export interface AttachmentUploaderProps {
    mode: HostMode;
    /** False on a model driven form for a record that has not been saved yet. */
    canUpload: boolean;
    /** 0 means no limit. */
    maxFileSizeKb: number;
    /**
     * Model driven only: the largest raw file the environment accepts, already scaled down from the
     * organization's base64 cap. 0 while it is still being read, or if it could not be read.
     */
    orgMaxFileBytes: number;
    /** Comma separated `.ext` / `mime/type` / `mime/*` list. Empty accepts everything. */
    acceptAttribute: string;
    theme: Theme;
    /** Canvas only: whether to show the "N file(s) uploaded." confirmation after a batch. */
    showSuccessMessage: boolean;
    /**
     * Canvas only: the app reports that it is uploading. The control has no visibility into a Power Fx
     * upload, so this is the only way for the drop zone to show the busy state a model driven upload
     * shows on its own. Merged with, rather than replacing, the control's own busy state.
     */
    externallyBusy: boolean;
    /** Canvas only: file the app is currently uploading, shown beside the spinner. */
    externalBusyFileName: string;
    translate: Translate;
    /** Set when the control could not initialize, for example when entity metadata failed to load. */
    initializationError?: string | null;
    onFiles: (files: File[], onProgress: ProgressCallback) => Promise<UploadResult[]>;
}

const useStyles = makeStyles({
    root: {
        boxSizing: "border-box",
        display: "flex",
        flexDirection: "column",
        rowGap: tokens.spacingVerticalS,
        width: "100%",
        height: "100%",
        minHeight: "150px",
        backgroundColor: tokens.colorTransparentBackground
    },
    dropZone: {
        boxSizing: "border-box",
        flexGrow: 1,
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        rowGap: tokens.spacingVerticalXS,
        minHeight: "120px",
        paddingTop: tokens.spacingVerticalM,
        paddingBottom: tokens.spacingVerticalM,
        paddingLeft: tokens.spacingHorizontalM,
        paddingRight: tokens.spacingHorizontalM,
        ...shorthands.border(tokens.strokeWidthThick, "dashed", tokens.colorNeutralStroke2),
        borderRadius: tokens.borderRadiusMedium,
        color: tokens.colorNeutralForeground1,
        backgroundColor: tokens.colorNeutralBackground1,
        cursor: "pointer",
        outlineStyle: "none",
        ":hover": {
            ...shorthands.borderColor(tokens.colorNeutralStroke1),
            backgroundColor: tokens.colorNeutralBackground1Hover
        },
        ":focus-visible": {
            ...shorthands.borderColor(tokens.colorStrokeFocus2)
        }
    },
    dropZoneActive: {
        ...shorthands.borderColor(tokens.colorBrandStroke1),
        backgroundColor: tokens.colorNeutralBackground1Selected
    },
    dropZoneDisabled: {
        cursor: "default",
        color: tokens.colorNeutralForegroundDisabled,
        backgroundColor: tokens.colorNeutralBackgroundDisabled
    },
    dropZoneBusy: {
        cursor: "default",
        backgroundColor: tokens.colorNeutralBackground3
    },
    icon: {
        width: "48px",
        height: "48px",
        color: tokens.colorNeutralForeground3
    },
    failureList: {
        marginTop: 0,
        marginBottom: 0,
        paddingLeft: tokens.spacingHorizontalXXL
    }
});

export const AttachmentUploader = (props: AttachmentUploaderProps): ReactElement => {
    const {
        mode,
        canUpload,
        maxFileSizeKb,
        orgMaxFileBytes,
        acceptAttribute,
        theme,
        showSuccessMessage,
        externallyBusy,
        externalBusyFileName,
        translate,
        initializationError,
        onFiles
    } = props;
    const styles = useStyles();

    const [busy, setBusy] = useState(false);
    const [completed, setCompleted] = useState(0);
    const [total, setTotal] = useState(0);
    const [failures, setFailures] = useState<string[]>([]);
    const [successCount, setSuccessCount] = useState(0);
    const [detail, setDetail] = useState("");

    const maxBytes = useMemo(() => {
        if (maxFileSizeKb > 0) {
            return maxFileSizeKb * 1024;
        }
        return mode === "canvas" ? CANVAS_DEFAULT_MAX_BYTES : 0;
    }, [maxFileSizeKb, mode]);
    const acceptTokens = useMemo(() => parseAcceptTokens(acceptAttribute), [acceptAttribute]);

    // Reading the files and the app uploading them are two separate phases of the same operation in
    // canvas, so either one puts the drop zone into the busy state.
    const isBusy = busy || externallyBusy;

    const onDrop = useCallback(
        async (droppedFiles: File[]) => {
            if (!droppedFiles || droppedFiles.length === 0) {
                return;
            }

            const problems: string[] = [];
            const allowed: File[] = [];
            for (const file of droppedFiles) {
                if (maxBytes > 0 && file.size > maxBytes) {
                    problems.push(`${file.name} ${translate("file_too_large")}`);
                } else if (orgMaxFileBytes > 0 && file.size > orgMaxFileBytes) {
                    // Distinct from the maker set limit: the fix here is an administrator raising
                    // the environment's maximum upload size, not editing the control's property.
                    problems.push(`${file.name} ${translate("file_exceeds_org_limit")}`);
                } else if (!matchesAccept(file, acceptTokens)) {
                    problems.push(`${file.name} ${translate("file_type_not_allowed")}`);
                } else {
                    allowed.push(file);
                }
            }

            setSuccessCount(0);
            if (allowed.length === 0) {
                setFailures(problems);
                return;
            }

            setFailures([]);
            setBusy(true);
            setCompleted(0);
            setTotal(allowed.length);

            try {
                // The files are handed over unread: reading them all up front would hold the whole
                // batch in memory as base64, which is roughly 2.7x their size. The consumer decides
                // how to read each one.
                const results = await onFiles(allowed, (done, count, note) => {
                    setCompleted(done);
                    setTotal(count);
                    setDetail(note ?? "");
                });

                let succeeded = 0;
                for (const result of results) {
                    if (result.succeeded) {
                        succeeded++;
                    } else {
                        problems.push(result.error ? `${result.fileName}: ${result.error}` : result.fileName);
                    }
                }
                setSuccessCount(succeeded);
            } catch (e) {
                problems.push(toErrorMessage(e));
            } finally {
                setBusy(false);
                setCompleted(0);
                setTotal(0);
                setDetail("");
                setFailures(problems);
            }
        },
        [acceptTokens, maxBytes, onFiles, orgMaxFileBytes, translate]
    );

    // useDropzone expects a void-returning handler, so the promise is explicitly discarded.
    // onDrop already handles its own errors in the try/catch above.
    const handleDrop = useCallback((droppedFiles: File[]) => void onDrop(droppedFiles), [onDrop]);

    const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop: handleDrop, disabled: !canUpload || isBusy });

    const dropZoneClassName = [
        styles.dropZone,
        isDragActive ? styles.dropZoneActive : "",
        canUpload ? "" : styles.dropZoneDisabled,
        isBusy ? styles.dropZoneBusy : ""
    ]
        .filter(Boolean)
        .join(" ");

    const message = initializationError ?? (canUpload ? null : translate("save_record_to_enable_content"));

    // The control's own phase wins while it is reading files, because only it knows those counts.
    // Once it has handed them over, the app is the only thing that knows what is happening.
    const busyLabel = busy
        ? [`${translate("uploading")} (${completed}/${total})`, detail].filter(Boolean).join(" ")
        : [translate("uploading"), externalBusyFileName].filter(Boolean).join(" ");

    return (
        <FluentProvider theme={theme} className={styles.root}>
            {message != null && (
                <MessageBar intent={initializationError ? "error" : "info"}>
                    <MessageBarBody>{message}</MessageBarBody>
                </MessageBar>
            )}

            <div {...getRootProps({ className: dropZoneClassName })}>
                <input {...getInputProps({ accept: acceptAttribute || undefined })} />

                {isBusy ? (
                    <Spinner size="large" labelPosition="below" label={busyLabel} />
                ) : (
                    <>
                        <CloudArrowUp48Regular className={styles.icon} />
                        <Text align="center">
                            {isDragActive
                                ? translate("drop_files_here")
                                : translate("drop_files_here_or_click_to_upload")}
                        </Text>
                    </>
                )}
            </div>

            {failures.length > 0 && (
                <MessageBar intent="error">
                    <MessageBarBody>
                        <MessageBarTitle>
                            {failures.length === 1
                                ? translate("error_while_uploading_attachment")
                                : translate("error_while_uploading_attachments")}
                        </MessageBarTitle>
                        <ul className={styles.failureList}>
                            {failures.map((failure, index) => (
                                <li key={`${index}-${failure}`}>{failure}</li>
                            ))}
                        </ul>
                    </MessageBarBody>
                </MessageBar>
            )}

            {/* successCount only means "handed to the app" in canvas, so holding the confirmation back
                until the app stops uploading keeps it from claiming success mid upload. */}
            {mode === "canvas" && showSuccessMessage && !externallyBusy && successCount > 0 && failures.length === 0 && (
                <MessageBar intent="success">
                    <MessageBarBody>{`${successCount} ${translate("files_uploaded")}`}</MessageBarBody>
                </MessageBar>
            )}
        </FluentProvider>
    );
};
