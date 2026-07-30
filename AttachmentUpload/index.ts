import { createElement } from "react";
import { createRoot, Root } from "react-dom/client";
import { webLightTheme, type Theme } from "@fluentui/react-components";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { AttachmentUploader, AttachmentUploaderProps, ProgressCallback } from "./AttachmentUploader";
import { FileInfo, UploadResult, toErrorMessage, uploadToDataverse } from "./attachmentService";
import { EntityRef, HostMode, getEntityReference, getEntitySetName, isApiUnavailable, refreshFormControl, resolveHostMode } from "./hostContext";
import { Translate, makeTranslator } from "./strings";

export class AttachmentUpload implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private root: Root;
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged: () => void;
    private translate: Translate;

    /** Cached entity set name lookup, keyed by the entity it was started for. */
    private entitySetNameFor = "";
    private entitySetNamePromise: Promise<string> | undefined;
    private initializationError: string | null = null;

    /** Set to "canvas" when the host turned out to have no Dataverse APIs after all. */
    private hostModeOverride: HostMode | undefined;

    /** Canvas app outputs. */
    private selectedFiles = "";
    private selectedFilesCount = 0;
    private lastBatchId = 0;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.translate = makeTranslator(context);
        this.container = container;
        this.container.style.height = "100%";
        this.root = createRoot(container);
        context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;

        const mode = this.resolveMode(context);
        const entityRef = getEntityReference(context);

        if (mode === "modeldriven") {
            // Prefetch; failures are reported through initializationError and re-raised at upload time.
            void this.ensureEntitySetName(entityRef.entityName).catch(() => undefined);
        }

        this.applyAllocatedSize(context);
        this.render(context, mode, entityRef);
    }

    public getOutputs(): IOutputs {
        return {
            SelectedFiles: this.selectedFiles,
            SelectedFilesCount: this.selectedFilesCount,
            LastBatchId: this.lastBatchId
        };
    }

    public destroy(): void {
        this.root.unmount();
    }

    private render(context: ComponentFramework.Context<IInputs>, mode: HostMode, entityRef: EntityRef): void {
        const acceptAttribute = context.parameters.AcceptedFileTypes.raw ?? "";
        const theme = (context as unknown as { fluentDesignLanguage?: { tokenTheme?: Theme } }).fluentDesignLanguage?.tokenTheme ?? webLightTheme;

        const props: AttachmentUploaderProps = {
            mode,
            canUpload: this.initializationError == null && (mode === "canvas" || entityRef.id !== ""),
            maxFileSizeKb: context.parameters.MaxFileSizeKB.raw ?? 0,
            acceptAttribute,
            theme,
            // Defaults to shown, so an unset property behaves like it did before this was added.
            showSuccessMessage: context.parameters.ShowSuccessMessage.raw !== "0",
            translate: this.translate,
            initializationError: this.initializationError,
            onFiles: this.handleFiles
        };

        this.root.render(createElement(AttachmentUploader, props));
    }

    /**
     * Model driven: uploads to Dataverse and refreshes the configured form control.
     * Canvas: WebAPI is unavailable, so the files are handed back through the output properties
     * for the app to write with Power Fx.
     */
    private handleFiles = async (files: FileInfo[], onProgress: ProgressCallback): Promise<UploadResult[]> => {
        const context = this.context;
        const mode = this.resolveMode(context);

        if (mode === "canvas") {
            this.selectedFiles = JSON.stringify(files);
            this.selectedFilesCount = files.length;
            // Guarantees an output change even when the same files are picked twice in a row.
            this.lastBatchId++;
            this.notifyOutputChanged();
            onProgress(files.length, files.length);
            return files.map((file) => ({ fileName: file.name, succeeded: true }));
        }

        // The record id is empty until the form is saved, so re-read it at upload time.
        const entityRef = getEntityReference(context);
        const useNoteAttachment = context.parameters.UseNoteAttachment.raw === "1";
        const needsEntitySetName = useNoteAttachment || !this.isActivityEntity(entityRef.entityName);

        let entitySetName = "";
        if (needsEntitySetName) {
            try {
                entitySetName = await this.ensureEntitySetName(entityRef.entityName);
            } catch (e) {
                const error = toErrorMessage(e);
                return files.map((file) => ({ fileName: file.name, succeeded: false, error }));
            }
        }

        const results = await uploadToDataverse(
            context,
            {
                id: entityRef.id,
                entityName: entityRef.entityName,
                entitySetName,
                useNoteAttachment,
                defaultNoteTitle: context.parameters.DefaultNoteTitle.raw
            },
            files,
            onProgress
        );

        refreshFormControl(context.parameters.ControlNameForRefresh.raw);
        return results;
    };

    private resolveMode(context: ComponentFramework.Context<IInputs>): HostMode {
        return this.hostModeOverride ?? resolveHostMode(context);
    }

    private isActivityEntity(entityName: string): boolean {
        const entity = entityName.toLowerCase();
        return entity === "email" || entity === "appointment";
    }

    private ensureEntitySetName(entityName: string): Promise<string> {
        if (entityName === "") {
            return Promise.resolve("");
        }

        if (this.entitySetNamePromise && this.entitySetNameFor === entityName) {
            return this.entitySetNamePromise;
        }

        this.entitySetNameFor = entityName;
        this.entitySetNamePromise = getEntitySetName(this.context, entityName)
            .then((name) => {
                this.initializationError = null;
                return name;
            })
            .catch((e: unknown) => {
                // Allow a later retry.
                this.entitySetNamePromise = undefined;

                if (isApiUnavailable(e)) {
                    // The host looked model driven but has no Dataverse APIs, so switch to handing
                    // the files back through the outputs rather than showing a dead error bar.
                    this.hostModeOverride = "canvas";
                    this.initializationError = null;
                } else {
                    // Surface the failure instead of leaving the control blank.
                    this.initializationError = toErrorMessage(e);
                }

                this.render(this.context, this.resolveMode(this.context), getEntityReference(this.context));
                throw e;
            });

        return this.entitySetNamePromise;
    }

    private applyAllocatedSize(context: ComponentFramework.Context<IInputs>): void {
        const { allocatedWidth, allocatedHeight } = context.mode;
        // Canvas apps report the size the maker gave the component; model driven apps report -1.
        this.container.style.width = allocatedWidth > 0 ? `${allocatedWidth}px` : "100%";
        this.container.style.height = allocatedHeight > 0 ? `${allocatedHeight}px` : "100%";
    }
}
