import { IInputs } from "./generated/ManifestTypes";

/**
 * The two hosts this control supports. They differ in one fundamental way:
 * Dataverse dependent APIs (WebAPI, Utility) are only available in model driven apps, so in a
 * canvas app the control can only read the files and hand them back through its output properties.
 * https://learn.microsoft.com/en-us/power-apps/developer/component-framework/limitations
 */
export type HostMode = "modeldriven" | "canvas";

export interface EntityRef {
    id: string;
    entityName: string;
}

/**
 * Reads the record the control is sitting on. `context.mode.contextInfo` is the supported API;
 * the legacy `context.page` bag is kept as a fallback for older model driven hosts.
 */
export function getEntityReference(context: ComponentFramework.Context<IInputs>): EntityRef {
    const entityRef: EntityRef = { id: "", entityName: "" };
    if (!context) {
        return entityRef;
    }

    const contextInfo = (context.mode as unknown as { contextInfo?: { entityId?: string; entityTypeName?: string } })?.contextInfo;
    const legacyPage = (context as unknown as { page?: { entityId?: string; entityTypeName?: string } }).page;
    const source = contextInfo?.entityTypeName ? contextInfo : legacyPage;

    if (source) {
        if (source.entityTypeName) {
            entityRef.entityName = source.entityTypeName;
        }
        if (source.entityId) {
            entityRef.id = source.entityId;
        }
    }

    return entityRef;
}

/**
 * Honours the HostMode property when the maker set it explicitly, otherwise detects the host.
 *
 * Detection has to be a *positive* test for model driven. Canvas apps still hand the control a
 * `webAPI` and a `utils` object - the methods are present and only throw "Method not implemented"
 * once called - so `typeof context.webAPI.createRecord === "function"` reports model driven
 * everywhere, and `context.page.entityTypeName` is populated in canvas as well. `mode.contextInfo`
 * is model driven only, so anything that matches neither signal is treated as canvas.
 */
export function resolveHostMode(context: ComponentFramework.Context<IInputs>): HostMode {
    const configured = context?.parameters?.HostMode?.raw;
    if (configured === "1") {
        return "modeldriven";
    }
    if (configured === "2") {
        return "canvas";
    }

    const contextInfo = (context?.mode as unknown as { contextInfo?: { entityTypeName?: string } })?.contextInfo;
    if (contextInfo?.entityTypeName) {
        return "modeldriven";
    }

    // Hosts that predate contextInfo only expose the legacy page bag, which canvas fills in too,
    // so pair it with the Xrm global. Deliberately checks this window rather than walking up to
    // window.top like findXrm() does: a custom page is canvas but is hosted inside a model driven
    // app, so its parent frame does have Xrm.
    const legacyPage = (context as unknown as { page?: { entityTypeName?: string } }).page;
    const hasOwnXrm = (window as unknown as { Xrm?: { Page?: unknown } }).Xrm?.Page != null;

    return legacyPage?.entityTypeName && hasOwnXrm ? "modeldriven" : "canvas";
}

/**
 * True when the host stubbed out a Dataverse API instead of implementing it, which is how canvas
 * reports that WebAPI/Utility are unavailable. Used as a safety net for a bad host detection.
 */
export function isApiUnavailable(error: unknown): boolean {
    const message = (error instanceof Error ? error.message : (error as { message?: string })?.message) ?? "";
    return /not implemented|not supported|not available/i.test(message);
}

/**
 * Resolves the entity set name used to build the `objectid_<entity>@odata.bind` url.
 * Model driven only. Unlike the previous implementation this rejects on failure instead of
 * swallowing the error, so the caller can surface it rather than rendering nothing at all.
 */
export async function getEntitySetName(context: ComponentFramework.Context<IInputs>, entityName: string): Promise<string> {
    const metadata = await context.utils.getEntityMetadata(entityName);
    // EntityMetadata is loosely typed (index signature returning any), so narrow explicitly.
    return metadata.EntitySetName as string;
}

interface XrmControl {
    refresh?: () => void;
    getName?: () => string;
}

interface XrmLike {
    Page?: {
        getControl?: (name: string) => XrmControl | null | undefined;
        ui?: { controls?: { get(): XrmControl[] } };
    };
    Utility?: {
        getGlobalContext?: () => { getClientUrl?: () => string } | undefined;
    };
}

const LOG_PREFIX = "[AttachmentUpload]";

/**
 * Locates the Xrm global. A PCF control usually shares the form's window, but on forms rendered
 * inside a nested iframe (dialogs, side panes, some custom page hosts) `window.Xrm` is undefined
 * while the parent frame still has it, so walk up before giving up.
 *
 * Takes a predicate because callers need different parts of Xrm, and a frame can expose one
 * without the other.
 */
function findXrm(hasWhatIsNeeded: (xrm: XrmLike) => boolean): XrmLike | undefined {
    const frames: (Window | null)[] = [window, window.parent, window.top];
    for (const frame of frames) {
        try {
            const xrm = (frame as unknown as { Xrm?: XrmLike } | null)?.Xrm;
            if (xrm && hasWhatIsNeeded(xrm)) {
                return xrm;
            }
        } catch {
            // Cross origin frame, keep looking.
        }
    }
    return undefined;
}

/**
 * Base url for the Web API calls that `context.webAPI` cannot make, namely the block upload
 * actions. Xrm reports the deployment specific url, which matters wherever the organization lives
 * under a path segment rather than on its own host; the page origin is only a fallback for when
 * Xrm cannot be reached at all.
 */
export function getClientUrl(): string {
    const xrm = findXrm((candidate) => typeof candidate.Utility?.getGlobalContext === "function");

    try {
        const clientUrl = xrm?.Utility?.getGlobalContext?.()?.getClientUrl?.();
        if (clientUrl) {
            return clientUrl.replace(/\/+$/, "");
        }
    } catch (e) {
        console.warn(`${LOG_PREFIX} getClientUrl() failed, falling back to the page origin.`, e);
    }

    return window.location.origin;
}

/**
 * Refreshes a control on the model driven form (typically the timeline) after an upload.
 * No-ops in canvas apps, where the Xrm global does not exist.
 *
 * Every failure path logs, because the previous silent no-op made a mistyped control name
 * indistinguishable from a control that does not support refresh().
 */
export function refreshFormControl(controlName: string | null): void {
    const name = controlName?.trim();
    if (!name) {
        return;
    }

    const xrm = findXrm((candidate) => typeof candidate.Page?.getControl === "function");
    if (!xrm) {
        console.warn(`${LOG_PREFIX} Xrm.Page is not available, cannot refresh "${name}".`);
        return;
    }

    let control: XrmControl | null | undefined;
    try {
        control = xrm.Page?.getControl?.(name);
    } catch (e) {
        console.warn(`${LOG_PREFIX} getControl("${name}") failed.`, e);
        return;
    }

    if (!control) {
        console.warn(
            `${LOG_PREFIX} No control named "${name}" on this form. Available controls: ${listControlNames(xrm)}`
        );
        return;
    }

    if (typeof control.refresh !== "function") {
        console.warn(`${LOG_PREFIX} Control "${name}" does not expose refresh().`);
        return;
    }

    try {
        control.refresh();
    } catch (e) {
        console.warn(`${LOG_PREFIX} refresh() on "${name}" threw.`, e);
    }
}

/** Best effort list of form control names, used only to make a bad ControlNameForRefresh obvious. */
function listControlNames(xrm: XrmLike): string {
    try {
        const controls = xrm.Page?.ui?.controls?.get() ?? [];
        const names = controls.map((c) => c.getName?.()).filter(Boolean);
        return names.length > 0 ? names.join(", ") : "(none reported)";
    } catch {
        return "(unavailable)";
    }
}
