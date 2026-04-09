import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class AttachmentUploader implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private context: ComponentFramework.Context<IInputs>;
    private container: HTMLDivElement;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {

        this.context = context;
        this.container = container;

        this.createUI();
    }

    // ✅ UI BUTTON
    private createUI() {

        const fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.style.display = "none";

        fileInput.onchange = (e: any) => {
            const file = e.target.files[0];
            if (file) {
                this.readFile(file);
            }
        };

        const button = document.createElement("button");
        button.innerText = "Upload Attachment";
        button.onclick = () => fileInput.click();

        this.container.appendChild(button);
        this.container.appendChild(fileInput);
    }

    // ✅ STEP 1: READ FILE
    private readFile(file: File) {

        const reader = new FileReader();

        reader.onload = async () => {
            const base64 = (reader.result as string).split(",")[1];
            await this.createRecord(file, base64);
        };

        reader.readAsDataURL(file);
    }

    // ✅ STEP 2: CREATE RECORD + LINK CASE
    private async createRecord(file: File, base64: string) {

        const caseId = this.context.page.entityId; // 🔥 Case GUID

        console.log("Case ID:", caseId);

        const entity: any = {
            "gg_name": file.name,

            // 🔥 LINK TO CASE
            "gg_CaseAttachmentId@odata.bind": `/incidents(${caseId})`
        };

        try {

            const result = await this.context.webAPI.createRecord("gg_attachment", entity);

            const recordId = result.id;

            // ✅ STEP 3: UPLOAD FILE
            await this.uploadFile(recordId, file.name, base64);

            // ✅ STEP 4: OPEN RECORD
            this.openRecord(recordId);

        } catch (error: any) {
            console.error("Error creating record:", error.message);
        }
    }

    // ✅ STEP 3: UPLOAD FILE TO FILE COLUMN
    private async uploadFile(recordId: string, fileName: string, base64: string) {

        const clientUrl = this.context.page.getClientUrl();

        const url = `${clientUrl}/api/data/v9.2/gg_attachments(${recordId})/gg_file`;

        await fetch(url, {
            method: "PATCH",
            headers: {
                "Content-Type": "application/octet-stream",
                "x-ms-file-name": fileName,
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0"
            },
            body: Uint8Array.from(atob(base64), c => c.charCodeAt(0))
        });
    }

    // ✅ STEP 4: OPEN RECORD
    private openRecord(recordId: string) {

        const pageInput = {
            pageType: "entityrecord",
            entityName: "gg_attachment",
            entityId: recordId
        };

        const navigationOptions = {
            target: 2,
            width: { value: 70, unit: "%" },
            height: { value: 80, unit: "%" }
        };

        (window as any).Xrm.Navigation.navigateTo(pageInput, navigationOptions);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {}

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {}
}
