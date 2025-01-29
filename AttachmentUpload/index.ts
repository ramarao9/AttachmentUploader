import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { AttachmentUploader, UploadProps } from './AttachmentUploader';


interface EntityRef {
	id: string,
	entityName: string
}

export class AttachmentUpload implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private UploadIconName: string = "uploadicn.png";
	private attachmentUploaderContainer: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private uploadProps: UploadProps = {
		id: "",
		entityName: "",
		entitySetName: "",
		controlToRefresh: "",
		uploadIcon: "",
		useNoteAttachment: false,
		context: undefined,
		defaultNoteTitle: ""
	};
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		let entityRef = this.getEntityReference(context);
		if (entityRef) {
			this.uploadProps.id = entityRef.id
			this.uploadProps.entityName = entityRef.entityName;
			this.uploadProps.context = context;
			this.uploadProps.controlToRefresh = context.parameters.ControlNameForRefresh.raw;
			this.uploadProps.uploadIcon = this.getImageBase64();
			this.uploadProps.useNoteAttachment = context.parameters.UseNoteAttachment.raw === "1";
			this.uploadProps.defaultNoteTitle = context.parameters.DefaultNoteTitle.raw;
			this.uploadProps.context = context;
		}
		this.attachmentUploaderContainer = container;
	}

	//since we want the image to also render during dev, we can use this approach until better support is provided in the future.
	private getImageBase64() {
		return "data:image/png;base64,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";
	}


	private getEntityReference(context: ComponentFramework.Context<IInputs>): EntityRef | undefined {
		let currentPageContext = context as any;
		currentPageContext = currentPageContext ? currentPageContext["page"] : undefined;
		var entityRef: EntityRef = { id: "", entityName: "" };
		if (currentPageContext) {
			if (currentPageContext.entityTypeName) {
				entityRef.entityName = currentPageContext.entityTypeName;
			}
			if (currentPageContext.entityId && currentPageContext.entityId !== "") {
				entityRef.id = currentPageContext.entityId;
			}
		}

		return entityRef;

	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this.uploadProps.context = context;

		let entityRef = this.getEntityReference(context);
		if (entityRef && entityRef.id !== "") {
			this.uploadProps.id = entityRef.id;
		}
		this.uploadProps.controlToRefresh = context.parameters.ControlNameForRefresh.raw;
		this.uploadProps.uploadIcon = this.getImageBase64();//when initially a new record tha's transitioning to an existing record, so the UI is now being updated to enable the content

		this.uploadProps.defaultNoteTitle = context.parameters.DefaultNoteTitle.raw;
		if (this.uploadProps.entitySetName === "") {
			this.retrieveEntitySetNameAndRender(context, this.uploadProps.entityName);
		}
		else {
			this.renderComponent();
		}


	}

	private retrieveEntitySetNameAndRender(context: ComponentFramework.Context<IInputs>, entityName: string) {
		var thisRef = this;
		context.utils.getEntityMetadata(entityName).then(function (response) {
			thisRef.uploadProps.entitySetName = response.EntitySetName;
			thisRef.renderComponent();
		},
			function (errorResponse: any) {
				console.log(`Error occurred while retrieving the entity metadata. ${errorResponse}`);
			});
	}

	private renderComponent() {
		ReactDOM.render(
			React.createElement(
				AttachmentUploader,
				this.uploadProps
			),
			this.attachmentUploaderContainer
		);
	}
	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.attachmentUploaderContainer);
	}
}