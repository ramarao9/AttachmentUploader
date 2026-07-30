# Attachment Uploader
A Power App code component to easily upload one or more attachments on Dynamics365 records. Works with Email and normal Notes attachments.

Works in both **model-driven apps** and **canvas apps**, and follows the app's modern (Fluent UI v9) theme, including dark mode.


## Installation

Download the unmanaged/managed solution from the [Releases](https://github.com/ramarao9/AttachmentUploader/releases)


### Setting up the control


* Insert a section with a single column on the form

* Add the field you would like to use that's will not be used on the form

  ![Setting control on form](https://ramarao.blob.core.windows.net/attachmentuploader/SettingControlOnForm.jpg)

* Also, uncheck 'Display label on the form' for the field

* Save and publish the form. 

* Navigate to the form and you should see the control

  ![ExistingRecord](https://ramarao.blob.core.windows.net/attachmentuploader/ExistingRecordState.jpg)


   When the record is not yet created, you would see the below

   ![NewRecordState](https://ramarao.blob.core.windows.net/attachmentuploader/NewRecordState.jpg)



* If using to upload note attachments, you could specify the name of the Timeline control as below to refresh after the upload
   
   ![TimelineRefresh](https://ramarao.blob.core.windows.net/attachmentuploader/TimelineControlRefresh.jpg)



## Canvas apps

Dataverse-dependent APIs, including the [Web API, are not available to code components in canvas
apps](https://learn.microsoft.com/en-us/power-apps/developer/component-framework/limitations). The
control therefore does **not** write to Dataverse in a canvas app — it reads the dropped files and hands
them back through its output properties, and the app decides what to do with them.

1. Enable **Power Apps component framework for canvas apps** in the environment (Admin center →
   Environment → Settings → Product → Features).
2. Import the solution, then in Power Apps Studio choose **Add** → **Get more components** → **Code** and
   import `AttachmentUpload`.
3. Read the files from the `SelectedFiles` output in the component's **OnChange**:

```powerfx
ForAll(
    ParseJSON(AttachmentUpload1.SelectedFiles),
    Patch(
        Notes,
        Defaults(Notes),
        {
            Title:       Text(ThisRecord.name),
            'File Name': Text(ThisRecord.name),
            Document:    Text(ThisRecord.content),
            Regarding:   Gallery1.Selected
        }
    )
)
```

`SelectedFiles` is a JSON array; each entry has `name`, `type`, `size` and `content` (the file bytes as
base64, without the `data:` prefix). `SelectedFilesCount` holds the number of files in the batch, and
`LastBatchId` increments on every batch so **OnChange** still fires when the same files are picked twice
in a row.

> Base64 is roughly 33% larger than the file itself, and the whole batch is held in app memory, so set
> **MaxFileSizeKB** in canvas apps rather than leaving it unlimited.


## Properties

| Property | Applies to | Description |
|---|---|---|
| `Attribute` | Model-driven | The (unused) field the control is bound to on the form. |
| `ControlNameForRefresh` | Model-driven | Name of a form control — typically the Timeline — to refresh after an upload. |
| `DefaultNoteTitle` | Model-driven | Subject applied to the created note. |
| `UseNoteAttachment` | Model-driven | `Yes` uploads to Notes even on Email/Appointment, instead of ActivityMimeAttachment. |
| `HostMode` | Both | `Auto` (default) detects the host. `ModelDriven` / `Canvas` force a behaviour. |
| `MaxFileSizeKB` | Both | Largest accepted file size in KB. Empty or `0` means no limit. |
| `AcceptedFileTypes` | Both | Comma separated extensions or mime types, e.g. `.pdf,.docx,image/*`. Empty accepts everything. |
| `SelectedFiles` (output) | Canvas | JSON array of the files in the last batch. |
| `SelectedFilesCount` (output) | Canvas | Number of files in the last batch. |
| `LastBatchId` (output) | Canvas | Increments on every batch so OnChange always fires. |


## Development

After cloning the project, run the below commands

`npm install` -- installs the required dependencies

`npm run start` -- local development and testing

`npm run build`  -- to build for production


If you are new to PCF, the [official documentation](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/implementing-controls-using-typescript) is a good place to start.
