# Attachment Uploader
A Power App code component to easily upload one or more attachments on Dynamics365 records. Works with Email and normal Notes attachments.

Works in both **model-driven apps** and **canvas apps**, and follows the app's modern (Fluent UI v9) theme, including dark mode, wherever the host provides one.

Large files are uploaded in blocks so they are not limited by a single request, the environment's own
maximum upload size is respected, and one file failing no longer abandons the rest of the batch.


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



## Large files and limits

In a model-driven app:

* Files larger than **10 MB** are uploaded to Dataverse in blocks instead of a single request. There is
  nothing to configure — the control picks the path per file, and the spinner shows which block is in
  flight next to the file name.

* The environment's **maximum upload file size** is read from the organization record and checked
  before anything is sent. A file over it is rejected with *"exceeds the maximum upload size configured
  for this environment."* — deliberately worded differently from the `MaxFileSizeKB` message, because
  the fix is an administrator raising the environment limit rather than editing a property on the
  control. If the setting cannot be read, the check is skipped rather than blocking uploads that would
  have worked.

* Every file in a batch is attempted. Any that fail are listed by name with the error Dataverse
  returned, and the ones that succeeded still land.


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

> Base64 is roughly 33% larger than the file itself, and the whole batch is held in app memory, so
> canvas apps fall back to a **10 MB per file** cap when `MaxFileSizeKB` is left empty. Set
> `MaxFileSizeKB` to raise or lower it.


### Showing progress while the app uploads

Once the control has handed the files over, it has no visibility into what Power Fx does with them, so
it cannot show progress on its own. `IsProcessing` and `ProcessingFileName` let the app drive the same
busy state a model-driven upload shows.

Power Apps does not repaint in the middle of a formula, which means setting these inside the `ForAll`
above displays nothing until the whole formula finishes. Drive one file per **Timer** tick instead:

```powerfx
// AttachmentUpload1.OnChange
Set(varQueue, ParseJSON(AttachmentUpload1.SelectedFiles));
Set(varIndex, 1);
Set(varBusy, CountRows(varQueue) > 0)

// Timer1.OnTimerEnd   (Duration 100, Repeat true, Start: varBusy)
Set(varFile, Index(varQueue, varIndex));
Patch(
    Notes,
    Defaults(Notes),
    {
        Title:       Text(varFile.name),
        'File Name': Text(varFile.name),
        Document:    Text(varFile.content),
        Regarding:   Gallery1.Selected
    }
);
Set(varIndex, varIndex + 1);
Set(varBusy, varIndex <= CountRows(varQueue))

// AttachmentUpload1 properties
IsProcessing:       varBusy
ProcessingFileName: Text(varFile.name)
```

While `IsProcessing` is true the drop zone shows the spinner and refuses further drops, and the
"file(s) uploaded." confirmation is held back until it goes false — so the message appears when the app
is actually done rather than when the control finished reading. Set `ShowSuccessMessage` to `No` to
suppress that message entirely.


## Properties

| Property | Applies to | Description |
|---|---|---|
| `Attribute` | Model-driven | The (unused) field the control is bound to on the form. |
| `ControlNameForRefresh` | Model-driven | Name of a form control — typically the Timeline — to refresh after an upload. |
| `DefaultNoteTitle` | Model-driven | Subject applied to the created note. |
| `UseNoteAttachment` | Model-driven | `Yes` uploads to Notes even on Email/Appointment, instead of ActivityMimeAttachment. |
| `HostMode` | Both | `Auto` (default) detects the host. `ModelDriven` / `Canvas` force a behaviour. |
| `MaxFileSizeKB` | Both | Largest accepted file size in KB. Empty or `0` means no limit in a model-driven app; canvas apps fall back to a 10 MB per file cap. |
| `AcceptedFileTypes` | Both | Comma separated extensions or mime types, e.g. `.pdf,.docx,image/*`. Empty accepts everything. |
| `ShowSuccessMessage` | Canvas | `Yes` (default) shows the "file(s) uploaded." confirmation after a batch. |
| `IsProcessing` | Canvas | Set to true while the app is uploading — shows the spinner and blocks further drops. |
| `ProcessingFileName` | Canvas | Name of the file the app is currently uploading, shown next to the spinner. |
| `SelectedFiles` (output) | Canvas | JSON array of the files in the last batch. |
| `SelectedFilesCount` (output) | Canvas | Number of files in the last batch. |
| `LastBatchId` (output) | Canvas | Increments on every batch so OnChange always fires. |


## Development

After cloning the project, run the below commands

`npm install` -- installs the required dependencies

`npm run start` -- local development and testing

`npm run build`  -- to build for production

`npm run rebuild` -- clean build

`npm run lint` -- runs ESLint (flat config in `eslint.config.mjs`); `npm run lint:fix` applies the fixes


If you are new to PCF, the [official documentation](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/implementing-controls-using-typescript) is a good place to start.
