# Attachment Uploader
A Power App code component to easily upload one or more attachments on Dynamics365 records.


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





## Development

After cloning the project, run the below commands

`npm install` -- installs the required dependencies

`npm run start` -- local development and testing

`npm run build`  -- to build for production


If you are new to PCF, the [official documentation](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/implementing-controls-using-typescript) is a good place to start.
