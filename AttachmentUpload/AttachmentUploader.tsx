import * as React from "react";
import { IInputs, IOutputs } from "./generated/ManifestTypes"
import {useDropzone} from 'react-dropzone'
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faSpinner, faRecordVinyl } from '@fortawesome/free-solid-svg-icons'

export interface UploadProps {
    id: string;
    entityName:string;
    entitySetName:string;
    controlToRefresh:string|null;
    uploadIcon:string;
    context: ComponentFramework.Context<IInputs>;
}

export interface FileInfo {
  name: string;
  type: string;
  body: string;
}

export const AttachmentUploader: React.FC<UploadProps> = (uploadProps: UploadProps) => {
    
    const [uploadIcn,setuploadIcn]=React.useState(uploadProps.uploadIcon);
    const [totalFileCount, setTotalFileCount] = React.useState(0);
    const [currentUploadCount, setCurrentUploadCount] = React.useState(0);
    const translate = (name:string) => uploadProps.context.resources.getString(name);

    const onDrop = React.useCallback((acceptedFiles:any) => {
       
        if(acceptedFiles && acceptedFiles.length){
            setTotalFileCount(acceptedFiles.length);
        }
       

        const toBase64 = async (file:any) => new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result);
            reader.onabort = () => reject();
            reader.onerror = error => reject(error);
        });

        const uploadFileToRecord = async (id: string, entity: string,entitySetName:string,
        fileInfo:  FileInfo,context: ComponentFramework.Context<IInputs>)=>{

            let isActivityMimeAttachment = (entity.toLowerCase() === "email" || entity.toLowerCase() === "appointment");
            let attachmentRecord: ComponentFramework.WebApi.Entity = {};
            if (isActivityMimeAttachment) {
                attachmentRecord["objectid_activitypointer@odata.bind"] = `/activitypointers(${id})`;
                attachmentRecord["body"] = fileInfo.body;
            }
            else {
                attachmentRecord[`objectid_${entity}@odata.bind`] = `/${entitySetName}(${id})`;
                attachmentRecord["documentbody"] = fileInfo.body;
            }
       
            if(fileInfo.type && fileInfo.type!==""){
              attachmentRecord["mimetype"] =fileInfo.type;
            }

            attachmentRecord["filename"] = fileInfo.name;
            attachmentRecord["objecttypecode"] = entity;
            let attachmentEntity = isActivityMimeAttachment ? "activitymimeattachment" : "annotation";
            await context.webAPI.createRecord(attachmentEntity, attachmentRecord)
        }


        const uploadFilesToCRM = async (files: any) => {

         
          try{

          
            for(let i=0;i<acceptedFiles.length;i++)
            {
                setCurrentUploadCount(i);
                let file=acceptedFiles[i] as any;
                let base64Data=await toBase64(acceptedFiles[i]);
                let base64DataStr=base64Data as string;
                var base64 =base64DataStr.replace(/^data:.+;base64,/, '');
                
                let fileInfo:FileInfo ={name:file.name,type:file.type,body:base64};
              let entityId = uploadProps.id;
              let entityName = uploadProps.entityName;

              if (entityId == null || entityId === "") {//this happens when the record is created and the user tries to upload
                let currentPageContext = uploadProps.context as any;
                currentPageContext = currentPageContext ? currentPageContext["page"] : undefined;
                entityId = currentPageContext.entityId;
                entityName = currentPageContext.entityTypeName;
              }

                await uploadFileToRecord(entityId,entityName,uploadProps.entitySetName, fileInfo,uploadProps.context);
            }
          }
          catch(e){         
            let errorMessagePrefix=(acceptedFiles.length===1) ? translate("error_while_uploading_attachment") : translate("error_while_uploading_attachments");
            let errOptions={message:`${errorMessagePrefix} ${e.message}`};
            uploadProps.context.navigation.openErrorDialog(errOptions)
          }

            setTotalFileCount(0);
            let xrmObj: any = (window as any)["Xrm"];
            if (xrmObj && xrmObj.Page && uploadProps.controlToRefresh) {
                var controlToRefresh = xrmObj.Page.getControl(uploadProps.controlToRefresh);
                if (controlToRefresh) {
                    controlToRefresh.refresh();
                }
            }
        }


        uploadFilesToCRM(acceptedFiles);




      }, [totalFileCount,currentUploadCount])
  
        const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop })

  
    if (uploadProps.id==null ||uploadProps.id === "") {
      return (
<div className={"defaultContentCont"}>
{translate("save_record_to_enable_content")}
</div>
      );
    }

        let fileStats = null;
        if (totalFileCount > 0) {
          fileStats = (          
            <div className={"filesStatsCont uploadDivs"}>
              <div>
                {/* <img src={spinner} alt="processing" /> */}
                <FontAwesomeIcon icon={faSpinner} inverse size="2x" spin/>
              </div>
              <div className={"uploadStatusText"}>
              {translate("uploading")} ({currentUploadCount}/{totalFileCount})
             </div>
            </div>
          );
        }      
      


        return (
          <div className={"dragDropCont"}>
            <div className={"dropZoneCont uploadDivs"} {...getRootProps()} style={{ backgroundColor: isDragActive ? '#F8F8F8' : 'white' }}>
              <input {...getInputProps()} />
              <div>
                <img className={"uploadImgDD"} src={uploadIcn} alt="Upload" />                
              </div>
              <div>
                {
                  isDragActive ?
                    <p>{translate("drop_files_here")}</p> :
                    <p>{translate("drop_files_here_or_click_to_upload")}</p>
                }
              </div>
            </div>
      
            { fileStats }
      
          </div>
        )
      


}