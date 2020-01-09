import * as React from "react";
import { IInputs, IOutputs } from "./generated/ManifestTypes"
import {useDropzone} from 'react-dropzone'
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faSpinner, faRecordVinyl } from '@fortawesome/free-solid-svg-icons'

export interface UploadProps {
    id: string;
    entityName:string,
    controlToRefresh:string|null,
    uploadIcon:string,
    context: ComponentFramework.Context<IInputs>;
}

export const AttachmentUploader: React.FC<UploadProps> = (uploadProps: UploadProps) => {

    const [uploadIcn,setuploadIcn]=React.useState(uploadProps.uploadIcon);
    const [totalFileCount, setTotalFileCount] = React.useState(0);
    const [currentUploadCount, setCurrentUploadCount] = React.useState(0);
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

        const uploadFileToRecord = async (id: string, entity: string,
             filename: string, body: string,context: ComponentFramework.Context<IInputs>)=>{

            let isActivityMimeAttachment = (entity.toLowerCase() === "email" || entity.toLowerCase() === "appointment");
            let attachmentRecord: ComponentFramework.WebApi.Entity = {};
            if (isActivityMimeAttachment) {
                attachmentRecord["objectid_activitypointer@odata.bind"] = `/activitypointers(${id})`;
                attachmentRecord["body"] = body;
            }
            else {
                attachmentRecord[`objectid_${entity}@odata.bind`] = `/${entity}s(${id})`;
                attachmentRecord["documentbody"] = body;
            }
       
            attachmentRecord["filename"] = filename;
            attachmentRecord["objecttypecode"] = entity;
            let attachmentEntity = isActivityMimeAttachment ? "activitymimeattachment" : "annotation";
            await context.webAPI.createRecord(attachmentEntity, attachmentRecord)
        }


        const uploadFilesToCRM = async (files: any) => {
            for(let i=0;i<acceptedFiles.length;i++)
            {
                setCurrentUploadCount(i);
                let file=acceptedFiles[i] as any;
                let base64Data=await toBase64(acceptedFiles[i]);
                let base64DataStr=base64Data as string;
                var base64 =base64DataStr.replace(/^data:.+;base64,/, '');
                
                await uploadFileToRecord(uploadProps.id,uploadProps.entityName,
                    file.name,base64,uploadProps.context);
            }

            setTotalFileCount(0);
            let xrmObj: any = (window as any)["Xrm"];
            if (xrmObj && xrmObj.Page && xrmObj.Page) {
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
To enable the content create the record.
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
                Uploading ({currentUploadCount}/{totalFileCount})
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
                    <p>Drop the files here ...</p> :
                    <p>Drop files here or click to upload.</p>
                }
              </div>
            </div>
      
            { fileStats }
      
          </div>
        )
      


}