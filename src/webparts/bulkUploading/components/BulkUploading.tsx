import * as React from "react";
import { InboxOutlined } from "@ant-design/icons";
import type { UploadProps } from "antd";
import { message, Upload } from "antd";
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IBulkUploadingProps } from "./IBulkUploadingProps";
import styles from "./BulkUploading.module.scss";

const { Dragger } = Upload;

const defaultUploadProps: UploadProps = {
  name: "file",
  multiple: true,
  action: "",
};

const BulkUploading = (props: IBulkUploadingProps) => {
  const getFileBuffer = async (file: File): Promise<ArrayBuffer | null> => {
    return new Promise((resolve, reject) => {
      const fileReader = new FileReader();

      fileReader.onerror = (event: ProgressEvent<FileReader>) => {
        reject(event.target?.error);
      };

      fileReader.onloadend = (event: ProgressEvent<FileReader>) => {
        resolve(event.target?.result as ArrayBuffer);
      };

      fileReader.readAsArrayBuffer(file);
    });
  };

  const uploadFile = async (
    fileData: ArrayBuffer | null,
    fileName: string
  ): Promise<void> => {
    if (!fileData) {
      throw new Error("No file data found");
    }

    try {
      const endpoint = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('File Uploads')/RootFolder/Files/add(overwrite=true,url='${fileName}')`;

      const options: ISPHttpClientOptions = {
        headers: { "CONTENT-LENGTH": fileData.byteLength.toString() },
        body: fileData,
      };

      await props.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        options
      );
    } catch (error) {
      message.error(`Error uploading file: ${error}`);
    }
  };

  const handleUpload = async (info: any) => {
    const { file } = info;
    const fileData = await getFileBuffer(file.originFileObj);

    if (fileData) {
      await uploadFile(fileData, file.name);
    }
  };

  const customStyles = `
  
  :where(.css-dev-only-do-not-override-1rqnfsa).ant-upload-wrapper .ant-upload-drag {
    width: 95%;
    margin: auto;
  }
  
  `;

  return (
    <div className={styles.card}>
      <style>{customStyles}</style>
      <div className={styles.headerBox}>File Upload</div>
      <div className={styles.contentBox}>
        <img
          src={require("../assets/server.png")}
          alt="Upload File"
          className={styles.uploadImg}
        />
        <p className={styles.text}>
          upload your{" "}
          <span style={{ fontWeight: "600", textDecoration: "underline" }}>
            files
          </span>{" "}
          to SharePoint
        </p>
      </div>
      <Dragger {...defaultUploadProps} {...props} onChange={handleUpload}>
        <p className="ant-upload-drag-icon">
          <InboxOutlined rev={undefined} />
        </p>
        <p className="ant-upload-text">
          Click or drag file to this area to upload
        </p>
        <p className="ant-upload-hint">
          Support for a single or bulk upload. Strictly prohibited from
          uploading company data or other banned files.
        </p>
      </Dragger>
    </div>
  );
};

export default BulkUploading;
