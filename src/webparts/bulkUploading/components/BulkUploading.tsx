import * as React from "react";
import { InboxOutlined } from "@ant-design/icons";
import type { UploadProps } from "antd";
import { message, Upload } from "antd";
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IBulkUploadingProps } from "./IBulkUploadingProps";

const { Dragger } = Upload;

const defaultUploadProps: UploadProps = {
  name: "file",
  multiple: true,
  action: "",
  onChange(info) {
    const { status } = info.file;
    if (status !== "uploading") {
      console.log(info.file, info.fileList);
    }
    if (status === "done") {
      message.success(`${info.file.name} file uploaded successfully.`);
    } else if (status === "error") {
      message.error(`${info.file.name} file upload failed.`);
    }
  },
};

const FileUploading = (props: IBulkUploadingProps) => {
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

      const response = await props.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        options
      );

      if (response.status === 200) {
        message.success("File uploaded successfully");
      } else {
        message.error(`Error uploading file: ${response.statusText}`);
      }
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

  return (
    <Dragger {...defaultUploadProps} {...props} onChange={handleUpload}>
      <p className="ant-upload-drag-icon">
        <InboxOutlined rev={undefined} />
      </p>
      <p className="ant-upload-text">
        Click or drag file to this area to upload
      </p>
      <p className="ant-upload-hint">
        Support for a single or bulk upload. Strictly prohibited from uploading
        company data or other banned files.
      </p>
    </Dragger>
  );
};

export default FileUploading;
