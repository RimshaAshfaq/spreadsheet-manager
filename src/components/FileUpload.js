import React from 'react';
import { useDropzone } from 'react-dropzone';

const FileUpload = ({ onFileUpload }) => {
const onDrop = (acceptedFiles) => {
// Filter for CSV files
const csvFiles = acceptedFiles.filter(file => file.type === 'text/csv');
onFileUpload(csvFiles);
};

const { getRootProps, getInputProps } = useDropzone({ onDrop, accept: { 'text/csv': [] } });

return (
<div {...getRootProps({ className: 'dropzone' })}>
    <input {...getInputProps()} />
    <p>Drag 'n' drop some CSV files here, or click to select files</p>
</div>
);
};

export default FileUpload;
