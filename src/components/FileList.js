import React from 'react';

const FileList = ({ files }) => {
return (
<div>
    <h2>Uploaded Files</h2>
    <ul>
    {files.map((file, index) => (
        <li key={index}>
        {file.name}
        {/* Add buttons for edit, delete, download */}
        <button onClick={() => {/* Logic to edit the file */}}>Edit</button>
        <button onClick={() => {/* Logic to delete the file */}}>Delete</button>
        <button onClick={() => {/* Logic to download the file */}}>Download</button>
        </li>
    ))}
    </ul>
</div>
);
};

export default FileList;
