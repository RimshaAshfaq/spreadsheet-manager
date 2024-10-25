import React from 'react';
import './Editor.css'; // Import CSS for styling

const Editor = ({ fileContent, onSave, fixedRows, fixedCols }) => {
const handleCellChange = (rowIndex, colIndex, value) => {
const updatedContent = [...fileContent];
updatedContent[rowIndex][colIndex] = value;
onSave(updatedContent);
};

return (
<div className="editor-container">
    <div className="header-row">
    {/* Column labels (A, B, C, ...) */}
    {Array.from({ length: fixedCols }, (_, colIndex) => (
        <div key={colIndex} className="header-cell">{String.fromCharCode(65 + colIndex)}</div>
    ))}
    </div>
    {fileContent.slice(0, fixedRows).map((row, rowIndex) => (
    <div key={rowIndex} className="row">
        {/* Row number (1, 2, 3, ...) */}
        <div className="row-number">{rowIndex + 1}</div>
        {row.slice(0, fixedCols).map((cell, colIndex) => (
        <input
            key={colIndex}
            type="text"
            value={cell}
            onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
            style={{ width: '100px', height: '40px', fontSize: '16px', padding: '5px' }} // Adjust size and style as needed
        />
        ))}
    </div>
    ))}
</div>
);
};

export default Editor;
