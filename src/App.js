import React, { useState } from 'react';
import './App.css'; // Import your CSS file
import * as XLSX from 'xlsx'; // For handling Excel files

const App = () => {
  const [data, setData] = useState([Array(35).fill().map(() => Array(15).fill(''))]); 
  const [sheetNames, setSheetNames] = useState(['Sheet1']);
  const [currentSheet, setCurrentSheet] = useState(0);
  const [isEditable, setIsEditable] = useState(false); // State to toggle edit mode

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const sheetData = workbook.SheetNames.map((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      });
      
      // Update state with data from all sheets
      setData(sheetData);
      setSheetNames(workbook.SheetNames);
      setCurrentSheet(0);
    };
    reader.readAsBinaryString(file);
  };

  const handleEdit = () => {
    setIsEditable(!isEditable); // Toggle edit mode
  };

  const handleSave = () => {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(data[currentSheet]);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetNames[currentSheet]);
    XLSX.writeFile(newWorkbook, `${sheetNames[currentSheet]}.xlsx`);
  };

  const handleDownload = () => {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(data[currentSheet]);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetNames[currentSheet]);
    XLSX.writeFile(newWorkbook, `${sheetNames[currentSheet]}.xlsx`);
  };

  const handleExit = () => {
    setData([]); // Reset data
    setSheetNames(['Sheet1']); // Reset sheet names
    setCurrentSheet(0);
    setIsEditable(false);
  };

  const handleNewSheet = () => {
    const newSheetName = `Sheet${sheetNames.length + 1}`;
    setSheetNames([...sheetNames, newSheetName]);
    setCurrentSheet(sheetNames.length);
    
    // Create a new empty sheet with 35 rows and 7 columns
    const newEmptySheet = Array(35).fill().map(() => Array(7).fill(''));
    
    setData([...data, newEmptySheet]); // Add the new sheet to data
  };

  const changeSheet = (index) => {
    setCurrentSheet(index);
  };

  return (
    <div className="App">
      <h1>Spreadsheet Manager</h1>
      <label className="custom-file-upload">
        Upload File
        <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} />
      </label>
      <button onClick={handleEdit}>{isEditable ? 'Stop Editing' : 'Edit'}</button>
      <button onClick={handleSave}>Save</button>
      <button onClick={handleDownload}>Download</button>
      <button onClick={handleExit}>Exit</button>
      <div className="scrollable-sheet">
        <table>
          <tbody>
            {data[currentSheet] && data[currentSheet].map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, colIndex) => (
                  <td key={colIndex}>
                    <input
                      type="text"
                      value={cell}
                      readOnly={!isEditable}
                      onChange={(e) => {
                        const newData = [...data];
                        newData[currentSheet][rowIndex][colIndex] = e.target.value; // Update cell value in the current sheet
                        setData(newData);
                      }}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="footer">
        <div className="footer-content">
          <button className="new-sheet-button" onClick={handleNewSheet}>New Sheet</button>
          <div className="sheet-name">
            <span>Sheet Name: {sheetNames[currentSheet]}</span>
          </div>
          <div className="sheet-buttons">
            {sheetNames.map((sheetName, index) => (
              <button key={index} onClick={() => changeSheet(index)}>{sheetName}</button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
