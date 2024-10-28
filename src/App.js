import React, { useState } from 'react';
import './App.css'; // Import your CSS file
import * as XLSX from 'xlsx'; // For handling Excel files

const App = () => {
  const [data, setData] = useState([Array(35).fill().map(() => Array(15).fill(''))]); 
  const [sheetNames, setSheetNames] = useState(['Sheet1']);
  const [currentSheet, setCurrentSheet] = useState(0);
  const [isEditable, setIsEditable] = useState(false); // State to toggle edit mode
  const [selectedRow, setSelectedRow] = useState(null); // State for selected row
  const [selectedColumn, setSelectedColumn] = useState(null); // State for selected column
  const [menuPosition, setMenuPosition] = useState({ top: 0, left: 0 }); // Position of the dropdown menu

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
    setData([Array(35).fill().map(() => Array(15).fill(''))]); // Reset to a default empty sheet
    setSheetNames(['Sheet1']); // Reset sheet names
    setCurrentSheet(0); // Go back to the first sheet
    setIsEditable(false); // Turn off edit mode
    setSelectedRow(null); // Deselect any selected row
    setSelectedColumn(null); // Deselect any selected column
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

  // Row and Column Selection Functions
  const handleRowSelect = (rowIndex, event) => {
    setSelectedRow(rowIndex);
    setSelectedColumn(null);
    setMenuPosition({ top: event.clientY, left: event.clientX });
  };

  const handleColumnSelect = (colIndex, event) => {
    setSelectedColumn(colIndex);
    setSelectedRow(null);
    setMenuPosition({ top: event.clientY, left: event.clientX });
  };

  const addRowAbove = () => {
    const newData = [...data];
    newData[currentSheet].splice(selectedRow, 0, Array(data[currentSheet][0].length).fill(''));
    setData(newData);
    setSelectedRow(null); // Deselect row after action
  };

  const addRowBelow = () => {
    const newData = [...data];
    newData[currentSheet].splice(selectedRow + 1, 0, Array(data[currentSheet][0].length).fill(''));
    setData(newData);
    setSelectedRow(null); // Deselect row after action
  };

  const deleteRow = () => {
    const newData = [...data];
    newData[currentSheet].splice(selectedRow, 1);
    setData(newData);
    setSelectedRow(null); // Deselect row after action
  };

  const addColumnLeft = () => {
    const newData = [...data];
    newData[currentSheet].forEach(row => row.splice(selectedColumn, 0, ''));
    setData(newData);
    setSelectedColumn(null); // Deselect column after action
  };

  const addColumnRight = () => {
    const newData = [...data];
    newData[currentSheet].forEach(row => row.splice(selectedColumn + 1, 0, ''));
    setData(newData);
    setSelectedColumn(null); // Deselect column after action
  };

  const deleteColumn = () => {
    const newData = [...data];
    newData[currentSheet].forEach(row => row.splice(selectedColumn, 1));
    setData(newData);
    setSelectedColumn(null); // Deselect column after action
  };

  return (
    <div className="App">
      <h1>Spreadsheet</h1>
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
          <thead>
            <tr>
              <th></th>
              {data[currentSheet][0] && data[currentSheet][0].map((_, colIndex) => (
                <th
                  key={`col-${colIndex}`}
                  onClick={(event) => handleColumnSelect(colIndex, event)}
                  style={{
                    backgroundColor: selectedColumn === colIndex ? '#d3d3d3' : 'inherit'
                  }}
                >
                  {colIndex + 1}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data[currentSheet] && data[currentSheet].map((row, rowIndex) => (
              <tr
                key={rowIndex}
                onClick={(event) => handleRowSelect(rowIndex, event)}
                style={{
                  backgroundColor: selectedRow === rowIndex ? '#d3d3d3' : 'inherit'
                }}
              >
                <td>{rowIndex + 1}</td>
                {row.map((cell, colIndex) => (
                  <td
                    key={colIndex}
                    style={{
                      backgroundColor: selectedColumn === colIndex ? '#d3d3d3' : 'inherit'
                    }}
                  >
                    <input
                      type="text"
                      value={cell}
                      readOnly={!isEditable}
                      onChange={(e) => {
                        const newData = [...data];
                        newData[currentSheet][rowIndex][colIndex] = e.target.value;
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

      {/* Dropdown Menu for Row and Column Actions */}
      {(selectedRow !== null || selectedColumn !== null) && (
        <div
          className="dropdown"
          style={{
            position: 'absolute',
            top: menuPosition.top,
            left: menuPosition.left,
            zIndex: 1,
            backgroundColor: '#ffffff',
            border: '1px solid #ddd',
            padding: '10px'
          }}
        >
          {selectedRow !== null && (
            <>
              <button onClick={addRowAbove}>Add Row Above</button>
              <button onClick={addRowBelow}>Add Row Below</button>
              <button onClick={deleteRow}>Delete Row</button>
            </>
          )}
          {selectedColumn !== null && (
            <>
              <button onClick={addColumnLeft}>Add Column Left</button>
              <button onClick={addColumnRight}>Add Column Right</button>
              <button onClick={deleteColumn}>Delete Column</button>
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default App;
