import React, { useState } from 'react';
import './App.css';
import * as XLSX from 'xlsx';

const App = () => {
  const [data, setData] = useState([Array(35).fill().map(() => Array(15).fill(''))]);
  const [sheetNames, setSheetNames] = useState(['Sheet1']);
  const [currentSheet, setCurrentSheet] = useState(0);
  const [isEditable, setIsEditable] = useState(true);
  const [selectedRow, setSelectedRow] = useState(null);
  const [selectedColumn, setSelectedColumn] = useState(null);
  const [menuPosition, setMenuPosition] = useState({ top: 0, left: 0 });

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
      setData(sheetData);
      setSheetNames(workbook.SheetNames);
      setCurrentSheet(0);
    };
    reader.readAsBinaryString(file);
  };

  const handleEdit = () => setIsEditable(!isEditable);
  
  const handleDownload = () => {
    const newWorkbook = XLSX.utils.book_new();
    data.forEach((sheetData, index) => {
      const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetNames[index]);
    });
    XLSX.writeFile(newWorkbook, 'Workbook.xlsx');
  };

  const handleExit = () => {
    setData([Array(35).fill().map(() => Array(15).fill(''))]);
    setSheetNames(['Sheet1']);
    setCurrentSheet(0);
    setIsEditable(false);
    setSelectedRow(null);
    setSelectedColumn(null);
  };

  const handleNewSheet = () => {
    const newSheetName = `Sheet${sheetNames.length + 1}`;
    setSheetNames([...sheetNames, newSheetName]);
    setCurrentSheet(sheetNames.length);
    const newEmptySheet = Array(35).fill().map(() => Array(15).fill(''));
    setData([...data, newEmptySheet]);
  };

  const changeSheet = (index) => setCurrentSheet(index);

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
    setSelectedRow(null);
  };

  const addRowBelow = () => {
    const newData = [...data];
    newData[currentSheet].splice(selectedRow + 1, 0, Array(data[currentSheet][0].length).fill(''));
    setData(newData);
    setSelectedRow(null);
  };

  const deleteRow = () => {
    const newData = [...data];
    newData[currentSheet].splice(selectedRow, 1);
    setData(newData);
    setSelectedRow(null);
  };

  const addColumnLeft = () => {
    const newData = [...data];
    newData[currentSheet].forEach(row => row.splice(selectedColumn, 0, ''));
    setData(newData);
    setSelectedColumn(null);
  };

  const addColumnRight = () => {
    const newData = [...data];
    newData[currentSheet].forEach(row => row.splice(selectedColumn + 1, 0, ''));
    setData(newData);
    setSelectedColumn(null);
  };

  const deleteColumn = () => {
    const newData = [...data];
    newData[currentSheet].forEach(row => row.splice(selectedColumn, 1));
    setData(newData);
    setSelectedColumn(null);
  };

  const sortData = (type) => {
    const newData = [...data];
    if (type === 'row' && selectedRow !== null) {
      newData[currentSheet][selectedRow].sort((a, b) => {
        if (!isNaN(a) && !isNaN(b)) return a - b;
        return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
      });
    } else if (type === 'column' && selectedColumn !== null) {
      const columnData = newData[currentSheet].map(row => row[selectedColumn]);
      columnData.sort((a, b) => {
        if (!isNaN(a) && !isNaN(b)) return a - b;
        return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
      });
      columnData.forEach((value, i) => newData[currentSheet][i][selectedColumn] = value);
    }
    setData(newData);
  };

  return (
    <div className="App">
      <h1>Spreadsheet Manager</h1>
      <label class="button custom-file-upload">Upload File <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} /> </label>
      <button className="button" onClick={handleEdit}>{isEditable ? 'Stop Editing' : 'Edit'}</button>
      <button className="button" onClick={handleDownload}>Download</button>
      <button className="button" onClick={handleExit}>Exit</button>
      <div className="scrollable-sheet">
        <table>
          <thead>
            <tr>
              <th></th>
              {data[currentSheet][0] && data[currentSheet][0].map((_, colIndex) => (
                <th key={`col-${colIndex}`} onClick={(event) => handleColumnSelect(colIndex, event)}
                    style={{
                      backgroundColor: selectedColumn === colIndex ? '#d3d3d3' : 'inherit'
                    }}>
                  {colIndex + 1}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data[currentSheet].map((row, rowIndex) => (
              <tr key={rowIndex} onClick={(event) => handleRowSelect(rowIndex, event)}
                  style={{
                    backgroundColor: selectedRow === rowIndex ? '#d3d3d3' : 'inherit'
                  }}>
                <td>{rowIndex + 1}</td>
                {row.map((cell, colIndex) => (
                  <td key={colIndex} style={{
                      backgroundColor: selectedColumn === colIndex ? '#d3d3d3' : 'inherit'
                    }}>
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
        <button className="button" onClick={handleNewSheet}>New Sheet</button>
        <span>Sheet Name: {sheetNames[currentSheet]}</span>
        <div>
          {sheetNames.map((name, idx) => (
            <button key={idx} onClick={() => changeSheet(idx)}>{name}</button>
          ))}
        </div>
      </div>

      {(selectedRow !== null || selectedColumn !== null) && (
        <div className="dropdown" style={{ top: menuPosition.top, left: menuPosition.left, position: 'absolute', backgroundColor: '#ffffff', border: '1px solid #ddd', padding: '10px' }}>
          {selectedRow !== null && (
            <>
              <button onClick={addRowAbove}>Add Row Above</button>
              <br />
              <button onClick={addRowBelow}>Add Row Below</button>
              <br />
              <button onClick={deleteRow}>Delete Row</button>
              <br />
              <button onClick={() => sortData('row')}>Sort Row</button>
            </>
          )}
          {selectedColumn !== null && (
            <>
              <button onClick={addColumnLeft}>Add Column Left</button>
              <br />
              <button onClick={addColumnRight}>Add Column Right</button>
              <br />
              <button onClick={deleteColumn}>Delete Column</button>
              <br />
              <button onClick={() => sortData('column')}>Sort Column</button>
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default App;
