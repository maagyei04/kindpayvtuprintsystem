import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

const Excel = () => {
    const [currentPage, setCurrentPage] = useState(1);
    const [pageSize] = useState(10000);
    const [totalPages, setTotalPages] = useState(0);
    const [cardData, setCardData] = useState([]);
    const [image, setImage] = useState(null);
    const [excelError, setExcelError] = useState(null);
    const [imageError, setImageError] = useState(null);
    const [currentPageData, setCurrentPageData] = useState([]);

    useEffect(() => {
        // Calculate total pages
        const totalPages = Math.ceil(cardData.length / pageSize);
        setTotalPages(totalPages);

        // Calculate start and end indices for the current page
        const startIndex = (currentPage - 1) * pageSize;
        const endIndex = Math.min(startIndex + pageSize, cardData.length);
        setCurrentPageData(cardData.slice(startIndex, endIndex));
    }, [cardData, currentPage, pageSize]);

    const handlePrint = () => {
        window.print();
    };

    const handleFileUpload = (event) => {
        const file = event.target.files[0];
        const reader = new FileReader();

        // Reset error messages
        setExcelError(null);

        // Validate file type
        if (!file || !file.name.endsWith('.xlsx')) {
            setExcelError('Please upload a valid Excel file (xlsx format).');
            return;
        }

        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            // Ensure header row exists
            if (jsonData.length < 1) {
                setExcelError('No data found in the Excel sheet.');
                return;
            }

            // Find the index of the 'pin' column
            const headerRow = jsonData[0];
            const pinIndex = headerRow.indexOf('digit');
            const amountIndex = headerRow.indexOf('amount');
            const serialIndex = headerRow.indexOf('serialNumber');

            // Ensure 'pin' column exists
            if (pinIndex === -1 || amountIndex === -1 || serialIndex === -1) {
                setExcelError('Columns "serial", "pin" and "amount" not found in the Excel sheet.');
                return;
            }

            const cardData = jsonData.slice(1).map((row) => ({
                pin: row[pinIndex],
                amount: row[amountIndex],
                serial: row[serialIndex],
            }));

            // Calculate total pages
            const totalPages = Math.ceil(cardData.length / pageSize);
            setTotalPages(totalPages);

            setCardData(cardData);
            setCurrentPage(1);
        };

        reader.readAsArrayBuffer(file);
    };

    const nextPage = () => {
        if (currentPage < totalPages) {
            setCurrentPage(currentPage + 1);
        }
    };

    const prevPage = () => {
        if (currentPage > 1) {
            setCurrentPage(currentPage - 1);
        }
    };

    const handleImageUpload = (event) => {
        const file = event.target.files[0];
        const reader = new FileReader();

        // Reset error messages
        setImageError(null);

        // Validate file type
        if (!file || !file.type.startsWith('image/')) {
            setImageError('Please upload a valid image file.');
            return;
        }

        reader.onload = (e) => {
            setImage(e.target.result);
        };

        reader.readAsDataURL(file);
    };

    return (
        <div className="container mt-[90px]">
            <h1 className='font-bold mb-5'>Scratch Cards / VTU Pins</h1>
            {excelError && <p className="error">{excelError}</p>}
            <input type="file" accept=".xlsx" onChange={handleFileUpload} />
            {imageError && <p className="error">{imageError}</p>}
            <label htmlFor="image-upload" className="custom-file-upload items-center flex flex-col justify-evenly">
                {image ? <div className="image-preview-container" style={{ position: 'relative' }}>
                    <img src={image} height={100} width={100} alt="Uploaded" />
                </div>
                    : "Upload Image"}
                <input id="image-upload" type="file" accept="image/*" onChange={handleImageUpload} />
            </label>
            <button onClick={handlePrint} className='btn bg-violet-600 text-white md:ml-4 font-semibold px-3 py-2 rounded-[10px] duration-500'>Print</button>
            <div className="card-container print-content">
                {currentPageData.map((card, index) => (
                    <div className="card" key={index}>
                        <div className="image-container image-preview-container" style={{ position: 'relative' }}>
                            <img src={image} alt="Scratch Card" />
                            <div className="image-text font-bold text-[12px] mr-[1px] mt-o">{card.amount}</div>
                            <div className="serial-text font-semibold text-[6px] mr-1">{card.serial}</div>
                        </div>
                        <div className="pin">
                            <p className='font-semibold text-sm'>*920*200*{card.pin}#</p>
                        </div>
                    </div>
                ))}
            </div>
            <div className="pagination">
                <button className='btn bg-violet-600 text-white md:ml-4 font-semibold px-3 py-2 rounded-[10px] duration-500' onClick={prevPage} disabled={currentPage === 1}>Previous</button>
                <span>{currentPage} of {totalPages}</span>
                <button className='btn bg-violet-600 text-white md:ml-4 font-semibold px-3 py-2 rounded-[10px] duration-500' onClick={nextPage} disabled={currentPage === totalPages}>Next</button>
            </div>
        </div>
    );
}

export default Excel;
