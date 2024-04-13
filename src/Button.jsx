import React, { useState } from "react";
import ExcelJS from "exceljs";
import moment from "moment";

const Button = () => {
  const [excelSheetData, setExcelSheetData] = useState([
    {
      createAt: new Date(),
      donorId: "001",
      donorDetails: {
        full_name: "John Doe",
        email: "john@example.com",
        address: "123 Main St",
      },
      amount: 100,
    },
    {
      createAt: new Date(),
      donorId: "002",
      donorDetails: {
        full_name: "Jane Smith",
        email: "jane@example.com",
        address: "456 Elm St",
      },
      amount: 150,
    },
    {
      createAt: new Date(),
      donorId: "003",
      donorDetails: {
        full_name: "Jane Smith",
        email: "jane@example.com",
        address: "456 Elm St",
      },
      amount: 2000,
    },
    // Add more sample data as needed
  ]);

  const exportExcelFile = () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Donation Report");
    sheet.properties.defaultRowHeight = 20;

    // Add title in the first column (A)
    sheet.getCell("A1").value = "Church Management System";

    // Define other headers starting from the second column (B)

    sheet.columns = [
      {
        header: "Donation Date",
        key: "donationDate",
        width: 20,
      },
      {
        header: "Donor Id",
        key: "donorId",
        width: 20,
      },
      {
        header: "Email",
        key: "email",
        width: 30,
      },
      {
        header: "Total",
        key: "total",
        width: 20,
      },
      {
        header: "Amount",
        key: "amount",
        width: 20,
      },
      {
        header: "Full Name",
        key: "full_name",
        width: 20,
      },
      {
        header: "Address",
        key: "address",
        width: 20,
      },
    ];

    // Format column headers
    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).alignment = { horizontal: "center" };

    let totalAmount = 0;

    excelSheetData?.forEach((item, index) => {
      const rowNumber = index + 2; // Start from the second row
      const donationDate = moment(item?.createAt).format("LL");
      const amount = item?.amount;

      totalAmount += amount; // Calculate total amount

      sheet.addRow({
        donationDate,
        donorId: item?.donorId,
        full_name: item?.donorDetails?.full_name,
        email: item?.donorDetails?.email,
        address: item?.donorDetails?.address,
        amount,
        total: "", // Leave this blank for now
      });
    });

    // Add total amount in the last row under the "Total" column
    const lastRowNumber = excelSheetData.length + 2;
    sheet.getCell(`D${lastRowNumber}`).value = totalAmount;

    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheet.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "download.xlsx";
      anchor.click();
      window.URL.revokeObjectURL(url);
    });
  };

  return (
    <div>
      <button onClick={exportExcelFile}>Export</button>
    </div>
  );
};

export default Button;
