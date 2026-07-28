const templateMapping = {
  version: 1,
  templateFile: "./templates/Travelling claim form - updated version.xlsx",
  templateSha256: "a3d78782754588db9223c0e751d8e892f9fc8cc282fa1cc8027c9d2f022e23e9",
  workbook: {
    sheets: ["Sheet1", "Sheet2", "Sheet3"],
    usedRanges: {
      Sheet1: "A1:N79",
      Sheet2: "A1:A1",
      Sheet3: "A1:A1"
    },
    sheetXml: {
      Sheet1: "xl/worksheets/sheet1.xml"
    },
    stylesXml: "xl/styles.xml"
  },
  fields: {
    reportPeriod: "B1",
    employeeName: "I1",
    employeeNumber: "I1",
    address: "B2",
    claimDate: "B3",
    weekEnding: "B1",
    remarks: "B3",
    grandTotalPage1: "N29",
    grandTotalPage2: "N58"
  },
  tripSections: [
    { name: "page1", startRow: 7, endRow: 27, totalCell: "N29" },
    { name: "page2", startRow: 34, endRow: 56, totalCell: "N58" }
  ],
  tripColumns: {
    date: "B",
    fromTo: "C",
    from: null,
    to: null,
    timeFrom: "E",
    timeTo: "F",
    client: "G",
    projectCode: "H",
    mtr: "I",
    bus: "J",
    van: "K",
    taxi: "L",
    other: null,
    subTotal: "M",
    deduction: "N",
    remarks: null
  },
  transportColumns: {
    MTR: "I",
    Bus: "J",
    Van: "K",
    Taxi: "L",
    Other: null
  },
  maxTrips: 44,
  mergedCells: {
    Sheet1: [
      "C34:D34", "B2:M2", "C24:D24", "I5:L5", "C15:D15", "C5:D5", "C42:D42", "E5:F5",
      "C14:D14", "C45:D45", "C26:D26", "C35:D35", "C20:D20", "C10:D10", "C16:D16",
      "C54:D54", "C47:D47", "C25:D25", "C41:D41", "C55:D55", "C50:D50", "C44:D44",
      "C22:D22", "C40:D40", "C9:D9", "B3:N3", "C56:D56", "G5:G6", "C12:D12", "C43:D43",
      "C46:D46", "C21:D21", "C11:D11", "H32:H33", "C51:D51", "C27:D27", "C36:D36",
      "I32:L32", "C52:D52", "C48:D48", "C17:D17", "C23:D23", "G32:G33", "C8:D8",
      "C39:D39", "H5:H6", "C32:D32", "C7:D7", "C38:D38", "E32:F32", "C19:D19",
      "C37:D37", "C13:D13", "C18:D18", "C53:D53", "C49:D49"
    ],
    Sheet2: [],
    Sheet3: []
  },
  formulaCells: {
    Sheet1: [],
    Sheet2: [],
    Sheet3: []
  },
  limitations: [
    "Template has no separate From and To cells because each trip row merges C:D; export writes a combined 'From → To' value into C.",
    "Template has no Other transport column; Other trips are recorded in Sub Total and annotated in the location text.",
    "Template has no row-level Remarks cell; trip remarks are kept in the app data and preview/history only.",
    "Template has no Employee Number, Week Ending, Claim Date or free Remarks cells; these are written into existing heading/address note cells without adding rows or columns.",
    "No worksheet image/logo object was detected in the supplied workbook. Existing package parts are still preserved."
  ]
};

export default templateMapping;
