import { generateExcel } from "./generate-excel";

let sampleData = [
  [
    { readOnly: true, value: "" },
    { value: "A", readOnly: true },
    { value: "B", readOnly: true },
    { value: "C", readOnly: true },
    { value: "D", readOnly: true },
    { value: "D", readOnly: true },
  ],
  [
    { readOnly: true, value: 1 },
    { value: 1 },
    { value: 3 },
    { value: 3 },
    { value: 3 },
  ],
  [
    { readOnly: true, value: 2 },
    { value: 2 },
    { value: 4 },
    { value: 4 },
    { value: 4 },
  ],
  [
    { readOnly: true, value: 3 },
    { value: 1 },
    { value: 3 },
    { value: 3 },
    { value: 3 },
  ],
  [
    { readOnly: true, value: 4 },
    { value: 2 },
    { value: 4 },
    { value: 4 },
    { value: 4 },
  ],
];

// generateExcel(sampleData);

export { generateExcel, sampleData };
