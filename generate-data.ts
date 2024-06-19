import * as fs from 'fs';
import * as XLSX from 'xlsx';
import { faker } from '@faker-js/faker';

// Define the Column interface
interface Column {
  columnName: string;
  type: string;
  length: number;
}

const NUM_RECORDS = 50;  // Number of records to generate

// Function to generate dummy data
const generateDummyData = (columns: Column[]): any[] => {
  const data = [];
  for (let i = 0; i < NUM_RECORDS; i++) {
    const record: any = {};
    const sex = faker.person.sexType();
    const first_name = faker.person.firstName(sex);
    const last_name = faker.person.lastName();
    const email = faker.internet.email({ firstName: first_name, lastName: last_name });
    columns.forEach(col => {
      switch (col.columnName) {
        case 'first_name':
          record[col.columnName] = first_name;
          break;
        case 'last_name':
          record[col.columnName] = last_name;
          break;
        case 'email':
          record[col.columnName] = email;
          break;
        case 'id':
          // Generate a string id with the format 'AB-1234'
          record[col.columnName] = first_name.charAt(0) + last_name.charAt(0) + faker.string.numeric(3);
          break;
        case 'phone':
          record[col.columnName] = faker.phone.number();
          break;
        case 'address':
          record[col.columnName] = faker.address.streetAddress();
          break;
        case 'sex':
          record[col.columnName] = faker.person.sex();
          break;
        case 'date_of_birth':
          record[col.columnName] = faker.date.birthdate();
          break;
        default:
          switch (col.type) {
            case 'number':
              record[col.columnName] = faker.number.int(100);
              break;
            case 'string':
              record[col.columnName] = faker.string.alphanumeric(col.length);
              break;
            default:
              record[col.columnName] = '';
          }
      }
    });
    data.push(record);
  }
  return data;
};

// Read and parse columns from JSON file
let columns: Column[];
try {
  const columnsData = fs.readFileSync('columns.json', 'utf-8');
  columns = JSON.parse(columnsData);
} catch (error) {
  console.error('Error reading or parsing columns.json:', error);
  process.exit(1);
}

// Generate data
const data = generateDummyData(columns);

// Create a new workbook and worksheet
const ws = XLSX.utils.json_to_sheet(data);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

// Write to Excel file
try {
  XLSX.writeFile(wb, 'output.xlsx');
  console.log('Excel file generated successfully!');
} catch (error) {
  console.error('Error writing Excel file:', error);
  process.exit(1);
}