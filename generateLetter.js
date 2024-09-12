import fs from "fs";
import path from 'path';
import { fileURLToPath } from 'url';
import XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Read the excel file
const workbook = XLSX.readFile(path.join(__dirname, "/docs/excel/admission_list.xlsx"));

// Access the particular sheet in the file
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Convert sheet to array of data
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Read the document template
const template = fs.readFileSync(path.join(__dirname, '/docs/template/template.docx'), 'binary');

// Ensure output directory exists
const outputDir = path.join(__dirname, '/docs/output');
if (!fs.existsSync(outputDir)) {
	fs.mkdirSync(outputDir, { recursive: true });
}

// Function to create personalized document
function createPersonalizedDocument(placeholders) {
	const zip = new PizZip(template);
	const doc = new Docxtemplater(zip, {
		paragraphLoop: true,
		linebreaks: true,
	});

	// Render the document (replace placeholders)
	doc.render(placeholders);

	const buf = doc.getZip().generate({
		type: 'nodebuffer',
		compression: "DEFLATE"
	});

	return buf;
}

// Generate a letter for each applicant
data.forEach((row, index) => {
	if (index === 0) return; // Skip header row
	const firstName = row[1]; // Column B (index 1) for First Name
	const surname = row[2];   // Column C (index 2) for Surname
	// Skip rows with missing names
	if (!firstName || !surname) return;
	// Combine the first name and surname into a full name
	const fullName = `${firstName} ${surname}`;

	// Create a personalized letter
	const personalizedLetter = createPersonalizedDocument({
		name: fullName,
		date: new Date().toLocaleDateString('en-US', {
			year: 'numeric',
			month: 'numeric',
			day: 'numeric'
		})
	});

	// Create a file name using the applicant's name, replacing spaces with underscores
	const fileName = `admission_letter_${fullName.replace(/ /g, '_')}.docx`;

	// Save the personalized letter to a file named after the applicant
	fs.writeFileSync(path.join(__dirname, `/docs/output/${fileName}`), personalizedLetter);
	console.log(`Generated letter for ${fullName} with file name: ${fileName}`);
});
