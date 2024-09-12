
import fs from "fs";
import path from "path";
import { exec } from "child_process";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Function to convert .docx to PDF using LibreOffice
function convertDocxToPdf(inputPath, outputPath) {
	const command = `libreoffice --headless --convert-to pdf --outdir ${path.dirname(outputPath)} ${inputPath}`;

	exec(command, (err, stdout, stderr) => {
		if (err) {
			console.error(`Failed to convert ${inputPath}:`, err);
			return;
		}
		console.log(`PDF created successfully at ${outputPath}`);
	});
}

// Function to process all .docx files in a directory
function processDocxFiles(inputDir) {
	fs.readdir(inputDir, (err, files) => {
		if (err) {
			console.error('Error reading directory:', err);
			return;
		}

		files.forEach((file) => {
			const fileNameWithoutExt = path.basename(file, '.docx');
			const inputFilePath = path.join(inputDir, file);
			const outputFilePath = path.join(__dirname, `/docs/output/pdf/${fileNameWithoutExt}.pdf`);

			convertDocxToPdf(inputFilePath, outputFilePath);
		});
	});
}

// Define your input directory
const inputDir = './docs/output/docx/'; // Replace with the path to your folder

// Process all .docx files in the specified directory
processDocxFiles(inputDir);

