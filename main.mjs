import express from 'express';
import cors from 'cors';
import 'dotenv/config';
const PORT = process.env.PORT || 3000;

const app = express();
const port = PORT;

// Middleware to parse JSON bodies
app.use(express.json());
app.use(cors({
    origin: '*',
}));

// Route to receive and return JSON
app.post('/', async (req, res) => {
    const data = req.body;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    // Set headers dynamically from the first row of the input matrix
    const headers = data.shift(); // Remove and retrieve the first row as headers
    if (!headers) throw new Error('No headers found');

    headers.forEach((header, index) => {
        let row = worksheet.getRow(1);
        // row.height = 40;

        let cell = row.getCell(index + 1);
        cell.value = header;
        cell.font = { bold: true };
    });

    data.forEach((row, rowIndex) => {
        row.forEach((cellValue, cellIndex) => {
            worksheet.getRow(rowIndex + 2).height = 40;

            if (isBase64EncodedImage(cellValue)) {

                const imageBuffer = base64ToBuffer((cellValue).replace(/^data:image\/\w+;base64,/, ''));
                const imageFileName = `image_${rowIndex}_${cellIndex}.png`; // Create a unique filename

                const imageFilePath = `/tmp/${imageFileName}`;
                fs.writeFileSync(imageFilePath, imageBuffer);

                const imageId = workbook.addImage({
                    filename: imageFilePath,
                    extension: 'png',
                });

                worksheet.addImage(imageId, {
                    // @ts-ignore tslint:disable-line
                    tl: { col: cellIndex, row: rowIndex + 1 },
                    // @ts-ignore tslint:disable-line
                    br: { col: cellIndex + 1, row: rowIndex + 2 },
                });

            } else {
                worksheet.getRow(rowIndex + 2).getCell(cellIndex + 1).value = cellValue;
            }

        });
    });

    const buffer = await workbook.xlsx.writeBuffer();


    res.send(buffer);
});

app.get('/', (_, res) => {
    res.json({ status: 'OK' });
})

app.get('*', (_, res) => {
    res.status(404).send('Not found');
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});

