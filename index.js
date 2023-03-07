const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const path = require('path');
const fs = require('fs');
const express = require('express')

const app = express()
const port = 3000

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '/index.html'));
});

app.get('/generate', (req, res) => {
    const { input, loadDate, unloadDate } = req.query;

    const filename = handleDocumentCreation(input, loadDate, unloadDate)
    res.download(filename);
});

app.listen(port, () => {
    console.log(`Serwer odpalony pod http://localhost:${port}`)
});

const mapDate = (date) => {
    const [year, month, day] = date.split('-');

    return `${day}/${month}/${year}`
}

const handleDocumentCreation = (input, loadDate, unloadDate) => {
    const extractUserData = (copiedText) => {
        const regex = new RegExp(/(?<address>.*)NIP: (?<nip>\w*) (?<name>.*)[(].*tel.*[)] ?(?<phonenumber>\d*)/);
        const noPhoneWorkaroundRegex = new RegExp(/(?<address>.*)NIP: (?<nip>\w*) (?<name>.*)[(].*tel.*[)]? ?(?<phonenumber>\d*)/);

        const result = copiedText.match(regex) ? copiedText.match(regex) : copiedText.match(noPhoneWorkaroundRegex);

        return {
            address: result.groups.address.trim(),
            nip: result.groups.nip,
            name: result.groups.name.trim(),
            phoneNumber: result.groups.phonenumber,
            loadDate: mapDate(loadDate),
            unloadDate: mapDate(unloadDate),
        }
    }

    const TEMPLATE_NAME = 'template.docx';
    const GENERATED_FOLDER_NAME = 'generated';

    const createDocumentsDirectoryIfNeeded = () => {
        const generatedDocumentsDirectoryPath = path.resolve(__dirname, GENERATED_FOLDER_NAME);
        if (!fs.existsSync(generatedDocumentsDirectoryPath)) {
            fs.mkdirSync(generatedDocumentsDirectoryPath);
        }
    }

    const generateDocument = (user) => {
        const templatePath = path.resolve(__dirname, TEMPLATE_NAME);
        const content = fs.readFileSync(templatePath, 'binary');

        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        doc.render(user); // this changes placeholders with actual data

        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });


        createDocumentsDirectoryIfNeeded();
        const fileName = path.resolve(__dirname, GENERATED_FOLDER_NAME, `${user.name.replace(' ', '-')}.docx`);
        fs.writeFileSync(fileName, buf);

        return fileName;
    }

    const text = input.replace(/\n/g, ' ').replace(/\r/g, ' ');
    let fileName;
    try {
        const user = extractUserData(text);
        fileName = generateDocument(user);
        console.log('DONE');
        console.log('plik utworzony w ', fileName);
    } catch (e) {
        console.error("Wywalilo się :(");
        console.error("Jeśli dla tego co wkleiles powinno dzialać to wyślij mi proszę to co wkleiles i poniższy log");
        console.error("--------------------------------------")
        console.error(e)
        console.error("--------------------------------------")
    }

    return fileName;
}