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
    const filename = handleDocumentCreation(req.query)
    res.download(filename);
});

app.listen(port, () => {
    console.log(`Serwer odpalony pod http://localhost:${port}`)
});

const mapDate = (date) => {
    const [year, month, day] = date.split('-');

    return `${day}/${month}/${year}`
}

const extractUserData = (copiedText) => {
    const regex = new RegExp(/(?<address>.*)NIP: (?<nip>\w*) (?<name>.*)[(].*tel.*[)] ?(?<phonenumber>\d*)/);
    const noPhoneWorkaroundRegex = new RegExp(/(?<address>.*)NIP: (?<nip>\w*) (?<name>.*)[(].*tel.*[)]? ?(?<phonenumber>\d*)/);

    const result = copiedText.match(regex) ? copiedText.match(regex) : copiedText.match(noPhoneWorkaroundRegex);

    return {
        address: result.groups.address.trim(),
        nip: result.groups.nip,
        name: result.groups.name.trim(),
        phoneNumber: result.groups.phonenumber
    }
}

const createOrderNumber = (orderNumberId, orderNumberMonth, orderNumberInitials) => `${orderNumberId}${orderNumberMonth}${orderNumberInitials}`;

const generateRate = (netRate, netRateCurrency) => `${netRate} ${netRateCurrency === 'EUR' ? 'euro all in' : 'PLN all in'}`;

const handleDocumentCreation = (requestQuery) => {
    const { input, loadDate, unloadDate, orderNumberId, orderNumberMonth, orderNumberInitials, registrationPlate, typeOfGoods, weight, netRate, netRateCurrency, payDays } = requestQuery;

    const TEMPLATE_NAME = 'template.docx';
    const GENERATED_FOLDER_NAME = 'generated';

    const createDocumentsDirectoryIfNeeded = () => {
        const generatedDocumentsDirectoryPath = path.resolve(__dirname, GENERATED_FOLDER_NAME);
        if (!fs.existsSync(generatedDocumentsDirectoryPath)) {
            fs.mkdirSync(generatedDocumentsDirectoryPath);
        }
    }

    const generateDocument = (data) => {
        const templatePath = path.resolve(__dirname, TEMPLATE_NAME);
        const content = fs.readFileSync(templatePath, 'binary');

        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        doc.render(data); // this changes placeholders with actual data

        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });


        createDocumentsDirectoryIfNeeded();
        const fileName = path.resolve(__dirname, GENERATED_FOLDER_NAME, `${data.user_name.replace(' ', '-')}.docx`);
        fs.writeFileSync(fileName, buf);

        return fileName;
    }

    const text = input.replace(/\n/g, ' ').replace(/\r/g, ' ');
    let fileName;
    try {
        const user = extractUserData(text);
        const data = {
            user_address: user.address,
            user_nip: user.nip,
            user_name: user.name,
            user_phoneNumber: user.phoneNumber,
            order_loadDate: mapDate(loadDate),
            order_unloadDate: mapDate(unloadDate),
            order_orderNumber: createOrderNumber(orderNumberId, orderNumberMonth, orderNumberInitials),
            order_registrationPlate: registrationPlate,
            order_typeOfGoods: typeOfGoods.replaceAll('\r', ''),
            order_weight: weight,
            order_rate: generateRate(netRate, netRateCurrency),
            order_payDays: `${payDays} dni`,
        };

        fileName = generateDocument(data);

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