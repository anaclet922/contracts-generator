const express = require('express')
const app = express()
const port = 3000

//to use public files; css, js, images...inside public folder
app.use(express.static('public'))
app.set('view engine', 'ejs');

// app.use(bodyParser.json()) // for parsing application/json
// app.use(bodyParser.urlencoded({ extended: true })) // for parsing application/x-www-form-urlencoded
app.use(express.json());
app.use(express.urlencoded({ extended: true }));


const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");


const fs = require("fs");
const path = require("path");



const fss = require('fs').promises;
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);


app.get('/', (req, res) => {
    console.log('Index called!');
    res.render('index')
});

app.get('/test', (req, res) => {
    res.status(200).json({ message: "API test endpoint reached" });
});
app.post('/generate-contract', async function (req, res) {

    let names = req.body.names;
    let nid = req.body.nid;
    let policy = req.body.policy;


    let type = req.body.type;
    let lang = req.body.lang;


    // Load the docx file as binary content
    let content = fs.readFileSync(
        path.resolve(__dirname, "docx-templates/template-cow-" + lang + ".docx"),
        "binary"
    );

    if (type == 'porc') {
        content = fs.readFileSync(
            path.resolve(__dirname, "/docx-templates/template-porc-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'crop') {
        content = fs.readFileSync(
            path.resolve(__dirname, "/docx-templates/template-porc-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'accident') {
        content = fs.readFileSync(
            path.resolve(__dirname, "/docx-templates/template-accident-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'school') {
        content = fs.readFileSync(
            path.resolve(__dirname, "/docx-templates/template-school-" + lang + ".docx"),
            "binary"
        );
    }


    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
    doc.render({
        names: names,
        nid: nid,
        policy: policy

    });

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    // buf is a nodejs Buffer, you can either write it to a
    // file or res.send it with express for example.
    let filename = policy + "-output.docx";
    let filename_pdf = policy + "-output";
    fs.writeFileSync(path.resolve(__dirname, "contracts/" + filename), buf);


    const ext = '.pdf'
    const inputPath = path.join(__dirname, '/contracts/' + filename);
    const outputPath = path.join(__dirname, `/contracts/${filename_pdf}${ext}`);
    // Read file
    const docxBuf = await fss.readFile(inputPath);
    // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
    let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);
    // Here in done you have pdf file which you can save or transfer in another stream
    await fss.writeFile(outputPath, pdfBuf);


    fs.unlink("contracts/" + filename, err => {
        if (err) {
            console.log(`An error occurred ${err.message}`);
        } else {
            console.log(`Deleted the file under ${path}`);
        }
    });


    let output = {
        filename: filename_pdf
    }
    res.status(200).json(output);

});

app.listen(port, () => { console.log('App Started at port ' + port) });