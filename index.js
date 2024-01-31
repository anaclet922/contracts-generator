const express = require('express')
const compression = require("compression");
var DocxMerger = require('docx-merger');

const app = express()
const port = 3000


app.use(compression()); // Compress all routes
//to use public files; css, js, images...inside public folder
app.use(express.static('public'));
// app.use(express.static('contracts'));
app.set('view engine', 'ejs');

// app.use(bodyParser.json()) // for parsing application/json
// app.use(bodyParser.urlencoded({ extended: true })) // for parsing application/x-www-form-urlencoded
app.use(express.json());
app.use(express.urlencoded({ extended: true }));


const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");


const fs = require("fs");
const path = require("path");


app.use(express.static(path.join(__dirname, 'contracts')));
app.use(express.static(path.join(__dirname, 'proposals')));
app.use(express.static(path.join(__dirname, 'tmp')));

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

    let policy = req.body.police;
    let type = req.body.type;
    let lang = req.body.lang;

    let all_body = req.body;

    // Load the docx file as binary content
    let content = fs.readFileSync(
        path.resolve(__dirname, "docx-templates/template-cow-" + lang + ".docx"),
        "binary"
    );

    if (type == 'crop') {
        content = fs.readFileSync(
            path.resolve(__dirname, "docx-templates/template-crop-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'chicken') {
        content = fs.readFileSync(
            path.resolve(__dirname, "docx-templates/template-chicken-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'pig') {
        content = fs.readFileSync(
            path.resolve(__dirname, "docx-templates/template-pig-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'cow') {
        content = fs.readFileSync(
            path.resolve(__dirname, "docx-templates/template-cow-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'student') {
        content = fs.readFileSync(
            path.resolve(__dirname, "docx-templates/template-student-" + lang + ".docx"),
            "binary"
        );
    } else if (type == 'accident') {
        content = fs.readFileSync(
            path.resolve(__dirname, "/docx-templates/template-accident-" + lang + ".docx"),
            "binary"
        );
    }


    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
    doc.render(all_body);

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
        filename: filename_pdf + ext
    }
    res.status(200).json(output);

});

app.post('/generate-proposal', async function (req, res) {


    let nsin = req.body.nsin;
    let typpol = req.body.typpol;
    let all_body = req.body;

    // Load the docx file as binary content
    let content = fs.readFileSync(
        path.resolve(__dirname, "docx-templates/template-payment-proposal-" + typpol + ".docx"),
        "binary"
    );



    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
    doc.render(all_body);

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    // buf is a nodejs Buffer, you can either write it to a
    // file or res.send it with express for example.
    let filename = nsin + "-output.docx";
    fs.writeFileSync(path.resolve(__dirname, "proposals/" + filename), buf);




    let output = {
        filename: filename
    }
    res.status(200).json(output);

});

app.get('/download/:filename', (req, res) => {

    const filePath = __dirname + '/contracts/' + req.params.filename;

    res.download(
        filePath,
        "downloaded-" + req.params.filename,
        (err) => {
            if (err) {
                res.send({
                    error: err,
                    msg: "Problem downloading the file"
                })
            }
        }
    );

});

app.get('/download-proposal/:filename', (req, res) => {

    const filePath = __dirname + '/proposals/' + req.params.filename;

    res.download(
        filePath,
        "downloaded-" + req.params.filename,
        (err) => {
            if (err) {
                res.send({
                    error: err,
                    msg: "Problem downloading the file"
                })
            }
        }
    );

});

app.get('/delete-old-files', (req, res) => {

    const directory = "contracts";
    const tmp = "tmp";

    fs.readdir(directory, (err, files) => {
        if (err) throw err;

        for (const file of files) {
            fs.unlink(path.join(directory, file), (err) => {
                if (err) throw err;
            });
        }
    });

    fs.readdir(tmp, (err, files) => {
        if (err) throw err;

        for (const file of files) {
            fs.unlink(path.join(tmp, file), (err) => {
                if (err) throw err;
            });
        }
    });

    let output = {
        response: 'cleaning'
    }
    res.status(200).json(output);

});



app.post('/generate-combined-proposal', async function (req, res) {

    try {

        let folder_name = req.body.folder_name;
        let files = req.body.files;

        console.log(files);

        // Array to hold file buffers
        let fileBuffers = [];

        // Read each file asynchronously
        for (let file of files) {
            let filePath = path.resolve(__dirname, "proposals/" + file);
            let fileBuffer = fs.readFileSync(filePath, 'binary');
            fileBuffers.push(fileBuffer);
        }

        var docx = new DocxMerger({}, fileBuffers);

        let filename = folder_name + '.docx';

        //SAVING THE DOCX FILE
        docx.save('nodebuffer', function (data) {
            fs.writeFile(path.resolve(__dirname, "tmp/" + filename), data, function (err) {
                if (err) {
                    console.error("Error saving file:", err);
                    return res.status(500).json({ error: "Error saving file" });
                }
                let output = {
                    filename: filename
                }
                return res.status(200).json(output);
            });
        });


    } catch (error) {

        console.error("Error generating combined proposal:", error);
        return res.status(500).json({ error: "Error generating combined proposal" });

    }
});


app.get('/download-combined-proposal/:filename', (req, res) => {

    const filePath = __dirname + '/tmp/' + req.params.filename;

    res.download(
        filePath,
        "downloaded-" + req.params.filename,
        (err) => {
            if (err) {
                res.send({
                    error: err,
                    msg: "Problem downloading the file"
                })
            }
        }
    );

});


app.listen(port, () => { console.log('App Started at port ' + port) });