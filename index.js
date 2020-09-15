const path = require("path");
const express = require('express')
const bodyParser = require('body-parser')
const exphbs = require('express-handlebars')
const xl = require('excel4node');
const fs = require('fs');
const nodemailer = require('nodemailer')
const passport = require('passport')
const flash = require('express-flash')
const session = require('express-session')
const POP3Strategy = require('passport-pop3')
const methodOverride = require('method-override')
const cookieParser = require('cookie-parser');
const multer = require('multer')
const upload = multer({ dest: __dirname + "/temps" })


let userr = ''
let passs = ''
let pop = new POP3Strategy({
    host: 'pop.secureserver.net',
    port: 995,
    enabletls: true,
    usernameField: userr,
    passwordField: passs,
}
)
passport.serializeUser(function (user, done) {
    done(null, user);
});

passport.deserializeUser(function (user, done) {
    done(null, user);
});
const app = express()
app.use(flash())
app.use(session({
    secret: 'keyboard cat',
    resave: false,
    saveUninitialized: false
}))
app.use(passport.initialize())
app.use(passport.session())
passport.use(pop);
app.use(methodOverride('_method'))
app.use(cookieParser());

//Settings
app.set('port', process.env.PORT || 3000)
//BODY-PARSER
app.use(bodyParser.urlencoded({ extended: true }))
app.use(express.json());
app.use(express.static(path.join(__dirname + '/views')));
app.use(express.static(path.join(__dirname + '/public')));

app.use(function (req, res, next) {
    req.flash('message', 'Enviado');
    next();
});
//HANDLEBARS
app.engine('.hbs', exphbs({ extname: '.hbs', defaultLayout: 'main.hbs' }));
app.set('view engine', '.hbs');

app.get('/', checkAuthenticated, function (req, res) {
    res.render(`inicio`, {
        style: 'home.css'
    });
});
app.get('/login', checkNotAuthenticated, function (req, res) {
    res.render(`login`, {
        style: 'login.css'
    });
});

app.get('/fisicas', checkAuthenticated, function (req, res) {
    res.render(`fisicas`, {
        style: 'formulario.css'
    });
});
app.get('/soloDatos', checkAuthenticated, function (req, res) {
    res.render(`soloDatos`, {
        style: 'formulario.css'
    });
});

app.get('/juridicas', checkAuthenticated, function (req, res) {
    res.render(`juridicas`, {
        style: 'formulario.css'
    });
});

app.get('/redir_login', checkAuthenticated, function (req, res) {
    res.render(`redir_login`);
});

app.post('/', passport.authenticate('pop3', { failureRedirect: '/' }),
    function (req, res) {
        userr = req.body.username
        passs = req.body.password
        res.render('inicio', { mensajeBienvenida: `BIENVENIDO: ${userr}`, style: 'home.css' });
    });



var cpUpload = upload.fields([{ name: 'constancia', maxCount: 1 }, { name: 'estatuto' }, { name: 'ultimobalance', maxCount: 1 }, { name: 'dnifrente', maxCount: 10 }, { name: 'dnidorso', maxCount: 10 }])
app.post('/juridicas', cpUpload, function (req, res) {


    // Create a new instance of a Workbook class
    var wb = new xl.Workbook();

    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet 1');

    // Create a reusable style
    var style = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleShare = wb.createStyle({
        font: {
            color: 'white',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#0008a3',
            fgColor: '#0008a3',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleHist = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#f50000',
            fgColor: '#f50000',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleTITULOS = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#FFFF00',
            fgColor: '#FFFF00',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleWeb = wb.createStyle({
        font: {
            color: 'white',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#007a00',
            fgColor: '#007a00',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });

    /* OBTENER FECHA */

    let today = new Date();
    let year = today.getFullYear();
    let month = today.getMonth() + 1;
    let day = today.getDate();
    let fechaCompleta = day + '/' + month + '/' + year;

    /* CELDAS HISTORIAL */

    // Set value of cell A1 to 100 as a number type styled with paramaters of style

    ws.cell(1, 5)
        .string('FILA HISTORIAL')
        .style(styleHist);

    ws.cell(2, 2)
        .string('FECHA CARGA SOLICITUD')
        .style(styleTITULOS);
    ws.cell(2, 3)
        .string('CUIT')
        .style(styleTITULOS);
    ws.cell(2, 4)
        .string('RAZON SOCIAL')
        .style(styleTITULOS);
    ws.cell(2, 5)
        .string('PROMOTOR RAILCOM')
        .style(styleTITULOS);
    ws.cell(2, 6)
        .string('PRODUCTO OFRECIDO')
        .style(styleTITULOS);
    ws.cell(2, 7)
        .string('N° REGISTRO STADER')
        .style(styleTITULOS);
    ws.cell(2, 8)
        .string('ESTADO (A - R - P - S)')
        .style(styleTITULOS);
    ws.cell(2, 9)
        .string('ENTREGADA A PROMOTOR')
        .style(styleTITULOS);
    ws.cell(2, 10)
        .string('N° SUC')
        .style(styleTITULOS);
    ws.cell(2, 11)
        .string('N° CTA CTE')
        .style(styleTITULOS);
    ws.cell(2, 12)
        .string('FECHA APERT')
        .style(styleTITULOS);

    ws.cell(3, 2)
        .string(fechaCompleta)
        .style(style);
    ws.cell(3, 3)
        .string(req.body.cuit)
        .style(style);
    ws.cell(3, 4)
        .string(req.body.razon)
        .style(style);
    ws.cell(3, 5)
        .string(req.body.promotor)
        .style(style);
    ws.cell(3, 6)
        .string('Cuenta PYME')
        .style(style);
    ws.cell(3, 7)
        .string('')
        .style(style);
    ws.cell(3, 8)
        .string('P')
        .style(style);
    ws.cell(3, 9)
        .string('')
        .style(style);
    ws.cell(3, 10)
        .string(req.body.n_sucursal)
        .style(style);
    ws.cell(3, 11)
        .string('')
        .style(style);
    ws.cell(3, 12)
        .string('')
        .style(style);
    /* ---------------- */


    /* CELDAS WEB */

    ws.cell(5, 5)
        .string('FILA WEB')
        .style(styleWeb);

    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    ws.cell(6, 1)
        .string('PROMOTOR RAILCOM')
        .style(styleTITULOS);

    // Set value of cell B1 to 200 as a number type styled with paramaters of style
    ws.cell(6, 2)
        .string('N° DE SUCURSAL')
        .style(styleTITULOS);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 3)
        .string('CUIT')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 4)
        .string('RAZON SOCIAL')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 5)
        .string('CONDICIÓN ANTE IVA E IIGG (RI, MT, EX)')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 6)
        .string('CONDICIÓN ANTE IIBB (LOCAL, CM, EX, RS)')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 7)
        .string('TELÉFONO')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 8)
        .string('E-MAIL')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 9)
        .string('FACT BRUTA ANUAL (SI POSEE MENOS DE 1 AÑO: $1.000.000)')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 10)
        .string('ULTIMA FECHA DE CIERRE DE EJERCICIO')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 11)
        .string('DNI, APELLIDO Y NOMBRE DE SOCIO - REPRESENTANTE LEGAL - FIRMANTE')
        .style(styleTITULOS);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 12)
        .string('% DE PARTICIPACIÓN SOCIETARIA.')
        .style(styleTITULOS);



    /* -------------- */
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 1)
        .string(req.body.promotor)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 2)
        .string(req.body.n_sucursal)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 3)
        .string(req.body.cuit)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 4)
        .string(req.body.razon)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 5)
        .string(req.body.cond_iva)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 6)
        .string(req.body.cond_iibb)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 7)
        .string(req.body.telefono)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(7, 8)
        .string(req.body.email)
        .style(style);

    // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
    ws.cell(7, 9)
        .string(req.body.fac_anual)
        .style(style);
    // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
    ws.cell(7, 10)
        .string(req.body.ult_cierre)
        .style(style);
    // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
    ws.cell(7, 11)
        .string(req.body.DNI + '/' + req.body.nombre + ' ' + req.body.apellido)
        .style(style);
    // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
    ws.cell(7, 12)
        .string(req.body.porc_part)
        .style(style);


    /* ---------- */


    /* CELDAS SHAREPOINT */

    ws.cell(9, 5)
        .string('FILA SHAREPOINT')
        .style(styleShare);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 1)
        .string('ID')
        .style(styleTITULOS);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 2)
        .string('Comercializadora Nombre')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 3)
        .string('Tipo de Venta')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 4)
        .string('CUIT')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 5)
        .string('Nombre Razon Social')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 6)
        .string('Sucursal')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 7)
        .string('Producto Ofrecido')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 8)
        .string('Nro Establecimiento')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 9)
        .string('Tipo de elemento')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 10)
        .string('Ruta de acceso')
        .style(styleTITULOS);

    /* ------------------------------------------------------------------ */


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 2)
        .string('Railcom')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 3)
        .string('Cuenta')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 4)
        .string(req.body.cuit)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 5)
        .string(req.body.razon)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 6)
        .string(req.body.n_sucursal)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 7)
        .string('Cuenta PYME')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 8)
        .string('')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 9)
        .string('')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(11, 10)
        .string('')
        .style(style);



    wb.write(__dirname + "/temps/" + `CC PYME CARGA WEB ${req.body.razon}.xlsx`); //creacion del archivo

    const output = `
    <h3>Detalle</h3>
    <ul>  
      <li>Promotor: ${req.body.promotor}</li>
      <li>Razon Social: ${req.body.razon}</li>
      <li>CUIT: ${req.body.cuit}</li>
    </ul>
  `;

    // create reusable transporter object using the default SMTP transport
    let transporter = nodemailer.createTransport({
        host: 'smtpout.secureserver.net',
        port: 465,
        secure: true, // true for 465, false for other ports
        auth: {
            user: 'cartas@railcom.com.ar', // generated ethereal user
            pass: 'Mayo2020!'  // generated ethereal password
        },
        tls: {
            rejectUnauthorized: false
        }
    });

    // setup email data with unicode symbols
    let mailOptions = {
        from: `GESTION PYME <${req.body.mpromotor}>`, // sender address
        to: `altas.railcom@gmail.com`, // list of receivers
        subject: `CC PYME CARGA WEB ${req.body.razon}`, // Subject line
        text: 'Hello world?', // plain text body
        html: output,
        // html body
        attachments: [
            {
                path: __dirname + "/temps/" + `CC PYME CARGA WEB ${req.body.razon}.xlsx`
            },
            {
                path: __dirname + "/temps/" + req.files['constancia'][0].filename,
                contentType: 'application/pdf'
            }
            ,
            {
                path: __dirname + "/temps/" + req.files['estatuto'][0].filename,
                contentType: 'application/pdf'
            }
            ,
            {
                filename: `ULTIMO BALANCE ${req.body.razon}.pdf`,
                path: __dirname + "/temps/" + req.files['ultimobalance'][0].filename,
                contentType: 'application/pdf'
            }
            ,
            {
                filename: `DNI FRENTE ${req.body.nombre} ${req.body.apellido}.jpg`,
                path: __dirname + "/temps/" + req.files['dnifrente'][0].filename,
                contentType: 'image/jpg'
            }
            ,
            {
                filename: `DNI DORSO ${req.body.nombre} ${req.body.apellido}.jpg`,
                path: __dirname + "/temps/" + req.files['dnidorso'][0].filename,
                contentType: 'image/jpg'
            }

        ]

    };

    // send mail with defined transport object
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.log(error);
        }
        fs.unlinkSync(__dirname + "/temps/" + `CC PYME CARGA WEB ${req.body.razon}.xlsx`)//Archivo eliminado
        fs.unlinkSync(__dirname + "/temps/" + req.files['constancia'][0].filename)//Archivo eliminado
        fs.unlinkSync(__dirname + "/temps/" + req.files['estatuto'][0].filename)//Archivo eliminado
        fs.unlinkSync(__dirname + "/temps/" + req.files['ultimobalance'][0].filename)//Archivo eliminado
        fs.unlinkSync(__dirname + "/temps/" + req.files['dnifrente'][0].filename)//Archivo eliminado
        fs.unlinkSync(__dirname + "/temps/" + req.files['dnidorso'][0].filename)//Archivo eliminado
    });
    let tlForm
    let promotorForm
    let razonForm
    let cuitForm

    if (req.body.promotor.length > 0) {
        promotorForm = req.body.promotor
    } else {
        promotorForm = '';
    }
    if (req.body.razon.length > 0) {
        razonForm = req.body.razon
    } else {
        razonForm = '';
    }
    if (req.body.cuit.length > 0) {
        cuitForm = req.body.cuit
    } else {
        cuitForm = '';
    }
    if (req.body.telefono.length > 0) {
        tlForm = req.body.telefono
    } else {
        tlForm = '';
    }
    const outputVendedor = `
  <h3>Detalle</h3>
  <ul>  
    <li>Promotor: ${promotorForm}</li>
    <li>Razon Social: ${razonForm}</li>
    <li>CUIT: ${cuitForm}</li>
    <li>Numero de Telefono: ${tlForm}</li>
  </ul>
`;
    let mailVendedor = {
        from: `GESTION PYME <${req.body.mpromotor}>`, // sender address
        to: req.body.mpromotor, // list of receivers
        subject: `CC PYME CARGA WEB ${req.body.razon}`, // Subject line
        html: outputVendedor, // html body
    };
    transporter.sendMail(mailVendedor, (error, info) => {
        if (error) {
            return console.log(error);
        }
    });
    res.render('juridicas', { mensajeJuridicas: `Formulario enviado con exito a Administración y a ${req.body.mpromotor}`, style: 'formulario.css' })



});

app.post('/fisicas', function (req, res) {

    // Create a new instance of a Workbook class
    var wb = new xl.Workbook();

    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet 1');

    // Create a reusable style
    var style = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleShare = wb.createStyle({
        font: {
            color: 'white',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#0008a3',
            fgColor: '#0008a3',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleTITULOS = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#FFFF00',
            fgColor: '#FFFF00',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleHist = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: '#f50000',
            fgColor: '#f50000',
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });


    /* CELDAS HISTORIAL */

    ws.cell(1, 5)
        .string('FILA HISTORIAL')
        .style(styleHist);

    /*  */

    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    ws.cell(2, 2)
        .string('FECHA CARGA SOLICITUD')
        .style(styleTITULOS);

    // Set value of cell B1 to 200 as a number type styled with paramaters of style
    ws.cell(2, 3)
        .string('CUIT')
        .style(styleTITULOS);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 4)
        .string('RAZON SOCIAL')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 5)
        .string('PROMOTOR RAILCOM')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 6)
        .string('PRODUCTO OFRECIDO')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 7)
        .string('N° SUC')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 8)
        .string('TELEFONO')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 9)
        .string('MAIL')
        .style(styleTITULOS);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 10)
        .string('ID(BCO)')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 11)
        .string('N° CTA CTE')
        .style(styleTITULOS);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(2, 12)
        .string('FECHA APERT')
        .style(styleTITULOS);

    /* OBTENER FECHA */

    let today = new Date();
    let year = today.getFullYear();
    let month = today.getMonth() + 1;
    let day = today.getDate();
    let fechaCompleta = day + '/' + month + '/' + year;



    /* -------------- */
    ws.cell(3, 2)
        .string(fechaCompleta)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 3)
        .string(req.body.cuit)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 4)
        .string(req.body.razon)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 5)
        .string(req.body.promotor)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 6)
        .string(req.body.productos)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 7)
        .string(req.body.n_sucursal)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 8)
        .string(req.body.mcliente)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 9)
        .string(req.body.telefono)
        .style(style);

    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 10)
        .string('')
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 11)
        .string('')
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(3, 12)
        .string('')
        .style(style);

    /* CELDAS SHAREPOINT */
    ws.cell(5, 5)
        .string('FILA SHAREPOINT')
        .style(styleShare);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 1)
        .string('ID')
        .style(styleTITULOS);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 2)
        .string('Comercializadora Nombre')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 3)
        .string('Tipo de Venta')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 4)
        .string('CUIT')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 5)
        .string('Nombre Razon Social')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 6)
        .string('Sucursal')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 7)
        .string('Producto Ofrecido')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 8)
        .string('MAIL')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 9)
        .string('TELEFONO')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 10)
        .string('Nro Establecimiento')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 11)
        .string('Tipo de elemento')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 12)
        .string('Ruta de acceso')
        .style(styleTITULOS);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 2)
        .string('Railcom')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 3)
        .string('Cuenta')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 4)
        .string(req.body.cuit)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 5)
        .string(req.body.razon)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 6)
        .string(req.body.n_sucursal)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 7)
        .string(req.body.productos)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 8)
        .string(req.body.mcliente)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 9)
        .string(req.body.telefono)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 10)
        .string('')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 11)
        .string('')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(7, 12)
        .string('')
        .style(style);


    wb.write(`CC ${req.body.razon}.xlsx`); //creacion del archivo
    let direcForm
    let localidadForm
    let tlForm

    if (req.body.direccion.length > 0) {
        direcForm = req.body.direccion
    } else {
        direcForm = '';
    }
    if (req.body.localidad.length > 0) {
        localidadForm = req.body.localidad
    } else {
        localidadForm = '';
    }
    if (req.body.telefono.length > 0) {
        tlForm = req.body.telefono
    } else {
        tlForm = '';
    }
    const outputVendedor = `
  <h3>Detalle</h3>
  <ul>  
    <li>Promotor: ${req.body.promotor}</li>
    <li>Razon Social: ${req.body.razon}</li>
    <li>CUIT: ${req.body.cuit}</li>
    <li>Dirección: ${direcForm}</li>
    <li>Localidad: ${localidadForm}</li>
    <li>Numero de Telefono: ${tlForm}</li>
    <li>Numero de Telefono: ${req.body.mcliente}</li>
    <li>Numero de Sucursal: ${req.body.n_sucursal}</li>
    <li>Producto Ofrecido: ${req.body.productos}</li>
  </ul>
`;
    const output = `
<h3>Detalle</h3>
<ul>  
<li>Promotor: ${req.body.promotor}</li>
<li>Razon Social: ${req.body.razon}</li>
<li>CUIT: ${req.body.cuit}</li>
<li>Dirección: ${direcForm}</li>
<li>Localidad: ${localidadForm}</li>
<li>Numero de Telefono: ${tlForm}</li>
<li>Numero de Telefono: ${req.body.mcliente}</li>
<li>Numero de Sucursal: ${req.body.n_sucursal}</li>
<li>Producto Ofrecido: ${req.body.productos}</li>
</ul>
`;

    // create reusable transporter object using the default SMTP transport
    let transporter = nodemailer.createTransport({
        host: 'smtpout.secureserver.net',
        port: 465,
        secure: true, // true for 465, false for other ports
        auth: {
            user: 'cartas@railcom.com.ar', // generated ethereal user
            pass: 'Mayo2020!'  // generated ethereal password
        },
        tls: {
            rejectUnauthorized: false
        }
    });

    // setup email data with unicode symbols
    let mailOptions = {
        from: `GESTION EN SUCURSAL <${req.body.mpromotor}>`, // sender address
        to: `jesus.parra@railcom.com.ar`, // list of receivers altas.railcom@gmail.com
        subject: `CC EN SUCURSAL ${req.body.razon}`, // Subject line
        text: 'Hello world?', // plain text body
        html: output, // html body
        attachments:
        {
            path: `CC ${req.body.razon}.xlsx`
        }
    };

    // send mail with defined transport object
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.log(error);
        }
        fs.unlinkSync(`CC ${req.body.razon}.xlsx`)//Archivo eliminado
    });
    let mailVendedor = {
        from: `GESTION EN SUCURSAL <${req.body.mpromotor}>`, // sender address
        to: `${req.body.mpromotor}`, // list of receivers
        subject: `CC EN SUCURSAL ${req.body.razon}`, // Subject line
        text: 'Hello world?', // plain text body
        html: outputVendedor, // html body
    };
    transporter.sendMail(mailVendedor, (error, info) => {
        if (error) {
            return console.log(error);
        }
    });
    res.render('fisicas', { mensajeFisicas: `Formulario enviado con exito a Administración y ${req.body.mpromotor}`, style: 'formulario.css' });



});

app.post('/soloDatos', function (req, res) {
    let direcForm
    let localidadForm
    let tlForm
    let promotorForm
    let razonForm
    let cuitForm

    if (req.body.promotor.length > 0) {
        promotorForm = req.body.promotor
    } else {
        promotorForm = '';
    }
    if (req.body.razon.length > 0) {
        razonForm = req.body.razon
    } else {
        razonForm = '';
    }
    if (req.body.cuit.length > 0) {
        cuitForm = req.body.cuit
    } else {
        cuitForm = '';
    }
    if (req.body.direccion.length > 0) {
        direcForm = req.body.direccion
    } else {
        direcForm = '';
    }
    if (req.body.localidad.length > 0) {
        localidadForm = req.body.localidad
    } else {
        localidadForm = '';
    }
    if (req.body.telefono.length > 0) {
        tlForm = req.body.telefono
    } else {
        tlForm = '';
    }
    const outputVendedor = `
  <h3>Detalle</h3>
  <ul>  
    <li>Promotor: ${promotorForm}</li>
    <li>Razon Social: ${razonForm}</li>
    <li>CUIT: ${cuitForm}</li>
    <li>Dirección: ${direcForm}</li>
    <li>Localidad: ${localidadForm}</li>
    <li>Numero de Telefono: ${tlForm}</li>
  </ul>
`;

    // create reusable transporter object using the default SMTP transport
    let transporter = nodemailer.createTransport({
        host: 'smtpout.secureserver.net',
        port: 465,
        secure: true, // true for 465, false for other ports
        auth: {
            user: 'cartas@railcom.com.ar', // generated ethereal user
            pass: 'Mayo2020!'  // generated ethereal password
        },
        tls: {
            rejectUnauthorized: false
        }
    });

    let mailVendedor = {
        from: `DATOS WebApp <${req.body.mpromotor}>`, // sender address
        to: req.body.mpromotor, // list of receivers
        subject: `DATOS WebApp ${req.body.razon}`, // Subject line
        text: 'Hello world?', // plain text body
        html: outputVendedor, // html body
    };
    transporter.sendMail(mailVendedor, (error, info) => {
        if (error) {
            return console.log(error);
        }
    });
    res.render('soloDatos', { mensaje: `Mensaje enviado con exito a ${req.body.mpromotor}`, style: 'formulario.css' });

});

app.delete('/logout', (req, res) => {
    req.logOut()
    res.redirect('/');
})


function checkAuthenticated(req, res, next) {
    if (req.isAuthenticated()) {
        return next()
    }

    res.redirect('/login')
}

function checkNotAuthenticated(req, res, next) {
    if (req.isAuthenticated()) {
        return res.redirect('/inicio')
    }
    next()
}
//START THE SERVER
const server = app.listen(app.get('port'), () => {
    console.log('Servidor en puerto', app.get('port'))
})



