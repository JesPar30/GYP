﻿const path = require("path");
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
const upload = multer({ dest: __dirname })
const Recaptcha = require('express-recaptcha').RecaptchaV3;

let recaptcha = new Recaptcha('6LcTvpMUAAAAALDaeDO8m-a6EfsNDQlbM7YQH8M2', '6LcTvpMUAAAAACqKJGMDn3WydcjPc3uHAn4MmAr7');
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
//HANDLEBARS
app.engine('.hbs', exphbs({ extname: '.hbs', defaultLayout: 'main.hbs' }));
app.set('view engine', '.hbs');

app.get('/', checkAuthenticated, function (req, res) {
    res.render(`inicio`);
});
app.get('/login', checkNotAuthenticated, function (req, res) {
    res.render(`login`);
});

app.get('/fisicas', checkAuthenticated, function (req, res) {
    res.render(`fisicas`);
});

app.get('/juridicas', checkAuthenticated, function (req, res) {
    res.render(`juridicas`);
});

app.post('/', passport.authenticate('pop3', { failureRedirect: '/' }),
    function (req, res) {
        userr = req.body.username
        passs = req.body.password
        res.render('inicio');
    });
var cpUpload = upload.fields([{ name: 'constancia', maxCount: 1 }, { name: 'estatuto', maxCount: 1 }, { name: 'ultimobalance', maxCount: 1 }, { name: 'dnifrente', maxCount: 1 }, { name: 'dnidorso', maxCount: 1 }])
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
            background: 'blue'
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleHist = wb.createStyle({
        font: {
            color: 'black',
            size: 10,
            background: 'yellow'
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });
    var styleWeb = wb.createStyle({
        font: {
            color: 'white',
            size: 10,
            background: 'black'
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

    ws.cell(1, 2)
        .string('FILA HISTORIAL')
        .style(styleHist);

    ws.cell(2, 2)
        .string('FECHA CARGA SOLICITUD')
        .style(style);
    ws.cell(2, 3)
        .string('CUIT')
        .style(style);
    ws.cell(2, 4)
        .string('RAZON SOCIAL')
        .style(style);
    ws.cell(2, 5)
        .string('PROMOTOR RAILCOM')
        .style(style);
    ws.cell(2, 6)
        .string('N° REGISTRO STADER')
        .style(style);
    ws.cell(2, 7)
        .string('ESTADO (A - R - P - S)')
        .style(style);
    ws.cell(2, 8)
        .string('ENTREGADA A PROMOTOR')
        .style(style);
    ws.cell(2, 9)
        .string('N° SUC')
        .style(style);
    ws.cell(2, 10)
        .string('N° CTA CTE')
        .style(style);
    ws.cell(2, 11)
        .string('FECHA APERT')
        .style(style);

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
        .string('')
        .style(style);
    ws.cell(3, 7)
        .string('P')
        .style(style);
    ws.cell(3, 8)
        .string('')
        .style(style);
    ws.cell(3, 9)
        .string(req.body.n_sucursal)
        .style(style);
    ws.cell(3, 10)
        .string('')
        .style(style);
    ws.cell(3, 11)
        .string('')
        .style(style);
    /* ---------------- */


    /* CELDAS WEB */

    ws.cell(5, 2)
        .string('FILA WEB')
        .style(styleWeb);

    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    ws.cell(6, 1)
        .string('PROMOTOR RAILCOM')
        .style(style);

    // Set value of cell B1 to 200 as a number type styled with paramaters of style
    ws.cell(6, 2)
        .string('N° DE SUCURSAL')
        .style(style);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 3)
        .string('CUIT')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 4)
        .string('RAZON SOCIAL')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 5)
        .string('CONDICIÓN ANTE IVA E IIGG (RI, MT, EX)')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 6)
        .string('CONDICIÓN ANTE IIBB (LOCAL, CM, EX, RS)')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 7)
        .string('TELÉFONO')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 8)
        .string('E-MAIL')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 9)
        .string('FACT BRUTA ANUAL (SI POSEE MENOS DE 1 AÑO: $1.000.000)')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 10)
        .string('ULTIMA FECHA DE CIERRE DE EJERCICIO')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 11)
        .string('DNI, APELLIDO Y NOMBRE DE SOCIO - REPRESENTANTE LEGAL - FIRMANTE')
        .style(style);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(6, 12)
        .string('% DE PARTICIPACIÓN SOCIETARIA.')
        .style(style);



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
        .style(style);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 2)
        .string('Comercializadora Nombre')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 3)
        .string('Tipo de Venta')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 4)
        .string('CUIT')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 5)
        .string('Nombre Razon Social--------------------------------------------')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 6)
        .string('Sucursal')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 7)
        .string('Producto Ofrecido-----------------------------------------')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 8)
        .string('Nro Establecimiento')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 9)
        .string('Tipo de elemento')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(10, 10)
        .string('Ruta de acceso')
        .style(style);

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



    wb.write(`CC PYME ${req.body.razon}.xlsx`); //creacion del archivo

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
            user: 'jesus.parra@railcom.com.ar', // generated ethereal user
            pass: '343434jesus.'  // generated ethereal password
        },
        tls: {
            rejectUnauthorized: false
        }
    });

    // setup email data with unicode symbols
    let mailOptions = {
        from: 'GESTION PYME <jesus.parra@railcom.com.ar>', // sender address
        to: `${req.body.email}`, // list of receivers
        subject: `CC PYME CARGA WEB ${req.body.razon}`, // Subject line
        text: 'Hello world?', // plain text body
        html: output, // html body
        attachments: [
            {
                path: `CC PYME ${req.body.razon}.xlsx`
            },
            {
                filename: `CONSTANCIA AFIP ${req.body.razon}.pdf`,
                path: req.files['constancia'][0].filename,
                contentType: 'application/pdf'
            }
            ,
            {
                filename: `ESTATUTO - C.S ${req.body.razon}.pdf`,
                path: req.files['estatuto'][0].filename,
                contentType: 'application/pdf'
            }
            ,
            {
                filename: `ULTIMO BALANCE ${req.body.razon}.pdf`,
                path: req.files['ultimobalance'][0].filename,
                contentType: 'application/pdf'
            }
            ,
            {
                filename: `DNI FRENTE ${req.body.nombre} ${req.body.apellido}.jpg`,
                path: req.files['dnifrente'][0].filename,
                contentType: 'image/jpg'
            }
            ,
            {
                filename: `DNI DORSO ${req.body.nombre} ${req.body.apellido}.jpg`,
                path: req.files['dnidorso'][0].filename,
                contentType: 'image/jpg'
            }

        ]

    };

    // send mail with defined transport object
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.log(error);
        }
        fs.unlink(`CC PYME ${req.body.razon}.xlsx`)//Archivo eliminado
        fs.unlink(req.files['constancia'][0].filename)//Archivo eliminado
        fs.unlink(req.files['estatuto'][0].filename)//Archivo eliminado
        fs.unlink(req.files['ultimobalance'][0].filename)//Archivo eliminado
        fs.unlink(req.files['dnifrente'][0].filename)//Archivo eliminado
        fs.unlink(req.files['dnidorso'][0].filename)//Archivo eliminado
    });
    res.render('juridicas')



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
            color: 'black',
            size: 10,
            background: 'blue'
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -',
    });

    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    ws.cell(1, 1)
        .string('FECHA CARGA SOLICITUD')
        .style(style);

    // Set value of cell B1 to 200 as a number type styled with paramaters of style
    ws.cell(1, 2)
        .string('CUIT')
        .style(style);

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 3)
        .string('RAZON SOCIAL')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 4)
        .string('PROMOTOR RAILCOM')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 5)
        .string('PRODUCTO OFRECIDO')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 6)
        .string('N° SUC')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 7)
        .string('ID(BCO)')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 8)
        .string('N° CTA CTE')
        .style(style);


    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(1, 9)
        .string('FECHA APERT')
        .style(style);

    /* OBTENER FECHA */

    let today = new Date();
    let year = today.getFullYear();
    let month = today.getMonth() + 1;
    let day = today.getDate();
    let fechaCompleta = day + '/' + month + '/' + year;



    /* -------------- */


    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 2)
        .string(fechaCompleta)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 3)
        .string(req.body.cuit)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 4)
        .string(req.body.razon)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 5)
        .string(req.body.promotor)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 6)
        .string(req.body.productos)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 7)
        .string(req.body.n_sucursal)
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 8)
        .string('')
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 9)
        .string('')
        .style(style);
    // Set value of cell A2 to 'string' styled with paramaters of style
    ws.cell(2, 10)
        .string('')
        .style(style);

    /* CELDAS SHAREPOINT */

    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(4, 5)
        .string('SHAREPOINT')
        .style(styleShare);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 2)
        .string('Railcom')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 3)
        .string('Cuenta')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 4)
        .string(req.body.cuit)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 5)
        .string(req.body.razon)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 6)
        .string(req.body.n_sucursal)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 7)
        .string(req.body.productos)
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 8)
        .string('')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 9)
        .string('')
        .style(style);
    // Set value of cell C1 to a formula styled with paramaters of style
    ws.cell(5, 10)
        .string('')
        .style(style);


    wb.write(`CC ${req.body.razon}.xlsx`); //creacion del archivo

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
            user: 'jesus.parra@railcom.com.ar', // generated ethereal user
            pass: '343434jesus.'  // generated ethereal password
        },
        tls: {
            rejectUnauthorized: false
        }
    });

    // setup email data with unicode symbols
    let mailOptions = {
        from: 'GESTION FISICA <jesus.parra@railcom.com.ar>', // sender address
        to: `jesus.parra@railcom.com.ar`, // list of receivers
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
    res.redirect('fisicas');



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



