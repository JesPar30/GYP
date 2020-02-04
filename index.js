const path = require("path");
const express = require('express')
const bodyParser = require('body-parser')
const exphbs = require('express-handlebars')
const imaps = require('imap-simple');
const xl = require('excel4node');
const fs = require ('fs');
const nodemailer = require('nodemailer')

const app = express()

let config = {
    imap: {
        user: '',
        password: '',
        host: 'imap.secureserver.net',
        port: 993,
        tls: true,
        authTimeout: 3000,
        socketTimeout: 0,
        keepalive:true
    }
};

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

app.get('/', function (req, res) {
    res.render('home');
});

app.post('/', function (req, res) {

    config.imap.user = req.body.email
    config.imap.password = req.body.password

    imaps.connect(config).then(function (connection) {
        connection.keepalive=true;
            console.log('sesion iniciada')
            res.render('inicio');

    }).catch((err) => {
            res.render('home');
    });

});
app.post('/inicio', function (req, res) {
    
// Create a new instance of a Workbook class
var wb = new xl.Workbook();
 
// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');
var ws2 = wb.addWorksheet('Sheet 2');
 
// Create a reusable style
var style = wb.createStyle({
  font: {
    color: 'black',
    size: 12,
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -',
});
 
// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1, 1)
  .string('PROMOTOR RAILCOM')
  .style(style);
 
// Set value of cell B1 to 200 as a number type styled with paramaters of style
ws.cell(1, 2)
  .string('N° DE SUCURSAL')
  .style(style);
 
// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 3)
  .string('CUIT')
  .style(style);

  
  // Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 4)
.string('RAZON SOCIAL')
.style(style);


// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 5)
  .string('CONDICIÓN ANTE IVA E IIGG (RI, MT, EX)')
  .style(style);

  
  // Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 6)
.string('CONDICIÓN ANTE IIBB (LOCAL, CM, EX, RS)')
.style(style);


// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 7)
  .string('TELÉFONO')
  .style(style);

  
  // Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 8)
.string('E-MAIL')
.style(style);


// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 9)
  .string('FACT BRUTA ANUAL (SI POSEE MENOS DE 1 AÑO: $1.000.000)')
  .style(style);

  
  // Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 10)
.string('ULTIMA FECHA DE CIERRE DE EJERCICIO')
.style(style);


// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 11)
  .string('DNI, APELLIDO Y NOMBRE DE SOCIO - REPRESENTANTE LEGAL - FIRMANTE')
  .style(style);

  // Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 12)
.string('% DE PARTICIPACIÓN SOCIETARIA.')
.style(style);


 
// Set value of cell A2 to 'string' styled with paramaters of style
ws.cell(2, 1)
  .string(req.body.nombre)
  .style(style);
 
// Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
ws.cell(2, 8)
  .string(req.body.email)
  .style(style)
  .style({font: {size: 14}});
 
  wb.write(`CC PYME ${req.body.nombre}.xlsx`); //creacion del archivo

  const output = `
    <h3>Detalle</h3>
    <ul>  
      <li>Nombre: ${req.body.nombre}</li>
      <li>Email: ${req.body.email}</li>
    </ul>
  `;

  // create reusable transporter object using the default SMTP transport
  let transporter = nodemailer.createTransport({
    host: 'smtpout.secureserver.net',
    port: 465,
    secure: true, // true for 465, false for other ports
    auth: {
        user: 'jesus.parra@railcom.com.ar', // generated ethereal user
        pass: '343434j.'  // generated ethereal password
    },
    tls:{
      rejectUnauthorized:false
    }
  });

  // setup email data with unicode symbols
  let mailOptions = {
      from: 'GESTION PYME <jesus.parra@railcom.com.ar>', // sender address
      to: `${req.body.email}`, // list of receivers
      subject: `CC PYME ${req.body.nombre}`, // Subject line
      text: 'Hello world?', // plain text body
      html: output, // html body
      attachments: 
        {   
            path: `CC PYME ${req.body.nombre}.xlsx`
        }
  };

  // send mail with defined transport object
  transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
          return console.log(error);
      }
      fs.unlinkSync(`CC PYME ${req.body.nombre}.xlsx`)//Archivo eliminado
      res.render('inicio', `<script>alert("CC PYME ${req.body.nombre}.xlsx Enviado con Exito")</script>`);
  });




});

app.use(function(req, res, next){
    res.status(404).render('404_error_template', {title: "Página no encontrada"});
});



app.get('*', function(req, res){
    res.status(404).render('404_error_template', {title: "Página no encontrada"});
});


//START THE SERVER
const server = app.listen(app.get('port'), () => {
    console.log('Servidor en puerto', app.get('port'))
})



