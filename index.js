//Import packages
require('dotenv').config();
const Axios = require("axios").default;


//const to get insurances to send
const getInsurances = async () => {
    try {
        const res = await Axios(process.env.URL);
        return res.data;
    } catch (error) {
        console.log(error);
    }
}

//generate csv
const GenerateCSV = async () => {
    var CSV = require('excel4node');
    var insurances = await getInsurances();
    var Libro = new CSV.Workbook();
    var Hoja = Libro.addWorksheet("clientes");

    let c = 1;
    Hoja.cell(1, c++).string("Tipo de Record");
    Hoja.cell(1, c++).string("Sponsor");
    Hoja.cell(1, c++).string("Producto");
    Hoja.cell(1, c++).string("Póliza");
    Hoja.cell(1, c++).string("Cédula");
    Hoja.cell(1, c++).string("Plan");
    Hoja.cell(1, c++).string("Prima");
    Hoja.cell(1, c++).string("Método de pago");
    Hoja.cell(1, c++).string("Frecuencia de pago");
    Hoja.cell(1, c++).string("Tipo de cobertura");
    Hoja.cell(1, c++).string("Fecha Efectividad creación");
    Hoja.cell(1, c++).string("Vendedor");
    Hoja.cell(1, c++).string("Nombre");
    Hoja.cell(1, c++).string("Segundo Nombre");
    Hoja.cell(1, c++).string("Apellidos");
    Hoja.cell(1, c++).string("Dirección 1");
    Hoja.cell(1, c++).string("Dirección 2");
    Hoja.cell(1, c++).string("Dirección 3");
    Hoja.cell(1, c++).string("Dirección 4");
    Hoja.cell(1, c++).string("Ciudad");
    Hoja.cell(1, c++).string("Provincia/Estado");
    Hoja.cell(1, c++).string("País");
    Hoja.cell(1, c++).string("Teléfono");
    Hoja.cell(1, c++).string("Teléfono 2");
    Hoja.cell(1, c++).string("Email");
    Hoja.cell(1, c++).string("Fecha nacimiento");
    Hoja.cell(1, c++).string("Sexo");
    Hoja.cell(1, c++).string("Relación dependiente");
    Hoja.cell(1, c++).string("Tarjeta");
    Hoja.cell(1, c++).string("Fecha expiración de tarjeta");
    Hoja.cell(1, c++).string("Tipo de tarjeta");
    Hoja.cell(1, c++).string("Número de cuenta");
    Hoja.cell(1, c++).string("Tipo de cuenta");
    Hoja.cell(1, c++).string("Código de colector");
    Hoja.cell(1, c++).string("Referencia 1");
    Hoja.cell(1, c++).string("Referencia 2");
    Hoja.cell(1, c++).string("Referencia 3 - Ciclo CMF");
    Hoja.cell(1, c++).string("Referencia 4");
    Hoja.cell(1, c++).string("Referencia 5");
    Hoja.cell(1, c++).string("Referencia 6");
    Hoja.cell(1, c++).string("Nombre Benef 1");
    Hoja.cell(1, c++).string("Cédula Benef 1");
    Hoja.cell(1, c++).string("Relación Benef 1");
    Hoja.cell(1, c++).string("Porcentaje Benef 1");
    Hoja.cell(1, c++).string("Nombre Benef 2");
    Hoja.cell(1, c++).string("Cédula Benef 2");
    Hoja.cell(1, c++).string("Relación Benef 2");
    Hoja.cell(1, c++).string("Porcentaje Benef 2");
    Hoja.cell(1, c++).string("Nombre Benef 3");
    Hoja.cell(1, c++).string("Cédula Benef 3");
    Hoja.cell(1, c++).string("Relación Benef 3");
    Hoja.cell(1, c++).string("Porcentaje Benef 3");
    Hoja.cell(1, c++).string("Nombre Benef 4");
    Hoja.cell(1, c++).string("Cédula Benef 4");
    Hoja.cell(1, c++).string("Relación Benef 4");
    Hoja.cell(1, c++).string("Porcentaje Benef 4");
    Hoja.cell(1, c++).string("Nombre Benef 5");
    Hoja.cell(1, c++).string("Cédula Benef 5");
    Hoja.cell(1, c++).string("Relación Benef 5");
    Hoja.cell(1, c++).string("Porcentaje Benef 5");
    
    for(let p=1; p<= insurances.length; p++) {
        Hoja.cell(p + 1, 5).string(insurances[p - 1].identityNumber);
        Hoja.cell(p + 1, 6).string((insurances[p - 1].insurancePlan).toString());
        Hoja.cell(p + 1, 10).string(insurances[p - 1].coverageType);
        Hoja.cell(p + 1, 11).string((insurances[p - 1].efectivityDate).toString());
        Hoja.cell(p + 1, 13).string(insurances[p - 1].firstname);
        Hoja.cell(p + 1, 14).string(insurances[p - 1].secondname);
        Hoja.cell(p + 1, 15).string(insurances[p - 1].surnames);
        Hoja.cell(p + 1, 16).string(insurances[p - 1].residentialAddress);
        Hoja.cell(p + 1, 22).string(insurances[p - 1].country);
        Hoja.cell(p + 1, 23).string(insurances[p - 1].phone);
        Hoja.cell(p + 1, 25).string(insurances[p - 1].email);
        Hoja.cell(p + 1, 26).string((insurances[p - 1].dateOfBirth).toString());
        Hoja.cell(p + 1, 27).string(insurances[p - 1].gender);
        Hoja.cell(p + 1, 28).string(insurances[p - 1].maritalStatus);
        Hoja.cell(p + 1, 29).string(insurances[p - 1].tokenizedCard);
        Hoja.cell(p + 1, 37).string(insurances[p - 1].cycle);
    }

    let docUrl = `document-`+ Date.now()+`.xlsx`;
    Libro.write('docs/' + docUrl);
}

GenerateCSV();

