// Generales
const ss = SpreadsheetApp.getActiveSpreadsheet()
const sheetEnvios = ss.getSheetByName('Envios')
const valueToCheckEnviosTV = sheetEnvios.getRange('BC1').getValue();
const lrdata = sheetEnvios.getRange('A3:A').getValues().filter(String).length.toString()
const data = sheetEnvios.getRange(3, 1, lrdata, sheetEnvios.getLastColumn()).getValues()
const currentHour = new Date().getHours();
//const salo = '5215559895500'
const quique = '5215552172262'
const touch = '5215562929714'

function WATV() {
  if (valueToCheckEnviosTV > 0) {
    if (!shouldRunTrigger()) return;
    WATV1() // On-Hold
    WATV2() // Pending
    WATV3() // Procesando
    WATV4() // Completado
    WATV5() // Cancelado
    WATV6() // Refunded
    WATV7() // Failed
    WATV8() // Hr Update
    // WATV9() //
  }
}

function shouldRunTrigger() {
  var date = new Date();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  date = Utilities.formatDate(date, timezone, 'yyyy/MM/dd HH:mm:ss');
  var hours = new Date(date).getHours();
  var hrentrada = 5; // Cambio en la hora de entrada
  var hrsalida = 6; // Cambio en la hora de salida

  console.log(date)
  console.log(hours)

  if (hours >= hrentrada && hours <= hrsalida) {
    console.log('No Corre por horarios') // Cambio en el mensaje de consola
    return false; // Cambio en el retorno cuando no está dentro del rango permitido
  } else {
    console.log('Si Corre') // Cambio en el mensaje de consola
    return true; // Cambio en el retorno cuando está dentro del rango permitido
  }
}

/*
function join(t, a, s) {
  function format(m) {
    let f = new Intl.DateTimeFormat('en', m);
    return f.format(t);
  }
  return a.map(format).join(s);
}
*/

// ON-HOLD local o foraneo
function WATV1() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    var a = [{ day: 'numeric' }, { month: 'short' }, { year: '2-digit' }];
    var date = join(Fecha, a, '-') + ' ' + Hora
    var orderdate = new Date(date)

    var now = new Date()
    var FechaInMS = orderdate.getTime()
    var add = 5 * 60 * 1000
    var minslater = FechaInMS + add
    var futureDate = new Date(minslater);

    // 1.0 Transferencia en On-Hold - Local o Foraneo
    if (
      Empresa == 'TV'
      //&& futureDate <= now
      && Fecha != ''
      && MetodoPago.includes("Transferencia")
      && Status == 'on-hold'
      && WA1 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Envíanos el comprobante de transferencia para revisarlo por la mañana y procesar tu pedido.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #*' + Pedido + '*\n' +
        'Con un total de *' + Total.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        'La cuenta para transferencia o deposito es:' + '\n' +
        'BBVA' + '\n' +
        'JOSEFINA RANGEL MARTINEZ' + '\n' +
        'Cta: 0477727364' + '\n' +
        'Clabe: 012180004777273647' + '\n\n' +

        'Esperamos el comprobante de pago por este medio para procesarlo.' + '\n' +
        '*Te pedimos depositar el monto exacto y poner en concepto o referencia el numero de pedido ' + Pedido + '*' + '\n\n' +

        'Te envio la clabe para que la puedas copiar y pegar sin errores.'

      var clabe = '012180004777273647'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      //send clabe
      SendWAApiSaloTV({
        to: WA.toString(),
        msg: clabe
      })

      var msg2 = 'Error al enviar mensaje de on-hold:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      //console.log(status_wa)

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      /*
      if (!status_wa.succes) {
        var wa_failed = SendWATV({
          to: touch,
          msg: msg2
        })
        //console.log(wa_failed)
      }
      */

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 40)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
      //range.setNote(JSON.stringify(status_wa))
      //console.log({ status_wa, Tel })
    }

    // 1.1 Clip en on-hold desues de 5 mins que no pago, local o foraneo
    else if (
      Empresa == 'TV'
      && futureDate <= now
      && Fecha != ''
      && MetodoPago.includes("Clip")
      && Status == 'on-hold'
      && WA1 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Puedes realizar el pago con confianza y procesaremos tu pedido en cuanto estemos de regreso.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #*' + Pedido + '*\n' +
        'Con un total de *' + Total.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        'El estado de tu pedido está *pendiente de pago*, elegiste Clip como metodo de pago.' + '\n' +
        //'Si aun no tienes el link para realizar el pago, por favor avisanos para enviartelo.' + '\n' +
        'Con gusto te paso por aqui el link para realizar el pago.' + '\n' +
        LinkClip + '\n\n' +
        'Si realizas el pago con tarjeta, lo veremos reflejado en automatico en nuestro sistema.' + '\n' +
        'Si realizas el pago en efectivo, te pedimos nos envies el comprobante en cuanto lo tengas.' + '\n\n' +

        'Gracias por tu compra!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de on-hold:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 40)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
      //range.setNote(JSON.stringify(status_wa))
      //console.log({ status_wa, Tel })

    }

    // 1.2 Transferencias Zelle en On-Hold - Local o Foraneo
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && MetodoPago.includes("Zelle")
      && Status == 'on-hold'
      && WA1 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          'Our service hours are daily from 10am to 10pm 🕙 and we are currently resting.' + '\n' +
          'Please send us the proof of transfer so we can review it in the morning and process your order.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hi *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'This is TouchVapes.com' + '\n\n' +

        '*We received your order #' + Pedido + '*\n' +
        'Your total ammount in USD is *' + TotalDLS.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        'Kindly make a payment for the total amount via Zelle to the following e-mail:' + '\n' +
        'zelle@touchvapes.com' + '\n' +

        'After completing the payment, please send us a screenshot as proof of payment, so we can promptly process your order.' + '\n'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de on-hold:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 40)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    /*
        // si no entra en on-hold
        else if (
          Empresa == 'TV'
          && futureDate <= now
          && Fecha != ''
          && Status != 'on-hold'
          && WA1 == ''
        ) {
          sheetEnvios.getRange(i + 3, 35).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
        }
    */
  });

}


// Pending clip
function WATV2() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    var a = [{ day: 'numeric' }, { month: 'short' }, { year: '2-digit' }];
    var date = join(Fecha, a, '-') + ' ' + Hora
    var orderdate = new Date(date)

    var now = new Date()
    var FechaInMS = orderdate.getTime()
    var add = 5 * 60 * 1000
    var minslater = FechaInMS + add
    var futureDate = new Date(minslater);

    // 2.0 Clip en pending desues de 5 mins que no pago, local o foraneo
    if (
      Empresa == 'TV'
      && futureDate <= now
      && Fecha != ''
      && PaymentID.includes("wc_clip")
      && Status == 'pending'
      && WA2 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Puedes realizar el pago con confianza y procesaremos tu pedido en cuanto estemos de regreso.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #*' + Pedido + '*\n' +
        'Con un total de *' + Total.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        'El estado de tu pedido está *pendiente de pago*, elegiste Clip como metodo de pago.' + '\n' +
        'Con gusto te paso por aqui el link para realizar el pago.' + '\n' +
        LinkClip + '\n\n' +

        'Si realizas el pago con tarjeta, lo veremos reflejado en automatico en nuestro sistema.' + '\n' +
        'Si realizas el pago en efectivo, te pedimos nos envies el comprobante en cuanto lo tengas.' + '\n\n' +

        'Gracias por tu compra!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de pending:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 41)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

    }

    // 2.1 MP en pending desues de 5 mins que no pago, local o foraneo
    else if (
      Empresa == 'TV'
      && futureDate <= now
      && Fecha != ''
      && MetodoPago.includes("Mercado")
      && Status == 'pending'
      && WA2 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Puedes realizar el pago con confianza y procesaremos tu pedido en cuanto estemos de regreso.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #*' + Pedido + '*\n' +
        'Con un total de *' + Total.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        'El estado de tu pedido está *pendiente de pago*, elegiste Mercado Pago como metodo de pago.' + '\n' +
        'Si aun no tienes el link para realizar el pago, por favor avisanos para enviartelo.' + '\n\n' +

        'Gracias por tu compra!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de pending:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 41)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

    }

  });

}


// Pedido en procesando
function WATV3() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    ///// Locales /////

    // 3.1 Local - Transferencia Pasa de on-hold a Procesando 
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && MetodoPago.includes("Transferencia")
      && WA1 != ''
      && WA3 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        '*Recibimos el pago de tu pedido* #' + Pedido + ', Muchas gracias!' + '\n\n' +

        'Ya lo estamos preparando, en cuanto esté por salir a ruta, te enviaremos una notificación con la hora aproximada de entrega, así como el contacto del mensajero.' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, en caso de alguna falla para poder hacer válida la garantía, es necesario que pruebes el producto frente al mensajero. NO HAY EXCEPCIONES*' + '\n\n' +

        'Gracias por tu compra!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
      //range.setValue(new Date()).setNumberFormat('hh:mm:ss').setHorizontalAlignment("center").setVerticalAlignment("middle");
      //range.setNote(JSON.stringify(status_wa))
      //console.log({ status_wa, Tel })
    }

    // 3.2.1 Clip pasa de on-hold o procesing a Procesando y NO le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && (MetodoPago.includes("Clip") || PaymentID.includes('wc_clip'))
      && (WA1 != '' || WA2 == '')
      && WA3 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Enviaremos tu pedido en la primera ruta por la mañana.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        '*Recibimos tu pedido #' + Pedido + ' y el pago*, Muchas gracias !' + '\n\n' +

        'Ya lo estamos preparando, en cuanto esté por salir a ruta, te enviaremos una notificación con la hora aproximada de entrega, así como el contacto del mensajero.' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, en caso de alguna falla para poder hacer válida la garantía, es necesario que pruebes el producto frente al mensajero. NO HAY EXCEPCIONES*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnviosTV.getRange(i + 2, 35).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle")
    }

    // 3.2.2 Clip pasa de on-hold o procesing a Procesando y ya le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && (MetodoPago.includes("Clip") || PaymentID.includes('wc_clip'))
      && (WA1 != '' || WA2 != '')
      && WA3 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + ' \n' +
        '*Recibimos el pago de tu pedido* #' + Pedido + ', Muchas gracias !' + '\n\n' +

        'Ya lo estamos preparando, en cuanto esté por salir a ruta, te enviaremos una notificación con la hora aproximada de entrega, así como el contacto del mensajero.' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, en caso de alguna falla para poder hacer válida la garantía, es necesario que pruebes el producto frente al mensajero. NO HAY EXCEPCIONES*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.3.1 Zelle pasa de on-hold a Procesando y NO le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && MetodoPago.includes("Zelle")
      && WA1 == ''
      && WA3 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + ' \n' +
        'This is TouchVapes.com' + '\n\n' +

        '*We received payment for your order* #' + Pedido + ' \n' +
        'Thank you very much!' + '\n\n' +

        'We are already preparing it, as soon as it is ready to leave, we will send you a notification with the approximate delivery time, as well as the delivery person contact.' + '\n\n' +

        'Thank you for your purchase!' + '\n\n' +

        '*REMINDER:*' + '\n' +
        '*As these are disposable devices, in case of any malfunction, it is necessary for you to test the product in front of the delivery person in order to validate the warranty. THERE ARE NO EXCEPTIONS*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnviosTV.getRange(i + 2, 35).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle")
    }

    // 3.3.2 Zelle pasa de on-hold a Procesando y ya le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && MetodoPago.includes("Zelle")
      && WA1 != ''
      && WA3 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + ' \n' +
        '*We received payment for your order* #' + Pedido + ' \n' +
        'Thank you very much!' + '\n\n' +

        'We are already preparing it, as soon as it is ready to leave, we will send you a notification with the approximate delivery time, as well as the delivery person contact.' + '\n\n' +

        'Thank you for your purchase!' + '\n\n' +

        '*REMINDER:*' + '\n' +
        '*As these are disposable devices, in case of any malfunction, it is necessary for you to test the product in front of the delivery person in order to validate the warranty. THERE ARE NO EXCEPTIONS*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.4.1 Entra en Procesando pago EFECTIVO obvio no le mande primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && PaymentID == 'cod'
      && WA3 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Enviaremos tu pedido en la primera ruta por la mañana.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #' + Pedido + '\n' +
        'Con pago en Efectivo' + '\n' +
        'Con un total de *' + Total.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        '*Ya lo estamos preparando, en cuanto esté por salir a ruta, te enviaremos una notificación con la hora aproximada de entrega, así como el contacto del mensajero.*' + '\n' +
        'Nuestros mensajeros siempre llevan cambio.' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        'Para que el mensajero te pueda entregar es necesario que seas mayor de edad y le muestres tu identificación oficial.' + '\n' +
        '*Por ser dispositivos desechables, en caso de alguna falla para poder hacer válida la garantía, es necesario que pruebes el producto frente al mensajero. NO HAY EXCEPCIONES*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.5.1 Entra en Procesando pago Tarjeta Fisica obvio no le mande primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && PaymentID == 'cheque'
      && WA3 == '') {

      var additionalText = '';

      if (currentHour >= 22 || currentHour < 10) {
        additionalText =
          '❗️ Nuestro horario de servicio es de 10am a 10pm 🕙 todos los dias y actualmente estamos descansando.' + '\n' +
          'Enviaremos tu pedido en la primera ruta por la mañana.' + '\n\n' +
          '----------------------------' + ' \n\n'
      }

      var message = additionalText +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #' + Pedido + '\n' +
        'Con un total de *' + Total.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) + '*\n\n' +

        '*Ya lo estamos preparando, en cuanto esté por salir a ruta, te enviaremos una notificación con la hora aproximada de entrega, así como el contacto del mensajero.*' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, en caso de alguna falla para poder hacer válida la garantía, es necesario que pruebes el producto frente al mensajero. NO HAY EXCEPCIONES*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }


    //// Foraneos /////

    // 3.4 Transferencia Pasa de on-hold a Procesando 
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Transferencia")
      && WA1 != ''
      && WA3 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        '*Recibimos el pago de tu pedido* #' + Pedido + ', Muchas gracias!' + '\n\n' +

        'Ya lo estamos preparando, en cuanto generemos tu guia, te enviaremos el link por este medio y a tu correo para rastrearla.' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, para poder ofrecerte una garantía por servicio en caso de que falle alguno, es necesario grabar un video desde que abres el sobre y probar todos en el video, sin cortes ni ediciones en cuadro*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.4.1 Foraneo Clip pasa de on-hold a Procesando y NO le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && ShippingID.includes("enviaya")
      && (MetodoPago.includes("Clip") || PaymentID.includes('wc_clip'))
      && WA1 == ''
      && WA3 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        '*Recibimos el pago de tu pedido* #' + Pedido + ', Muchas gracias!' + '\n\n' +

        'Ya lo estamos preparando, en cuanto generemos tu guia, te enviaremos el link por este medio y a tu correo para rastrearla.' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, para poder ofrecerte una garantía por servicio en caso de que falle alguno, es necesario grabar un video desde que abres el sobre y probar todos en el video, sin cortes ni ediciones en cuadro*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnvios.getRange(i + 3, 35).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.4.2 Foraneo Clip pasa de on-hold a Procesando y ya le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && ShippingID.includes("enviaya")
      && (MetodoPago.includes("Clip") || PaymentID.includes('wc_clip'))
      && WA1 != ''
      && WA3 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + ' \n' +
        '*Recibimos el pago de tu pedido* #' + Pedido + ', Muchas gracias !' + '\n\n' +

        'Ya lo estamos preparando, en cuanto generemos tu guia, te enviaremos el link por este medio y a tu correo para rastrearla.' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, para poder ofrecerte una garantía por servicio en caso de que falle alguno, es necesario grabar un video desde que abres el sobre y probar todos en el video, sin cortes ni ediciones en cuadro*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.5.1 Foraneo Zelle pasa de on-hold a Procesando y NO le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Zelle")
      && WA1 == ''
      && WA3 == '') {

      var message =
        'Hello ' + Cliente.slice(0, Cliente.indexOf(" ")) + '' + '\n' +
        'This is TouchVapes.com' + '\n\n' +

        '*We received payment for your order* #' + Pedido + ' \n' +
        'Thank you very much!' + '\n\n' +

        'We are already preparing it, and as soon as we generate your shipping label, we will send you the link through this channel and to your email to track it.' + '\n\n' +

        'Thank you for your purchase!' + '\n\n' +

        '*REMINDER:*' + '\n' +
        '*As these are disposable devices, in order to offer you a warranty for service in case any of them fail, it is necessary to record a video from the moment you open the package and test all the devices in the video, without cuts or edits in the frame.*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnvios.getRange(i + 3, 35).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 3.5.2 Foraneo Zelle pasa de on-hold a Procesando y ya le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Zelle")
      && WA1 != ''
      && WA3 == '') {

      var message =
        'Hello ' + Cliente.slice(0, Cliente.indexOf(" ")) + '' + '\n' +

        '*We received payment for your order* #' + Pedido + ' \n' +
        'Thank you very much!' + '\n\n' +

        'We are already preparing it, and as soon as we generate your shipping label, we will send you the link through this channel and to your email to track it.' + '\n\n' +

        'Thank you for your purchase!' + '\n\n' +

        '*REMINDER:*' + '\n' +
        '*As these are disposable devices, in order to offer you a warranty for service in case any of them fail, it is necessary to record a video from the moment you open the package and test all the devices in the video, without cuts or edits in the frame.*'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de procesing:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }


    // CREO YA NO HACE FALTA 3.5.1 pasa de pending a Procesando y NO le mande un primer mensaje osea no mando el de pending
    /*
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && TipoEnvio.includes("MXN")
      && MetodoPago.includes("Mercado")
      && WA2 == ''
      && WA3 == ''    ) {

      var message =
        '3.5.1 FORANEO - pasa de pasa de pending a Procesando y NO le mande un primer mensaje ' + '\n\n' +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        'Te escribo de TouchVapes.com' + '\n\n' +

        'Recibimos tu pedido #' + Pedido + '\n\n' +

        '*Ya lo estamos preparando, en cuanto generemos tu guia, te enviaremos el link por este medio y a tu correo para rastrearla.*' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, para poder ofrecerte una garantía por servicio en caso de que falle alguno, es necesario grabar un video desde que abres el sobre y probar todos en el video, sin cortes ni ediciones en cuadro*'

      //send notif to whatsapp
      const status_wa = SendWATV({
        //to: Tel.toString(),
        to: salo,
        msg: message
      })

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(new Date()).setNumberFormat('hh:mm:ss').setHorizontalAlignment("center").setVerticalAlignment("middle");
      range.setNote(JSON.stringify(status_wa))
      console.log({ status_wa, Tel })

      //sheetEnvios.getRange(i + 3, 36).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");

    }

    // CREO YA NO HACE FALTA 3.5.2 pasa de pending a Procesando y ya le mande un primer mensaje
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'processing'
      && TipoEnvio.includes("MXN")
      && MetodoPago.includes("Mercado")
      && WA2 != 'N/A'
      && WA3 == ''    ) {

      var message =
        '3.5.2 FORANEO - pasa de pasa de pending a Procesando y ya le mande un primer mensaje ' + '\n\n' +
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n' +
        '*Recibimos el pago de tu pedido* #' + Pedido + ', Muchas gracias!' + '\n\n' +

        '*Ya lo estamos preparando, en cuanto generemos tu guia, te enviaremos el link por este medio y a tu correo para rastrearla.*' + '\n\n' +

        'Gracias por tu compra!' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        '*Por ser dispositivos desechables, para poder ofrecerte una garantía por servicio en caso de que falle alguno, es necesario grabar un video desde que abres el sobre y probar todos en el video, sin cortes ni ediciones en cuadro.*'

      //send notif to whatsapp
      const status_wa = SendWATV({
        //to: Tel.toString(),
        to: salo,
        msg: message
      })

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 42)
      range.setValue(new Date()).setNumberFormat('hh:mm:ss').setHorizontalAlignment("center").setVerticalAlignment("middle");
      range.setNote(JSON.stringify(status_wa))
      console.log({ status_wa, Tel })
    }
*/
  });

}


// Completado
function WATV4() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    //// Locales /////

    // 4.1 Completed despues de pasar por processing
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && !MetodoPago.includes("Zelle")
      && Mensajero != ''
      && HrEntrega != ''
      && WA3 != ''
      && WA4 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Tu pedido #' + Pedido + ' ya está en camino*' + '\n' +
        'La hora estimada de entrega es entre ' + HrEntrega + ' hrs' + '\n\n' +

        'Tu mensajero es *' + Mensajero + '*, si necesitas algo escríbele aquí ' + GetMensajero(Mensajero)[2] + '\n' +
        //'*Por favor no le llames ya que retrasa el proceso*, el se comunicara contigo cuando tu pedido sea el siguiente en entregarse.' + '\n\n' +
        'El se comunicara contigo cuando tu pedido sea el siguiente en entregarse.' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        'Para evitar retrasos en las entregas, los mensajeros no pueden ingresar ni subir a los domicilios, te pedimos que los veas en Planta Baja para que pueda seguir su ruta lo antes posible. El mensajero no puede esperar mas de 5 minutos en tu domicilio, si deseas reagendar la entrega, avisanos aqui antes de que llegue el mensajero, asi evitaras que se te cobre doble envio.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.1.2 Completed despues de pasar por processing para ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && MetodoPago.includes("Zelle")
      && Mensajero != ''
      && HrEntrega != ''
      && WA3 != ''
      && WA4 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Your Order #' + Pedido + ' it\'s on its way*' + '\n' +
        'The estimated delivery time is between ' + HrEntrega + ' hrs' + '\n\n' +

        'Your delivery person is *' + Mensajero + '*, if you need anything please text him here ' + GetMensajero(Mensajero)[2] + '\n' +
        'He will contact you when your order is the next one to be delivered.' + '\n\n' +

        '*REMINDER:*' + '\n' +
        'To avoid delays in deliveries, couriers are not allowed to enter or go up to residences. We ask that you meet them on the ground floor so that they can continue their route as soon as possible. The courier cannot wait for more than 5 minutes at your residence. If you wish to reschedule the delivery, please let us know here before the courier arrives to avoid being charged for double shipping.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.2 Completed pero no paso por processing
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && !MetodoPago.includes("Zelle")
      && Mensajero != ''
      && HrEntrega != ''
      && WA3 == ''
      && WA4 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Recibimos el pago de tu pedido* # ' + Pedido + ', Muchas gracias!, *ya esta en camino*' + '\n' +
        'La hora estimada de entrega es entre ' + HrEntrega + ' hrs' + '\n\n' +

        'Tu mensajero es *' + Mensajero + '*, si necesitas algo escríbele aquí ' + GetMensajero(Mensajero)[2] + '\n' +
        //'*Por favor no le llames ya que retrasa el proceso*, el se comunicara contigo cuando tu pedido sea el siguiente en entregarse.' + '\n\n' +
        'El se comunicara contigo cuando tu pedido sea el siguiente en entregarse.' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        'Para evitar retrasos en las entregas, los mensajeros no pueden ingresar ni subir a los domicilios, te pedimos que los veas en Planta Baja para que pueda seguir su ruta lo antes posible. El mensajero no puede esperar mas de 5 minutos en tu domicilio, si deseas reagendar la entrega, avisanos aqui antes de que llegue el mensajero, asi evitaras que se te cobre doble envio.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      // Pone N/A a processing
      //sheetEnvios.getRange(i + 3, 37).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");

    }

    // 4.2 Completed pero no paso por processing para ZELLE
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      && MetodoPago.includes("Zelle")
      && Mensajero != ''
      && HrEntrega != ''
      && WA3 == ''
      && WA4 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*We have recieved the payment for your order # ' + Pedido + ', Thank you very much!, it\'s on its way*' + '\n' +
        'The estimated delivery time is between ' + HrEntrega + ' hrs' + '\n\n' +

        'Your delivery person is *' + Mensajero + '*, if you need anything please text him here ' + GetMensajero(Mensajero)[2] + '\n' +
        'He will contact you when your order is the next one to be delivered.' + '\n\n' +

        '*REMINDER:*' + '\n' +
        'To avoid delays in deliveries, couriers are not allowed to enter or go up to residences. We ask that you meet them on the ground floor so that they can continue their route as soon as possible. The courier cannot wait for more than 5 minutes at your residence. If you wish to reschedule the delivery, please let us know here before the courier arrives to avoid being charged for double shipping.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      // Pone N/A a processing
      //sheetEnvios.getRange(i + 3, 37).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");

    }


    ///// Foraneos /////

    // 4.3 Completed despues de pasar por processing NO PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && !MetodoPago.includes("Zelle")
      && Seguro == 0
      && WA3 != ''
      && WA4 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Ya generamos tu guia del pedido #' + Pedido + '*, te paso el link de rastreo.' + '\n' +
        LinkGuia + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        'Nuestra resposabilidad termina al entregar tu pedido en paqueteria, ya que no pagaste el seguro contra decomiso, por lo que no nos podemos hacer responsables por perdidas o decomisios por parte de la autoridad.' + '\n\n' +

        'Muchas gracias, y seguimos a la orden!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.3.1 Completed despues de pasar por processing SI PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && !MetodoPago.includes("Zelle")
      && Seguro > 0
      && WA3 != ''
      && WA4 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Ya generamos tu guia del pedido #' + Pedido + '*, te paso el link de rastreo.' + '\n' +
        LinkGuia + '\n\n' +

        'Muchas gracias, y seguimos a la orden!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.3.2 Completed despues de pasar por processing para ZELLE NO PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Zelle")
      && Seguro == 0
      && WA3 != ''
      && WA4 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*We have generated the shipping label for the order #' + Pedido + '*, here is the link to track it.' + '\n' +
        LinkGuia + '\n\n' +

        '*REMINDER:*' + '\n' +
        'Our liability ends upon delivery of your order to the courier, since you declined to obtain insurance against confiscation. Therefore, we cannot assume responsibility for any losses or confiscations by the authorities.' + '\n\n' +

        'Thank you very much and we remain at your service!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.3.3 Completed despues de pasar por processing para ZELLE SI PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Zelle")
      && Seguro > 0
      && WA3 != ''
      && WA4 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*We have generated the shipping label for the order #' + Pedido + '*, here is the link to track it.' + '\n' +
        LinkGuia + '\n\n' +

        'Thank you very much and we remain at your service!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.4.1 Completed pero no paso por processing NO PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && !MetodoPago.includes("Zelle")
      && Seguro == 0
      && WA3 == ''
      && WA4 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Ya generamos tu guia del pedido #' + Pedido + '*, te paso el link de rastreo.' + '\n' +
        LinkGuia + '\n\n' +
        //'*Ya generamos tu guia del pedido #' + Pedido + '*, debes de tener el link para rastrearlo en tu correo, si no lo encuentras, por favor avisanos.' + '\n\n' +

        '*TE RECORDAMOS:*' + '\n' +
        'Nuestra resposabilidad termina al entregar tu pedido en paqueteria, ya que no pagaste el seguro contra decomiso, por lo que no nos podemos hacer responsables por perdidas o decomisios por parte de la autoridad.' + '\n\n' +

        'Muchas gracias, y seguimos a la orden!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnvios.getRange(i + 3, 38).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.4.2 Completed pero no paso por processing SI PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && !MetodoPago.includes("Zelle")
      && Seguro > 0
      && WA3 == ''
      && WA4 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Ya generamos tu guia del pedido #' + Pedido + '*, te paso el link de rastreo.' + '\n' +
        LinkGuia + '\n\n' +

        'Muchas gracias, y seguimos a la orden!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnvios.getRange(i + 3, 38).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.4.3 Completed pero no paso por processing para ZELLE NO PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Zelle")
      && Seguro == 0
      && WA3 == ''
      && WA4 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*We have generated the shipping label for the order #' + Pedido + '*, here is the link to track it.' + '\n' +
        LinkGuia + '\n\n' +

        '*REMINDER:*' + '\n' +
        'Our liability ends upon delivery of your order to the courier, since you declined to obtain insurance against confiscation. Therefore, we cannot assume responsibility for any losses or confiscations by the authorities.' + '\n\n' +

        'Thank you very much and we remain at your service!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnvios.getRange(i + 3, 38).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // 4.4.4 Completed pero no paso por processing para ZELLE SI PAGO SEGURO
    else if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && ShippingID.includes("enviaya")
      && MetodoPago.includes("Zelle")
      && Seguro > 0
      && WA3 == ''
      && WA4 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*We have generated the shipping label for the order #' + Pedido + '*, here is the link to track it.' + '\n' +
        LinkGuia + '\n\n' +

        'Thank you very much and we remain at your service!'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de completed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 43)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");

      //sheetEnvios.getRange(i + 3, 38).setValue('N/A').setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

  });

}

// Cancelado
function WATV5() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    // NO ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'cancelled'
      && !MetodoPago.includes("Zelle")
      && WA5 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        'Tu *pedido #' + Pedido + ' ha sido cancelado*' + '\n' +
        'Si sientes que es un error, avisanos.'

      /* 
      //Start Test
      var url = "http://3.144.48.182:8083/message/sendText/codechat";

      // Verificamos si el número de teléfono comienza con '521'
      if (WA.startsWith('521')) {
        // Si es así, entonces reemplazamos '521' con '52'
        WASinElUno = WA.replace('521', '52');
      }

      var data = {
        "number": WASinElUno,
        "options": {
          "delay": 1200
        },
        "textMessage": {
          "text": "*❗️❗️❗️ Nuevo*" + '\n\n\n' + message
        }
      };
      var options = {
        method: 'post',
        headers: {
          "Content-Type": "application/json",
          "apikey": "zYzP7ocstxh3SJ23D4FZTCu4ehnM8v4hu"
        },
        payload: JSON.stringify(data),
        muteHttpExceptions: true
      };
      // Intentamos realizar la petición
      try {
        var response = UrlFetchApp.fetch(url, options);
        Logger.log(response.getContentText());
      } catch (e) {
        // Si hay una excepción (como 'Address unavailable'), la registramos y continuamos con el script
        Logger.log('Exception: ' + e.toString());
      }
      // End Test
      */

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de cancelled:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 44)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    //  ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'cancelled'
      && MetodoPago.includes("Zelle")
      && WA5 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        'Your *Order #' + Pedido + ' has been canceled*' + '\n' +
        'If you feel that it\'s a mistake, let us know.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de cancelled:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 44)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

  });

}

// Refunded
function WATV6() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    // NO ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && (MetodoPago.includes("Clip") || PaymentID.includes("wc_clip"))
      && !MetodoPago.includes("Zelle")
      && Status == 'refunded'
      && WA6 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Tu pedido ' + Pedido + ' ha sido reembolsado*' + '\n' +
        'Para cualquier duda, aqui estamos para apoyarte.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de refunded:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 45)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && (MetodoPago.includes("Clip") || PaymentID.includes("wc_clip"))
      && MetodoPago.includes("Zelle")
      && Status == 'refunded'
      && WA6 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Your Order ' + Pedido + ' has been refunded*' + '\n' +
        'If you have any questions, please let us know.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de refunded:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 45)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

  });

}

// Failed
function WATV7() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    // NO ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && (MetodoPago.includes("Clip") || PaymentID.includes("wc_clip"))
      && !MetodoPago.includes("Zelle")
      && Status == 'failed'
      && WA7 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        'Una disculpa, *el pago de tu pedido ' + Pedido + ' ha sido rechazado.*' + '\n' +
        'Para cualquier duda, aqui estamos para apoyarte.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de failed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 46)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && (MetodoPago.includes("Clip") || PaymentID.includes("wc_clip"))
      && MetodoPago.includes("Zelle")
      && Status == 'failed'
      && WA7 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        'We are sorry, *the payment of your order ' + Pedido + ' has been declined.*' + '\n' +
        'If you have any questions, please let us know.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de failed:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 46)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

  });

}


// Hr Updated
function WATV8() {

  data.forEach(function (row, i) {

    var Empresa = row[0]; var Fecha = row[1]; var Hora = row[2]; var Pedido = row[3]; var Cliente = row[4]; var DireccionCompleta = row[5]; var CalleYNumero = row[6]; var Notas = row[7]; var Colonia = row[8]; var Ciudad = row[9]; var Estado = row[10]; var CP = row[11]; var CountryCode = row[12]; var Tel = row[13]; var WA = row[14]; var ShippingID = row[15]; var TipoEnvio = row[16]; var PaymentID = row[17]; var MetodoPago = row[18]; var Subtotal = row[19]; var CostoEnvio = row[20]; var Seguro = row[21]; var Comision = row[22]; var Total = row[23]; var EfeReal = row[24]; var Mensajero = row[25]; var Status = row[26]; var HrEntrega = row[27]; var HrUpdate = row[28]; var Terminado = row[29]; var LinkGuia = row[30]; var LinkClip = row[31]; var TotalDLS = row[32]; var Vacia1 = row[33]; var Vacia2 = row[34]; var Vacia3 = row[35]; var Vacia4 = row[36]; var Vacia5 = row[37]; var Cobro = row[38]; var WA1 = row[39]; var WA2 = row[40]; var WA3 = row[41]; var WA4 = row[42]; var WA5 = row[43]; var WA6 = row[44]; var WA7 = row[45]; var WA8 = row[46]; var WA9 = row[47]; var WA10 = row[48];

    // Nueva validación para números de teléfono que comienzan con '521'
    if (WA.startsWith('521')) {
      WA = '52' + WA.slice(3);
    }

    // Local 
    // NO ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      & !MetodoPago.includes("Zelle")
      && Mensajero != ''
      && HrEntrega != ''
      && HrUpdate != ''
      && WA8 == '') {

      var message =
        'Hola *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*Se actualizó el tiempo estimado de entrega de tu pedido #' + Pedido + '*\n' +
        'Calculamos que tu pedido se va entregar entre *' + HrUpdate + ' hrs' + '*\n\n' +

        'Tu mensajero es *' + Mensajero + '*, si necesitas algo escríbele aquí ' + GetMensajero(Mensajero)[2] + '\n' +
        //'*Por favor no le llames ya que retrasa el proceso*, el se comunicara contigo cuando tu pedido sea el siguiente en entregarse.' + '\n\n' +
        'El se comunicara contigo cuando tu pedido sea el siguiente en entregarse.'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de Hr Update:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 47)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    // ZELLE
    if (
      Empresa == 'TV'
      && Fecha != ''
      && Status == 'completed'
      && (ShippingID == 'distance_rate' || ShippingID == 'free_shipping')
      & MetodoPago.includes("Zelle")
      && Mensajero != ''
      && HrEntrega != ''
      && HrUpdate != ''
      && WA8 == '') {

      var message =
        'Hello *' + Cliente.slice(0, Cliente.indexOf(" ")) + '*' + '\n\n' +

        '*The estimated delivery time of your order has been updated #' + Pedido + '*\n' +
        'We estimate that your order will be delivered between *' + HrUpdate + ' hrs' + '*\n\n' +

        'Your delivery person is *' + Mensajero + '*, If you need anything, text him here ' + GetMensajero(Mensajero)[2] + '\n' +
        'He will contact you when your order is the next to be delivered'

      // API DE SALO
      //send notif to whatsapp
      var status_wa = SendWAApiSaloTV({
        to: WA.toString(),
        msg: message
      })

      var msg2 = 'Error al enviar mensaje de Hr Update:' + '\n' +
        'Pedido: ' + Pedido + '\n' +
        'Cliente: ' + Cliente + '\n' +
        'Telefono: ' + Tel

      if (status_wa.success) {
        status = 'OK'
      }
      else {
        status = 'Error'
        var wa_failed = SendWAApiSaloTV({
          to: touch,
          msg: msg2
        })
      }

      // Pone confirmacion de enviado  
      const range = sheetEnvios.getRange(i + 3, 47)
      range.setValue(status + ' - ' + Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm')).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

  });

}


/*

'Tu mensajero es : ' + Mensajero + '\n' +

 'El contacto del mensajero es : ' + GetMensajero(Mensajero)[2] + '\n' +
 // Utilities.formatDate(Fecha, "GMT", "dd-MMM-yy") + '\n' +

         'El detalle de tu pedido lo puedes revisar aqui: ' + '\n' +
        'https://touchvapes.com/my-account/view-order/' + Orden + '\n' +

*/

/*
var Empresa = row[0];
var Fecha = row[1];
var Hora = row[2];
var Pedido = row[3];
var Cliente = row[4];
var DireccionCompleta = row[5];
var CalleYNumero = row[6];
var Notas = row[7];
var Colonia = row[8];
var Ciudad = row[9];
var Estado = row[10];
var CP = row[11];
var CountryCode = row[12];
var Tel = row[13];
var WA = row[14];
var ShippingID = row[15];
var TipoEnvio = row[16];
var PaymentID = row[17];
var MetodoPago = row[18];
var Subtotal = row[19];
var CostoEnvio = row[20];
var Seguro = row[21];
var Comision = row[22];
var Total = row[23];
var EfeReal = row[24];
var Mensajero = row[25];
var Status = row[26];
var HrEntrega = row[27];
var HrUpdate = row[28];
var Terminado = row[29];
var LinkGuia = row[30];
var LinkClip = row[31];
var TotalDLS = row[32];
var Vacia1 = row[33];
var Vacia2 = row[34];
var Vacia3 = row[35];
var Vacia4 = row[36];
var Vacia5 = row[37];
var Cobro = row[38];
var WA1 = row[39];
var WA2 = row[40];
var WA3 = row[41];
var WA4 = row[42];
var WA5 = row[43];
var WA6 = row[44];
var WA7 = row[45];
var WA8 = row[46];
var WA9 = row[47];
var WA10 = row[48];
*/


