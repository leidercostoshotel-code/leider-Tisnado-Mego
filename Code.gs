// =====================================================
// LTM Soluciones Digitales — Google Apps Script
//
// INSTRUCCIONES DE CONFIGURACIÓN:
// 1. Ve a sheets.google.com y crea una hoja nueva
// 2. En el menú: Extensiones > Apps Script
// 3. Borra el código que aparece y pega TODO este archivo
// 4. Haz clic en "Implementar" > "Nueva implementación"
// 5. Tipo: "Aplicación web"
//    - Ejecutar como: "Yo (tu cuenta)"
//    - Acceso: "Cualquier persona"
// 6. Haz clic en "Implementar" y copia la URL
// 7. En index.html busca APPS_SCRIPT_URL y pega la URL
// =====================================================

const EMAIL_DESTINO = 'leidercostoshotel@gmail.com';

function doPost(e) {
  try {
    const p        = e.parameter;
    const nombre   = p.nombre   || '';
    const telefono = p.telefono || '';
    const email    = p.email    || '';
    const servicio = p.servicio || '';
    const mensaje  = p.mensaje  || '';
    const fecha    = new Date();

    guardarEnSheet(fecha, nombre, telefono, email, servicio, mensaje);
    enviarCorreo(fecha, nombre, telefono, email, servicio, mensaje);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function guardarEnSheet(fecha, nombre, telefono, email, servicio, mensaje) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];

  if (sheet.getLastRow() === 0) {
    const cabeceras = ['Fecha', 'Nombre', 'Teléfono', 'Email', 'Servicio', 'Mensaje'];
    sheet.appendRow(cabeceras);
    const rango = sheet.getRange(1, 1, 1, cabeceras.length);
    rango.setFontWeight('bold')
         .setBackground('#061a44')
         .setFontColor('#ffffff')
         .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 210);
    sheet.setColumnWidth(5, 190);
    sheet.setColumnWidth(6, 320);
  }

  sheet.appendRow([fecha, nombre, telefono, email, servicio, mensaje]);
}

function enviarCorreo(fecha, nombre, telefono, email, servicio, mensaje) {
  const tz       = Session.getScriptTimeZone();
  const fechaStr = Utilities.formatDate(fecha, tz, 'dd/MM/yyyy HH:mm');
  const asunto   = `🚀 Nueva solicitud: ${servicio} | LTM Soluciones Digitales`;

  const html = `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:'Segoe UI',Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8;padding:40px 16px;">
  <tr><td align="center">
  <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">

    <!-- ENCABEZADO -->
    <tr><td style="background:linear-gradient(135deg,#061a44 0%,#0b3b91 50%,#00c2ff 100%);
                   border-radius:18px 18px 0 0;padding:38px 40px;text-align:center;">
      <div style="font-size:12px;font-weight:700;color:#00c2ff;letter-spacing:4px;
                  text-transform:uppercase;margin-bottom:10px;">
        LTM Soluciones Digitales
      </div>
      <div style="font-size:32px;font-weight:900;color:#ffffff;margin-bottom:6px;">
        🚀 Nueva Solicitud
      </div>
      <div style="font-size:13px;color:rgba(255,255,255,.65);">
        Recibido el ${fechaStr}
      </div>
    </td></tr>

    <!-- CUERPO -->
    <tr><td style="background:#ffffff;padding:40px 40px 32px;">

      <p style="margin:0 0 26px;font-size:16px;color:#334155;line-height:1.6;">
        Hola <strong style="color:#061a44;">Leider</strong>, tienes una nueva solicitud de cotización.
        Aquí están todos los detalles:
      </p>

      <!-- BADGE SERVICIO -->
      <div style="background:linear-gradient(135deg,#eaf7ff,#dbeafe);
                  border:1px solid #bfdbfe;border-radius:14px;
                  padding:18px 24px;margin-bottom:30px;text-align:center;">
        <div style="font-size:11px;font-weight:700;color:#0069aa;
                    text-transform:uppercase;letter-spacing:2px;margin-bottom:6px;">
          Servicio solicitado
        </div>
        <div style="font-size:22px;font-weight:900;color:#061a44;">
          ${servicio || '—'}
        </div>
      </div>

      <!-- TABLA DE DATOS -->
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr><td style="padding:14px 0;border-bottom:1px solid #f1f5f9;">
          <div style="font-size:11px;font-weight:700;color:#94a3b8;
                      text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">
            👤 Nombre
          </div>
          <div style="font-size:17px;font-weight:700;color:#1e293b;">${nombre || '—'}</div>
        </td></tr>
        <tr><td style="padding:14px 0;border-bottom:1px solid #f1f5f9;">
          <div style="font-size:11px;font-weight:700;color:#94a3b8;
                      text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">
            📱 Teléfono / WhatsApp
          </div>
          <div style="font-size:17px;font-weight:700;color:#1e293b;">${telefono || '—'}</div>
        </td></tr>
        <tr><td style="padding:14px 0;border-bottom:1px solid #f1f5f9;">
          <div style="font-size:11px;font-weight:700;color:#94a3b8;
                      text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;">
            ✉️ Correo electrónico
          </div>
          <div style="font-size:17px;font-weight:700;color:#0d6efd;">${email || '—'}</div>
        </td></tr>
        <tr><td style="padding:16px 0 0;">
          <div style="font-size:11px;font-weight:700;color:#94a3b8;
                      text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">
            💬 Mensaje
          </div>
          <div style="font-size:15px;color:#334155;line-height:1.7;
                      background:#f8fafc;border-radius:12px;padding:16px 18px;
                      border-left:4px solid #00c2ff;">
            ${mensaje || '—'}
          </div>
        </td></tr>
      </table>

      <!-- BOTÓN RESPONDER -->
      <div style="margin-top:34px;text-align:center;">
        <a href="mailto:${email}?subject=Re: Tu solicitud en LTM Soluciones Digitales"
           style="display:inline-block;
                  background:linear-gradient(135deg,#061a44,#0b3b91,#00c2ff);
                  color:#ffffff;font-weight:800;font-size:15px;
                  padding:16px 36px;border-radius:999px;text-decoration:none;
                  letter-spacing:.02em;">
          Responder a ${nombre || 'el cliente'} →
        </a>
      </div>

    </td></tr>

    <!-- PIE -->
    <tr><td style="background:#f8fafc;border-radius:0 0 18px 18px;
                   padding:20px 40px;border-top:1px solid #e2e8f0;text-align:center;">
      <p style="margin:0;font-size:12px;color:#94a3b8;">
        LTM Soluciones Digitales · Lima, Perú · leidercostoshotel@gmail.com
      </p>
    </td></tr>

  </table>
  </td></tr>
</table>

</body>
</html>`;

  MailApp.sendEmail({
    to:       EMAIL_DESTINO,
    subject:  asunto,
    htmlBody: html
  });
}
