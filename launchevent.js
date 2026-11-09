/*
 * Aviso de desconexión laboral - complemento de Outlook (Smart Alerts / OnMessageSend)
 * Agrega el aviso al inicio del correo cuando se envía fuera del horario laboral:
 *   - Antes de las 7:45 a.m. o después de las 5:15 p.m. (hora del equipo)
 *   - Sábados y domingos (todo el día)
 *
 * Versión 1.2 (diagnóstico): escribe cada paso en la consola con el prefijo
 * [AvisoDesconexion] y nunca deja el envío colgado (try/catch + tiempo límite).
 */

// ====== CONFIGURACIÓN ======
var HORA_INICIO = { h: 7, m: 45 };   // 7:45 a.m.
var HORA_FIN = { h: 17, m: 15 };     // 5:15 p.m.

var MARCADOR = "tu jornada laboral y tu derecho";

var AVISO_HTML =
  "<p style='color:#666666;font-size:10pt;font-family:Calibri, Arial, sans-serif;'>" +
  "<b>La empresa respeta tu jornada laboral y tu derecho a la desconexión.</b><br>" +
  "Cualquier envío de correos fuera de este horario es voluntario y bajo responsabilidad del colaborador." +
  "</p><br>";

var TIEMPO_LIMITE_MS = 8000; // si algo no responde en 8 s, se deja enviar igual
// ===========================

function log(msg) {
  try { console.log("[AvisoDesconexion] " + msg); } catch (e) {}
}

log("script cargado v1.2 | horario " + HORA_INICIO.h + ":" + HORA_INICIO.m + " a " + HORA_FIN.h + ":" + HORA_FIN.m);

// Obligatorio: indica a Outlook que el complemento terminó de inicializarse.
// Sin esto, la página carga pero Outlook nunca llama al handler.
Office.onReady(function (info) {
  log("onReady OK | host: " + (info && info.host) + " | plataforma: " + (info && info.platform));
});

function esFueraDeHorario(fecha) {
  var dia = fecha.getDay(); // 0 = domingo, 6 = sábado
  if (dia === 0 || dia === 6) {
    return true;
  }
  var minutos = fecha.getHours() * 60 + fecha.getMinutes();
  var inicio = HORA_INICIO.h * 60 + HORA_INICIO.m;
  var fin = HORA_FIN.h * 60 + HORA_FIN.m;
  return minutos < inicio || minutos > fin;
}

function onMessageSendHandler(event) {
  var terminado = false;

  function terminar(motivo) {
    if (terminado) { return; }
    terminado = true;
    log("fin: " + motivo);
    event.completed({ allowEvent: true });
  }

  // Red de seguridad: nunca dejar el correo colgado.
  setTimeout(function () { terminar("tiempo limite alcanzado"); }, TIEMPO_LIMITE_MS);

  try {
    var ahora = new Date();
    var fuera = esFueraDeHorario(ahora);
    log("handler llamado " + ahora.toString() + " | fuera de horario: " + fuera);

    if (!fuera) {
      terminar("dentro del horario");
      return;
    }

    var item = Office.context.mailbox.item;

    item.body.getAsync(Office.CoercionType.Html, function (resultado) {
      try {
        if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
          terminar("getAsync fallo: " + JSON.stringify(resultado.error));
          return;
        }

        var cuerpo = resultado.value || "";
        if (cuerpo.toLowerCase().indexOf(MARCADOR) !== -1) {
          terminar("ya tenia el aviso");
          return;
        }

        item.body.prependAsync(
          AVISO_HTML,
          { coercionType: Office.CoercionType.Html },
          function (r) {
            terminar("prependAsync " + r.status + (r.error ? " " + JSON.stringify(r.error) : ""));
          }
        );
      } catch (e2) {
        terminar("error en callback: " + e2);
      }
    });
  } catch (e) {
    terminar("error: " + e);
  }
}

// Conecta el nombre declarado en el manifiesto con la función.
try {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  log("associate OK");
} catch (e) {
  log("associate FALLO: " + e);
}
