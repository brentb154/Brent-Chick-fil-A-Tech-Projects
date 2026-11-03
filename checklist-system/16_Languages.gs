/**
 * ============================================================
 * CHECKLIST SYSTEM - Form Languages
 * ============================================================
 * A checklist's form can be in Spanish (Checklists "Language" =
 * Spanish). Everything a responder sees on it is translated: the
 * questions, the answer choices, the photo tag, the photo lines,
 * the date and the closed message. Each task uses Items "Spanish
 * task", or the English task when that's blank.
 *
 * Submissions are logged in English: answers are mapped back to
 * Complete / Could not complete and Item Results gets the English
 * task, so the morning email and scorecard read as always.
 */

var FORM_TEXT = {
  en: {
    leader: Q_LEADER,
    position: Q_POSITION,
    code: Q_STATION_CODE,
    notes: Q_NOTES,
    done: ANSWER_DONE,
    notDone: ANSWER_NOT_DONE,
    photo: PHOTO_SUFFIX,
    uploadLine: 'Pictures: after you submit, upload them here (no sign-in). Add them all, then tap Upload: ',
    emailFirst: 'Email all pictures to ',
    emailOr: 'Or email them to ',
    subject: ' with the subject ',
    thanks: 'Thanks, your checklist is in.',
    nowUpload: 'Now upload your pictures',
    samePosition: ' for the same position',
    closed: 'No checklist today.'
  },
  es: {
    leader: 'Nombre del líder que completa esta lista',
    position: '¿Qué posición estás completando?',
    code: 'Código de estación (se llena solo al escanear el código QR de la estación)',
    notes: '¿Algo que no se completó o que necesita atención?',
    done: 'Completado',
    notDone: 'No se pudo completar',
    photo: ' **TOMA FOTOS**',
    uploadLine: 'Fotos: después de enviar, súbelas aquí (sin iniciar sesión). Agrégalas todas y luego toca Upload: ',
    emailFirst: 'Manda todas las fotos por correo a ',
    emailOr: 'O mándalas por correo a ',
    subject: ' con el asunto ',
    thanks: 'Gracias, tu lista ya se envió.',
    nowUpload: 'Ahora sube tus fotos',
    samePosition: ' de la misma posición',
    closed: 'Hoy no hay lista.'
  }
};

var DAY_ES = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
var MONTH_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
  'septiembre', 'octubre', 'noviembre', 'diciembre'];

// Checklists "Language": Spanish (or Español / es) -> 'es'; anything else -> 'en'
function languageCode_(text) {
  return /^(spanish|espa[nñ]ol|es)$/i.test(String(text || '').trim()) ? 'es' : 'en';
}

function formText_(checklist) {
  return FORM_TEXT[checklist.language] || FORM_TEXT.en;
}

// "Saturday, October 10, 2026" / "Sábado, 10 de octubre de 2026"
function dateLabel_(key, lang) {
  if (lang !== 'es') return longLabel_(key);
  var p = key.split('-');
  var day = DAY_ES[dayIndex_(key)];
  return day.charAt(0).toUpperCase() + day.slice(1) + ', ' + Number(p[2]) + ' de ' + MONTH_ES[Number(p[1]) - 1] + ' de ' + p[0];
}

// Which fixed question a form title is, in any language: 'leader', 'position', 'code', 'notes' or ''
function questionRole_(title) {
  var role = '';
  Object.keys(FORM_TEXT).forEach(function (lang) {
    ['leader', 'position', 'code', 'notes'].forEach(function (r) {
      if (!role && FORM_TEXT[lang][r] === title) role = r;
    });
  });
  return role;
}

// An answer as Item Results stores it: "Complete" / "Could not complete" in English; typed answers as typed
function englishAnswer_(answer) {
  var out = answer;
  Object.keys(FORM_TEXT).forEach(function (lang) {
    if (answer === FORM_TEXT[lang].done) out = ANSWER_DONE;
    else if (answer === FORM_TEXT[lang].notDone) out = ANSWER_NOT_DONE;
  });
  return out;
}

// The English task as an English form would have shown it that day (same random pick, same photo tag)
function englishTask_(item, dateKey) {
  return pickOne_(item.task, dateKey) + (item.photo ? PHOTO_SUFFIX : '');
}
